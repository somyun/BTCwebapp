"use strict";

const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function createHarness({ testBackend = false, overrideToken = true } = {}) {
  const cache = new Map();
  const properties = new Map();
  const sent = [];
  let uuidCounter = 0;
  const cacheApi = {
    get: (key) => cache.get(key) ?? null,
    put: (key, value) => cache.set(key, String(value)),
    remove: (key) => cache.delete(key)
  };
  const context = vm.createContext({
    console,
    Math,
    Date,
    JSON,
    String,
    Number,
    RegExp,
    CacheService: { getScriptCache: () => cacheApi },
    LockService: {
      getScriptLock: () => ({ waitLock() {}, releaseLock() {} })
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => properties.get(key) ?? null,
        setProperty: (key, value) => properties.set(key, value)
      })
    },
    MailApp: {
      getRemainingDailyQuota: () => 100,
      sendEmail: (message) => sent.push(message)
    },
    Utilities: {
      Charset: { UTF_8: "utf8" },
      DigestAlgorithm: { SHA_256: "sha256" },
      getUuid: () => `00000000-0000-4000-8000-${String(++uuidCounter).padStart(12, "0")}`,
      computeDigest: (_algorithm, value) => [...crypto.createHash("sha256").update(String(value)).digest()],
      base64EncodeWebSafe: (value) => Buffer.from(value).toString("base64url"),
      computeRsaSha256Signature: () => [1, 2, 3]
    },
    SERVICE_ACCOUNT_KEY: {
      client_email: "test@example.iam.gserviceaccount.com",
      private_key: "unused-in-test",
      project_id: testBackend ? "btcwebapp-test" : "btcwebapp-551bd"
    },
    TEST_FIREBASE_PROJECT_ID: "btcwebapp-test",
    assertTestEnvironment_() {}
  });
  const source = fs.readFileSync(path.join(
    __dirname,
    "..",
    testBackend ? "apps-script-test" : "apps-script",
    "auth_code.js"
  ), "utf8");
  vm.runInContext(source, context);
  if (overrideToken) context.createMapAuthCustomToken_ = (email) => `token:${email}`;
  return { context, cache, sent };
}

test("only the Humetro email domain can request a code", () => {
  const { context, sent } = createHarness();
  const result = context.requestMapAuthCode("person@gmail.com");
  assert.equal(result.success, false);
  assert.equal(result.code, "INVALID_DOMAIN");
  assert.equal(sent.length, 0);
});

test("a six-digit code is mailed and can be exchanged once", () => {
  const { context, sent } = createHarness();
  const requested = context.requestMapAuthCode("User@humetro.busan.kr");
  assert.equal(requested.success, true);
  assert.equal(sent.length, 1);
  const code = sent[0].body.match(/\b(\d{6})\b/)[1];

  const verified = context.verifyMapAuthCode("user@humetro.busan.kr", code);
  assert.equal(verified.success, true);
  assert.equal(verified.token, "token:user@humetro.busan.kr");

  const replay = context.verifyMapAuthCode("user@humetro.busan.kr", code);
  assert.equal(replay.success, false);
  assert.equal(replay.code, "CODE_EXPIRED");
});

test("resending within the cooldown is rejected", () => {
  const { context } = createHarness();
  assert.equal(context.requestMapAuthCode("user@humetro.busan.kr").success, true);
  const repeated = context.requestMapAuthCode("user@humetro.busan.kr");
  assert.equal(repeated.success, false);
  assert.equal(repeated.code, "RATE_LIMITED");
});

test("five incorrect attempts invalidate the code", () => {
  const { context, sent } = createHarness();
  assert.equal(context.requestMapAuthCode("user@humetro.busan.kr").success, true);
  const issuedCode = sent[0].body.match(/\b(\d{6})\b/)[1];
  const wrongCode = issuedCode === "000000" ? "000001" : "000000";
  for (let attempt = 1; attempt < 5; attempt += 1) {
    const result = context.verifyMapAuthCode("user@humetro.busan.kr", wrongCode);
    assert.equal(result.code, "INVALID_CODE");
  }
  const finalAttempt = context.verifyMapAuthCode("user@humetro.busan.kr", wrongCode);
  assert.equal(finalAttempt.code, "TOO_MANY_ATTEMPTS");
});

test("bwa_test auth keeps the same code flow and rejects an operational service account", () => {
  const { context, sent } = createHarness({ testBackend: true });
  assert.equal(context.requestMapAuthCode("user@humetro.busan.kr").success, true);
  const code = sent[0].body.match(/\b(\d{6})\b/)[1];
  assert.equal(context.verifyMapAuthCode("user@humetro.busan.kr", code).success, true);

  const guarded = createHarness({ testBackend: true, overrideToken: false }).context;
  guarded.SERVICE_ACCOUNT_KEY.project_id = "btcwebapp-551bd";
  assert.throws(() => guarded.createMapAuthCustomToken_("user@humetro.busan.kr"), {
    message: "TEST_FIREBASE_SERVICE_ACCOUNT_REQUIRED"
  });
});
