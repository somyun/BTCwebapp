"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const {
  createLegacyNotificationMigrator,
  legacyDeviceId,
  normalizeLegacyTokenRows
} = require("../lib/legacy-notification-migration");
const { hash } = require("../lib/notification-service");
const { FakeFirestore } = require("./helpers/fake-firestore");

const tokenA = `token-${"a".repeat(80)}`;
const tokenB = `token-${"b".repeat(80)}`;

test("normalizes active legacy rows without exposing duplicates or invalid tokens", () => {
  const result = normalizeLegacyTokenRows([
    [tokenA, "Chrome", "2026-08-01", "결혼", true],
    [tokenA, "Chrome", "2026-08-02", "부고", "TRUE"],
    [tokenB, "Safari", "2026-08-03", "", false],
    ["short", "Unknown", "", "", true]
  ]);
  assert.deepEqual(result.summary, { sourceRows: 4, activeUnique: 1, inactive: 1, invalid: 1, duplicate: 1 });
  assert.equal(result.devices[0].keywords, "부고");
  assert.equal(result.devices[0].deviceId, legacyDeviceId(tokenA));
});

test("dry run reports counts and confirmed migration creates one protected owner", async () => {
  const firestore = new FakeFirestore();
  const sheetsGateway = { async readRange() { return [[tokenA, "Chrome", "2026-08-02", "부고", true]]; } };
  const migrator = createLegacyNotificationMigrator({
    firestore,
    sheetsGateway,
    serverTimestamp: () => "SERVER_TIMESTAMP"
  });
  const dryRun = await migrator.migrate({ dryRun: true });
  assert.equal(dryRun.activeUnique, 1);
  assert.equal(dryRun.imported, 0);
  assert.equal(firestore.documents.size, 0);

  const migrated = await migrator.migrate({ dryRun: false });
  assert.equal(migrated.imported, 1);
  const tokenHash = hash(tokenA);
  const device = firestore.documents.get(`notificationDevices/${legacyDeviceId(tokenA)}`);
  assert.equal(device.legacy, true);
  assert.equal(device.keywords, "부고");
  assert.equal(firestore.documents.get(`notificationTokenOwners/${tokenHash}`).deviceId, legacyDeviceId(tokenA));
});

test("migration never overwrites a token already owned by a regular Firebase device", async () => {
  const tokenHash = hash(tokenA);
  const currentId = `d_${"z".repeat(43)}`;
  const firestore = new FakeFirestore({
    [`notificationTokenOwners/${tokenHash}`]: { tokenHash, deviceId: currentId, legacy: false },
    [`notificationDevices/${currentId}`]: { deviceId: currentId, token: tokenA, tokenHash, legacy: false, active: true }
  });
  const migrator = createLegacyNotificationMigrator({
    firestore,
    sheetsGateway: { async readRange() { return [[tokenA, "Chrome", "", "부고", true]]; } },
    serverTimestamp: () => "SERVER_TIMESTAMP"
  });
  const result = await migrator.migrate({ dryRun: false });
  assert.equal(result.protected, 1);
  assert.equal(firestore.documents.get(`notificationDevices/${currentId}`).legacy, false);
});
