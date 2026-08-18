"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const {
  MAX_CHUNK_DOCUMENT_BYTES,
  PRODUCTION_SPREADSHEET_ID,
  buildFormDocument,
  buildStoragePlan,
  createPublisher,
  formKeyForSheet,
  normalizeFormList,
  normalizeFormRows,
  serializedBytes
} = require("../lib/publisher");
const { FakeFirestore } = require("./helpers/fake-firestore");

const sourceList = [
  {
    sheetName: "율리24",
    spreadsheetId: PRODUCTION_SPREADSHEET_ID,
    lastModifiedDate: "2026-07-29T00:00:00.000Z"
  }
];

test("form key is stable, URL-safe, and sheet-specific", () => {
  const first = formKeyForSheet("율리24");
  assert.equal(first, formKeyForSheet("율리24"));
  assert.notEqual(first, formKeyForSheet("호포24"));
  assert.match(first, /^f_[a-f0-9]{32}$/);
});

test("form list rejects any non-test spreadsheet reference", () => {
  assert.throws(() => normalizeFormList([{ ...sourceList[0], spreadsheetId: "production-id" }]), {
    message: "PRODUCTION_SPREADSHEET_ID_MISMATCH"
  });
});

test("empty GAS responses are rejected so existing cache cannot be erased", () => {
  assert.throws(() => normalizeFormList([]), { message: "EMPTY_FORM_LIST_REJECTED" });
  assert.throws(() => normalizeFormRows([]), { message: "EMPTY_FORM_RESPONSE_REJECTED" });
});

test("form list normalization is deterministic", () => {
  const first = normalizeFormList(sourceList);
  const second = normalizeFormList(structuredClone(sourceList));
  assert.deepEqual(first, second);
  assert.equal(first.itemCount, 1);
  assert.equal(first.items[0].displayName, "율리24");
});

test("form row normalization preserves the GAS response contract", () => {
  const rows = normalizeFormRows([{
    uniqueId: "u1",
    location: "24KV SIS",
    item: "수전점 전압",
    value: "24.00",
    unit: "kV",
    validation: { minValue: "20", maxValue: "26" },
    recentInfo: { value: "23.9", date: "7월28일" }
  }]);
  assert.deepEqual(rows[0].validation, { minValue: "20", maxValue: "26" });
  assert.deepEqual(rows[0].recentInfo, { value: "23.9", date: "7월28일" });
});

test("small forms stay inline and keep a deterministic content hash", () => {
  const list = normalizeFormList(sourceList);
  const document = buildFormDocument(list.items[0], [{ uniqueId: "u1", location: "L", item: "I" }]);
  const plan = buildStoragePlan(document);
  assert.equal(plan.root.storageMode, "inline");
  assert.equal(plan.root.chunkCount, 0);
  assert.equal(plan.root.rows.length, 1);
  assert.equal(document.contentHash, buildFormDocument(list.items[0], [{ uniqueId: "u1", location: "L", item: "I" }]).contentHash);
});

test("large forms are split into bounded chunks without losing rows", () => {
  const list = normalizeFormList(sourceList);
  const sourceRows = Array.from({ length: 50 }, (_, index) => ({
    uniqueId: `u${index}`,
    location: "L".repeat(20_000),
    item: `item-${index}`
  }));
  const document = buildFormDocument(list.items[0], sourceRows);
  const plan = buildStoragePlan(document);
  assert.equal(plan.root.storageMode, "chunked");
  assert.equal(plan.chunks.flatMap((chunk) => chunk.rows).length, sourceRows.length);
  for (const chunk of plan.chunks) {
    assert.ok(serializedBytes({ rows: chunk.rows }) <= MAX_CHUNK_DOCUMENT_BYTES);
  }
});

test("submission publishing writes only the date-specific measurement cache", async () => {
  const initialList = normalizeFormList(sourceList);
  const initialForm = buildFormDocument(initialList.items[0], [{
    uniqueId: "u1",
    location: "L",
    item: "I",
    value: "1",
    unit: "V"
  }]);
  const initialPlan = buildStoragePlan(initialForm);
  const firestore = new FakeFirestore({
    "publicCache/formList": initialList,
    [`publicForms/${initialForm.formKey}`]: initialPlan.root
  });
  const publisher = createPublisher({
    firestore,
    fetchImpl: async () => { throw new Error("GAS_MUST_NOT_BE_CALLED"); },
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {} }
  });
  const nextRevision = "2026-07-29T01:02:03.000Z";
  const result = await publisher.publishDailyMeasurementCache({
    formDocument: initialPlan.root,
    rows: initialForm.rows,
    measurements: [{ uniqueId: "u1", location: "L", item: "I", value: "2", unit: "V" }],
    sourceRevision: nextRevision
  });
  assert.equal(result.status, "published");
  assert.equal(result.cacheDate, "2026-07-29");
  const storedCache = firestore.documents.get(`dailyMeasurementCaches/${result.cacheId}`);
  assert.equal(storedCache.sourceRevision, nextRevision);
  assert.equal(storedCache.measurements[0].value, "2");
  assert.equal(firestore.documents.get(`publicForms/${initialForm.formKey}`).sourceRevision,
    initialForm.sourceRevision);
  assert.equal(firestore.documents.get("publicCache/formList").sourceRevision,
    initialList.sourceRevision);
});

test("daily measurement cache maps sorted measurements by identity, not display order", async () => {
  const initialList = normalizeFormList(sourceList);
  const initialForm = buildFormDocument(initialList.items[0], [
    { uniqueId: "u1", location: "L1", item: "I1", value: "1", unit: "V" },
    { uniqueId: "u2", location: "L2", item: "I2", value: "2", unit: "A" }
  ]);
  const initialPlan = buildStoragePlan(initialForm);
  const firestore = new FakeFirestore({
    "publicCache/formList": initialList,
    [`publicForms/${initialForm.formKey}`]: initialPlan.root
  });
  const publisher = createPublisher({
    firestore,
    fetchImpl: async () => { throw new Error("GAS_MUST_NOT_BE_CALLED"); },
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {} }
  });
  const result = await publisher.publishDailyMeasurementCache({
    formDocument: initialPlan.root,
    rows: initialForm.rows,
    measurements: [
      { uniqueId: "u2", location: "L2", item: "I2", value: "20", unit: "A" },
      { uniqueId: "u1", location: "L1", item: "I1", value: "10", unit: "V" }
    ],
    sourceRevision: "2026-07-29T02:00:00.000Z"
  });
  const storedRows = firestore.documents.get(`dailyMeasurementCaches/${result.cacheId}`).measurements;
  assert.deepEqual(storedRows.map((row) => row.value), ["10", "20"]);
});

test("daily cleanup deletes past measurement caches and keeps today's cache", async () => {
  const firestore = new FakeFirestore({
    "dailyMeasurementCaches/yesterday": { cacheDate: "2026-07-28" },
    "dailyMeasurementCaches/today": { cacheDate: "2026-07-29" }
  });
  const publisher = createPublisher({
    firestore,
    fetchImpl: async () => { throw new Error("GAS_MUST_NOT_BE_CALLED"); },
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {} }
  });
  const result = await publisher.deleteExpiredDailyMeasurementCaches({
    currentDate: "2026-07-29"
  });
  assert.equal(result.deletedCount, 1);
  assert.equal(firestore.documents.has("dailyMeasurementCaches/yesterday"), false);
  assert.equal(firestore.documents.has("dailyMeasurementCaches/today"), true);
});

test("scheduled publishing cannot regress a newer submission revision", async () => {
  const newerRevision = "2026-07-29T01:02:03.000Z";
  const newerList = normalizeFormList([{ ...sourceList[0], lastModifiedDate: newerRevision }]);
  const newerForm = buildFormDocument(newerList.items[0], [{
    uniqueId: "u1", location: "L", item: "I", value: "2", unit: "V"
  }]);
  const firestore = new FakeFirestore({
    "publicCache/formList": newerList,
    [`publicForms/${newerForm.formKey}`]: buildStoragePlan(newerForm).root
  });
  const publisher = createPublisher({
    firestore,
    fetchImpl: async (url) => ({
      ok: true,
      json: async () => url.searchParams.get("action") === "getFormList"
        ? sourceList
        : [{ uniqueId: "u1", location: "L", item: "I", value: "1", unit: "V" }]
    }),
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {} }
  });
  const result = await publisher.publishAllChangedForms({ force: true });
  assert.equal(result.list.status, "stale_skipped");
  assert.equal(result.forms[0].status, "stale_skipped");
  assert.equal(firestore.documents.get("publicCache/formList").sourceRevision, newerRevision);
  assert.equal(firestore.documents.get(`publicForms/${newerForm.formKey}`).rows[0].value, "2");
});
