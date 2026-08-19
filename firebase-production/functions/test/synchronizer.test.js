"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const { formKeyForSheet } = require("../lib/publisher");
const { createSynchronizer } = require("../lib/synchronizer");
const { FakeFirestore } = require("./helpers/fake-firestore");

const sheetName = "율리24";
const formKey = formKeyForSheet(sheetName);
const revision = "2026-07-29T00:00:00.000Z";
const submissionId = "123e4567-e89b-12d3-a456-426614174000";
const measurement = { uniqueId: "u1", location: "L", item: "I", value: "1", unit: "V" };

function makeFirestore() {
  return new FakeFirestore({
    [`publicForms/${formKey}`]: {
      formKey,
      sheetName,
      sourceRevision: revision,
      rowCount: 1,
      storageMode: "inline",
      rows: [measurement]
    },
    [`measurementSubmissions/${submissionId}`]: {
      schemaVersion: 1,
      formKey,
      sheetName,
      formRevision: revision,
      measurements: [measurement],
      status: "queued",
      attemptCount: 0,
      acceptedAt: "2026-07-29T01:02:03.000Z"
    }
  });
}

test("a repeated worker event performs the Sheets side effect once after synced", async () => {
  const firestore = makeFirestore();
  let sheetCalls = 0;
  let publishCalls = 0;
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: {
      syncMeasurements: async () => {
        sheetCalls += 1;
        return { updatedCellCount: 2 };
      }
    },
    publisher: {
      publishDailyMeasurementCache: async () => {
        publishCalls += 1;
        return { status: "published", cacheId: "daily", cacheDate: "2026-07-29" };
      }
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });

  const first = await synchronizer.syncSubmission(submissionId, "event-1");
  const duplicate = await synchronizer.syncSubmission(submissionId, "event-1");
  assert.equal(first.status, "synced");
  assert.equal(duplicate.status, "synced");
  assert.equal(sheetCalls, 1);
  assert.equal(publishCalls, 1);
});

test("publishes the Firebase cache before starting the Sheets update", async () => {
  const firestore = makeFirestore();
  const stages = [];
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: {
      syncMeasurements: async () => {
        stages.push(firestore.documents.get(`measurementSubmissions/${submissionId}`).status);
        return { updatedCellCount: 2 };
      }
    },
    publisher: {
      publishDailyMeasurementCache: async () => {
        stages.push("publisher");
        return { status: "published", cacheId: "daily", cacheDate: "2026-07-29" };
      }
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });

  const result = await synchronizer.syncSubmission(submissionId, "event-cache-stage");
  assert.deepEqual(stages, ["publisher", "cached"]);
  assert.equal(result.status, "synced");
  assert.equal(result.sourceRevisionAfterCache, "2026-07-29T01:02:03.000Z");
});

test("keeps a successful Firebase cache when the later Sheets update fails", async () => {
  const firestore = makeFirestore();
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: { syncMeasurements: async () => { throw new Error("SHEETS_HTTP_400"); } },
    publisher: {
      publishDailyMeasurementCache: async () => ({
        status: "published", cacheId: "daily", cacheDate: "2026-07-29"
      })
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });

  const result = await synchronizer.syncSubmission(submissionId, "event-sheet-failure");
  assert.equal(result.status, "failed");
  assert.equal(result.cachedAt, "2026-07-29T01:02:04.000Z");
  assert.equal(result.sourceRevisionAfterCache, "2026-07-29T01:02:03.000Z");
});

test("a retryable Sheets failure is tracked and can be retried safely", async () => {
  const firestore = makeFirestore();
  let shouldFail = true;
  let publishCalls = 0;
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: {
      syncMeasurements: async () => {
        if (shouldFail) throw new Error("SHEETS_HTTP_503");
        return { updatedCellCount: 2 };
      }
    },
    publisher: {
      publishDailyMeasurementCache: async () => {
        publishCalls += 1;
        return {
          status: "published", cacheId: "daily", cacheDate: "2026-07-29"
        };
      }
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });

  const failed = await synchronizer.syncSubmission(submissionId, "event-1");
  assert.equal(failed.status, "failed");
  assert.equal(failed.retryable, true);
  shouldFail = false;
  const synced = await synchronizer.syncSubmission(submissionId, "manual-retry");
  assert.equal(synced.status, "synced");
  assert.equal(synced.attemptCount, 2);
  assert.equal(publishCalls, 1);
});

test("queued save continues after an earlier save advances the published form revision", async () => {
  const firestore = makeFirestore();
  firestore.documents.get(`publicForms/${formKey}`).sourceRevision = "2026-07-29T01:00:00.000Z";
  const stored = firestore.documents.get(`measurementSubmissions/${submissionId}`);
  stored.status = "failed";
  stored.retryable = true;
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: { syncMeasurements: async () => ({ updatedCellCount: 2 }) },
    publisher: {
      publishDailyMeasurementCache: async () => ({
        status: "published", cacheId: "daily", cacheDate: "2026-07-29"
      })
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });
  const synced = await synchronizer.syncSubmission(submissionId, "retry-after-publish");
  assert.equal(synced.status, "synced");
});

test("queues XLSX preparation after Sheets sync without coupling XLSX failure to save status", async () => {
  const firestore = makeFirestore();
  const queued = [];
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: { syncMeasurements: async () => ({ updatedCellCount: 2 }) },
    publisher: {
      publishDailyMeasurementCache: async () => ({
        status: "published", cacheId: "daily", cacheDate: "2026-07-29"
      })
    },
    xlsxCache: {
      enqueue: async (job) => {
        queued.push(job);
        throw new Error("XLSX_QUEUE_TEMPORARY_FAILURE");
      }
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });
  const result = await synchronizer.syncSubmission(submissionId, "xlsx-queue-failure");
  assert.equal(result.status, "synced");
  assert.equal(result.xlsxStatus, "queue_failed");
  assert.equal(queued.length, 1);
  assert.equal(queued[0].revision, "2026-07-29T01:02:03.000Z");
});
