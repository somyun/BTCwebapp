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
      publishSubmissionSnapshot: async () => {
        publishCalls += 1;
        return { form: { status: "published" }, list: { status: "published" } };
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

test("a retryable Sheets failure is tracked and can be retried safely", async () => {
  const firestore = makeFirestore();
  let shouldFail = true;
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
      publishSubmissionSnapshot: async () => ({
        form: { status: "published" },
        list: { status: "published" }
      })
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
});

test("retry continues when a previous attempt already published the deterministic target revision", async () => {
  const firestore = makeFirestore();
  const acceptedAt = "2026-07-29T01:02:03.000Z";
  firestore.documents.get(`publicForms/${formKey}`).sourceRevision = acceptedAt;
  const stored = firestore.documents.get(`measurementSubmissions/${submissionId}`);
  stored.status = "failed";
  stored.retryable = true;
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "2026-07-29T01:02:04.000Z",
    sheetsGateway: { syncMeasurements: async () => ({ updatedCellCount: 2 }) },
    publisher: {
      publishSubmissionSnapshot: async () => ({
        form: { status: "published" },
        list: { status: "published" }
      })
    },
    now: () => new Date("2026-07-29T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });
  const synced = await synchronizer.syncSubmission(submissionId, "retry-after-publish");
  assert.equal(synced.status, "synced");
});
