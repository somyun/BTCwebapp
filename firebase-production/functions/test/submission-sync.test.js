"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const { PRODUCTION_SPREADSHEET_ID, formKeyForSheet } = require("../lib/publisher");
const {
  buildSheetsBatchUpdate,
  createSubmissionService,
  createSynchronizer,
  normalizeSubmissionRequest,
  validateAgainstPublishedForm
} = require("../lib/submission-sync");
const { FakeFirestore } = require("./helpers/fake-firestore");

const sheetName = "호포24";
const formKey = formKeyForSheet(sheetName);
const revision = "2026-08-10T00:00:00.000Z";
const submissionId = "123e4567-e89b-12d3-a456-426614174000";
const rows = [
  { uniqueId: "u1", location: "24KV SIS", item: "전압", value: "24", unit: "kV" },
  { uniqueId: "u2", location: "정류기", item: "전류", value: "10", unit: "A" }
];
const request = {
  schemaVersion: 1,
  idempotencyKey: submissionId,
  formKey,
  sheetName,
  formRevision: revision,
  measurements: [
    { uniqueId: "u2", location: "정류기", item: "전류", value: "11", unit: "A" },
    { uniqueId: "u1", location: "24KV SIS", item: "전압", value: "24.1", unit: "kV" }
  ]
};

function publishedForm() {
  return {
    schemaVersion: 1,
    formKey,
    sheetName,
    sourceRevision: revision,
    rowCount: rows.length,
    storageMode: "inline",
    rows
  };
}

test("production submissions accept UI row reordering while preserving row identity", () => {
  const normalized = normalizeSubmissionRequest(request);
  assert.equal(validateAgainstPublishedForm(normalized, publishedForm(), rows), true);
  assert.match(normalized.requestHash, /^[a-f0-9]{64}$/);
});

test("production Sheets update writes measurement F cells and the FormList revision", () => {
  const update = buildSheetsBatchUpdate({
    sheetName,
    measurements: request.measurements,
    revision: "2026-08-10T01:02:03.000Z",
    sheetRows: [
      ["u1", "24KV SIS", "전압", "old", "kV", "", "", "2"],
      ["u2", "정류기", "전류", "old", "A", "", "", "0"]
    ],
    formListRows: [[sheetName, PRODUCTION_SPREADSHEET_ID, revision]]
  });
  assert.deepEqual(update.data, [
    { range: "'호포24'!F3", majorDimension: "ROWS", values: [["11"]] },
    { range: "'호포24'!F2", majorDimension: "ROWS", values: [["24.10"]] },
    { range: "'FormList'!C2", majorDimension: "ROWS", values: [["2026-08-10T01:02:03.000Z"]] }
  ]);
});

test("production submission service queues once and rejects conflicting reuse", async () => {
  const firestore = new FakeFirestore({
    [`publicForms/${formKey}`]: publishedForm(),
    "systemConfig/submissions": { enabled: true }
  });
  const service = createSubmissionService({
    firestore,
    serverTimestamp: () => "SERVER_TIMESTAMP",
    now: () => new Date("2026-08-10T01:02:03.000Z")
  });
  const first = await service.submit(request);
  const duplicate = await service.submit(structuredClone(request));
  assert.equal(first.created, true);
  assert.equal(duplicate.created, false);
  const changed = structuredClone(request);
  changed.measurements[0].value = "12";
  await assert.rejects(service.submit(changed), { message: "IDEMPOTENCY_KEY_CONFLICT" });
});

test("production worker performs the Sheets side effect only once after sync", async () => {
  const acceptedAt = "2026-08-10T01:02:03.000Z";
  const firestore = new FakeFirestore({
    [`publicForms/${formKey}`]: publishedForm(),
    [`measurementSubmissions/${submissionId}`]: {
      ...normalizeSubmissionRequest(request),
      status: "queued",
      attemptCount: 0,
      acceptedAt
    }
  });
  let sheetCalls = 0;
  let publishCalls = 0;
  const synchronizer = createSynchronizer({
    firestore,
    serverTimestamp: () => "SERVER_TIMESTAMP",
    sheetsGateway: {
      syncMeasurements: async () => {
        sheetCalls += 1;
        return { updatedCellCount: 3 };
      }
    },
    publisher: {
      publishSubmissionSnapshot: async () => {
        publishCalls += 1;
        return { form: { status: "published" }, list: { status: "published" } };
      }
    },
    now: () => new Date("2026-08-10T01:02:04.000Z"),
    logger: { info() {}, error() {} }
  });
  const first = await synchronizer.syncSubmission(submissionId, "event-1");
  const repeated = await synchronizer.syncSubmission(submissionId, "event-1");
  assert.equal(first.status, "synced");
  assert.equal(repeated.status, "synced");
  assert.equal(sheetCalls, 1);
  assert.equal(publishCalls, 1);
});
