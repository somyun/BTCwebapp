"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const { formKeyForSheet } = require("../lib/publisher");
const {
  createSubmissionService,
  normalizeSubmissionRequest,
  publicSubmissionStatus,
  validateAgainstPublishedForm
} = require("../lib/submission");
const { FakeFirestore } = require("./helpers/fake-firestore");

const sheetName = "율리24";
const formKey = formKeyForSheet(sheetName);
const revision = "2026-07-29T00:00:00.000Z";
const rows = [{
  uniqueId: "u1",
  location: "24KV SIS",
  item: "수전점 전압",
  value: "24.00",
  unit: "kV"
}];
const request = {
  schemaVersion: 1,
  idempotencyKey: "123e4567-e89b-12d3-a456-426614174000",
  formKey,
  sheetName,
  formRevision: revision,
  measurements: [{
    uniqueId: "u1",
    location: "24KV SIS",
    item: "수전점 전압",
    value: "24",
    unit: "kV"
  }]
};

function seedFirestore() {
  return new FakeFirestore({
    [`publicForms/${formKey}`]: {
      schemaVersion: 1,
      formKey,
      sheetName,
      sourceRevision: revision,
      rowCount: 1,
      storageMode: "inline",
      rows
    },
    "systemConfig/submissions": { enabled: true }
  });
}

test("normalizes a bounded submission and hashes its semantic content", () => {
  const normalized = normalizeSubmissionRequest(request);
  assert.equal(normalized.formRevision, revision);
  assert.match(normalized.requestHash, /^[a-f0-9]{64}$/);
  assert.equal(normalized.measurements[0].value, "24");
});

test("rejects invalid idempotency keys and non-numeric values", () => {
  assert.throws(() => normalizeSubmissionRequest({ ...request, idempotencyKey: "short" }), {
    message: "INVALID_IDEMPOTENCY_KEY"
  });
  const invalid = structuredClone(request);
  invalid.measurements[0].value = "24kV";
  assert.throws(() => normalizeSubmissionRequest(invalid), { message: "INVALID_MEASUREMENT_VALUE" });
});

test("rejects stale revisions and row identity changes", () => {
  const normalized = normalizeSubmissionRequest(request);
  const form = { formKey, sheetName, sourceRevision: revision, rowCount: 1 };
  assert.equal(validateAgainstPublishedForm(normalized, form, rows), true);
  assert.throws(() => validateAgainstPublishedForm(
    normalized,
    { ...form, sourceRevision: "2026-07-28T00:00:00.000Z" },
    rows
  ), { message: "STALE_SUBMISSION_FORM_REVISION" });
  assert.throws(() => validateAgainstPublishedForm(
    normalized,
    form,
    [{ ...rows[0], item: "다른 항목" }]
  ), { message: "SUBMISSION_ROW_IDENTITY_MISMATCH" });
});

test("accepts the same published rows after the UI changes their display order", () => {
  const secondRow = {
    uniqueId: "u2",
    location: "정류기",
    item: "출력 전압",
    value: "750",
    unit: "V"
  };
  const reorderedRequest = {
    ...request,
    measurements: [
      { ...secondRow, value: "751" },
      request.measurements[0]
    ]
  };
  const normalized = normalizeSubmissionRequest(reorderedRequest);
  const form = { formKey, sheetName, sourceRevision: revision, rowCount: 2 };
  assert.equal(validateAgainstPublishedForm(normalized, form, [...rows, secondRow]), true);
});

test("same idempotency key returns the original receipt and conflicting content is rejected", async () => {
  const firestore = seedFirestore();
  const service = createSubmissionService({
    firestore,
    serverTimestamp: () => "SERVER_TIMESTAMP",
    now: () => new Date("2026-07-29T01:02:03.000Z")
  });
  const first = await service.submit(request);
  const duplicate = await service.submit(structuredClone(request));
  assert.equal(first.created, true);
  assert.equal(duplicate.created, false);
  assert.equal(duplicate.status.id, request.idempotencyKey);

  const conflicting = structuredClone(request);
  conflicting.measurements[0].value = "25";
  await assert.rejects(service.submit(conflicting), { message: "IDEMPOTENCY_KEY_CONFLICT" });
});

test("public status never exposes measurements or request hashes", () => {
  const status = publicSubmissionStatus(request.idempotencyKey, {
    status: "queued",
    sheetName,
    formKey,
    measurements: request.measurements,
    requestHash: "secret-hash",
    acceptedAt: revision
  });
  assert.equal(status.status, "queued");
  assert.equal("measurements" in status, false);
  assert.equal("requestHash" in status, false);
});

test("submission gate is closed when disabled and can restrict one production form", async () => {
  const firestore = seedFirestore();
  const service = createSubmissionService({
    firestore,
    serverTimestamp: () => "SERVER_TIMESTAMP",
    now: () => new Date("2026-07-29T01:02:03.000Z")
  });

  await service.setGate(false);
  await assert.rejects(service.submit(request), { message: "SUBMISSIONS_DISABLED" });

  await service.setGate(true, { allowedFormKeys: [`f_${"b".repeat(32)}`] });
  await assert.rejects(service.submit(request), { message: "SUBMISSION_FORM_NOT_ALLOWED" });

  const gate = await service.setGate(true, { allowedFormKeys: [formKey] });
  assert.deepEqual(gate.allowedFormKeys, [formKey]);
  assert.equal((await service.submit(request)).created, true);
});
