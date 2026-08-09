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

const sourceList = [{
  sheetName: "동래24",
  spreadsheetId: PRODUCTION_SPREADSHEET_ID,
  lastModifiedDate: "2026-07-29T00:00:00.000Z"
}];

test("production form keys are deterministic and URL-safe", () => {
  const first = formKeyForSheet("동래24");
  assert.equal(first, formKeyForSheet("동래24"));
  assert.notEqual(first, formKeyForSheet("명륜24"));
  assert.match(first, /^f_[a-f0-9]{32}$/);
});

test("publisher rejects a response from any other spreadsheet", () => {
  assert.throws(() => normalizeFormList([{ ...sourceList[0], spreadsheetId: "wrong-id" }]), {
    message: "PRODUCTION_SPREADSHEET_ID_MISMATCH"
  });
});

test("empty source responses cannot erase the current cache", () => {
  assert.throws(() => normalizeFormList([]), { message: "EMPTY_FORM_LIST_REJECTED" });
  assert.throws(() => normalizeFormRows([]), { message: "EMPTY_FORM_RESPONSE_REJECTED" });
});

test("small forms remain inline", () => {
  const list = normalizeFormList(sourceList);
  const document = buildFormDocument(list.items[0], [{ uniqueId: "u1", location: "L", item: "I" }]);
  const plan = buildStoragePlan(document);
  assert.equal(plan.root.storageMode, "inline");
  assert.equal(plan.root.chunkCount, 0);
  assert.equal(plan.root.rows.length, 1);
});

test("large forms are chunked without data loss", () => {
  const list = normalizeFormList(sourceList);
  const sourceRows = Array.from({ length: 50 }, (_, index) => ({
    uniqueId: `u${index}`,
    location: "L".repeat(20_000),
    item: `item-${index}`
  }));
  const plan = buildStoragePlan(buildFormDocument(list.items[0], sourceRows));
  assert.equal(plan.root.storageMode, "chunked");
  assert.equal(plan.chunks.flatMap((chunk) => chunk.rows).length, sourceRows.length);
  plan.chunks.forEach((chunk) => {
    assert.ok(serializedBytes({ rows: chunk.rows }) <= MAX_CHUNK_DOCUMENT_BYTES);
  });
});

test("unchanged revisions skip the expensive GAS form-data request", async () => {
  const normalizedList = normalizeFormList(sourceList);
  const item = normalizedList.items[0];
  const existingForm = buildStoragePlan(buildFormDocument(item, [{
    uniqueId: "u1", location: "L", item: "I", value: "1", unit: "V"
  }])).root;
  const firestore = new FakeFirestore({
    "publicCache/formList": normalizedList,
    [`publicForms/${item.formKey}`]: existingForm
  });
  const requestedActions = [];
  const publisher = createPublisher({
    firestore,
    fetchImpl: async (url) => {
      requestedActions.push(url.searchParams.get("action"));
      return { ok: true, json: async () => sourceList };
    },
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {} }
  });

  const result = await publisher.publishAllChangedForms();
  assert.deepEqual(requestedActions, ["getFormList"]);
  assert.equal(result.forms[0].status, "unchanged_revision");
  assert.equal(result.list.status, "unchanged");
});

test("a changed revision refreshes only that form and then the list", async () => {
  const oldList = normalizeFormList(sourceList);
  const oldItem = oldList.items[0];
  const newSourceList = [{ ...sourceList[0], lastModifiedDate: "2026-07-29T01:00:00.000Z" }];
  const existingForm = buildStoragePlan(buildFormDocument(oldItem, [{
    uniqueId: "u1", location: "L", item: "I", value: "1", unit: "V"
  }])).root;
  const firestore = new FakeFirestore({
    "publicCache/formList": oldList,
    [`publicForms/${oldItem.formKey}`]: existingForm
  });
  const requestedActions = [];
  const publisher = createPublisher({
    firestore,
    fetchImpl: async (url) => {
      const action = url.searchParams.get("action");
      requestedActions.push(action);
      return {
        ok: true,
        json: async () => action === "getFormList"
          ? newSourceList
          : [{ uniqueId: "u1", location: "L", item: "I", value: "2", unit: "V" }]
      };
    },
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {} }
  });

  const result = await publisher.publishAllChangedForms();
  assert.deepEqual(requestedActions, ["getFormList", "getFormDataForWeb"]);
  assert.equal(result.forms[0].status, "published");
  assert.equal(result.list.status, "published");
  assert.equal(firestore.documents.get(`publicForms/${oldItem.formKey}`).rows[0].value, "2");
});
