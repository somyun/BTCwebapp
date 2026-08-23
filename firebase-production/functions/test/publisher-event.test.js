"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { createPublisher, formKeyForSheet, PRODUCTION_SPREADSHEET_ID } = require("../lib/publisher");

function fakeFirestore(seed = {}) {
  const documents = new Map(Object.entries(seed));
  function snapshot(value) {
    return {
      exists: value !== undefined,
      data: () => value,
      get: (field) => value?.[field]
    };
  }
  function reference(path) {
    return {
      path,
      collection(name) {
        return collection(`${path}/${name}`);
      },
      async get() {
        return snapshot(documents.get(path));
      },
      async set(value) {
        documents.set(path, value);
      },
      async delete() {
        documents.delete(path);
      }
    };
  }
  function collection(path) {
    return {
      doc: (id) => reference(`${path}/${id}`),
      async get() {
        const prefix = `${path}/`;
        const docs = [...documents.entries()]
          .filter(([key]) => key.startsWith(prefix) && !key.slice(prefix.length).includes("/"))
          .map(([key, value]) => ({ id: key.slice(prefix.length), ref: reference(key), ...snapshot(value) }));
        return { docs };
      }
    };
  }
  return {
    documents,
    collection,
    async runTransaction(handler) {
      return handler({
        get: (ref) => ref.get(),
        set: (ref, value) => documents.set(ref.path, value)
      });
    }
  };
}

test("event publish advances the form and public list together", async () => {
  const sheetName = "일일점검표";
  const formKey = formKeyForSheet(sheetName);
  const firestore = fakeFirestore();
  const fetchImpl = async (url) => {
    const action = new URL(url).searchParams.get("action");
    const payload = action === "getFormList"
      ? [{
        sheetName,
        spreadsheetId: PRODUCTION_SPREADSHEET_ID,
        lastModifiedDate: "2026-08-23T01:02:03.000Z"
      }]
      : [{
        uniqueId: "A-1",
        location: "기계실",
        item: "압력",
        value: "10",
        unit: "bar",
        validation: null,
        recentInfo: null
      }];
    return { ok: true, json: async () => payload };
  };
  const publisher = createPublisher({
    firestore,
    fetchImpl,
    serverTimestamp: () => "SERVER_TIMESTAMP",
    logger: { info() {}, error() {} }
  });

  const result = await publisher.publishFormAndList({ sheetName });
  const form = firestore.documents.get(`publicForms/${formKey}`);
  const list = firestore.documents.get("publicCache/formList");
  assert.equal(result.form.status, "published");
  assert.equal(form.sourceRevision, "2026-08-23T01:02:03.000Z");
  assert.equal(list.items[0].lastModifiedDate, form.sourceRevision);
  assert.equal(list.items[0].formKey, form.formKey);
});
