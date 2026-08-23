"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { createFormPublishQueue, jobIdFor, optionalRevision } = require("../lib/form-publish-jobs");

function fakeFirestore() {
  const documents = new Map();
  function reference(path) {
    return {
      id: path.split("/").at(-1),
      async get() {
        const value = documents.get(path);
        return {
          exists: value !== undefined,
          data: () => value,
          get: (field) => value?.[field]
        };
      },
      async update(patch) {
        if (!documents.has(path)) throw new Error("NOT_FOUND");
        documents.set(path, { ...documents.get(path), ...patch });
      }
    };
  }
  return {
    documents,
    collection(name) {
      return { doc: (id) => reference(`${name}/${id}`) };
    },
    async runTransaction(handler) {
      return handler({
        get: (ref) => ref.get(),
        create(ref, value) {
          const path = `formPublishJobs/${ref.id}`;
          if (documents.has(path)) throw new Error("ALREADY_EXISTS");
          documents.set(path, value);
        }
      });
    }
  };
}

test("job identity is deterministic for an event", () => {
  const input = {
    sheetName: "점검표",
    revision: "2026-08-23T01:02:03.000Z",
    eventId: "submission-1",
    source: "measurement-submission"
  };
  assert.equal(jobIdFor(input), jobIdFor({ ...input, revision: "2026-08-24T00:00:00.000Z" }));
  assert.equal(optionalRevision("2026-08-23T10:02:03+09:00"), "2026-08-23T01:02:03.000Z");
});

test("enqueue is idempotent and processing publishes once", async () => {
  const firestore = fakeFirestore();
  let publishCount = 0;
  const queue = createFormPublishQueue({
    firestore,
    publisher: {
      async publishFormAndList({ sheetName }) {
        publishCount += 1;
        return { form: { sheetName, status: "published" }, list: { status: "published" } };
      }
    },
    serverTimestamp: () => "SERVER_TIMESTAMP"
  });
  const payload = {
    sheetName: "점검표",
    revision: "2026-08-23T01:02:03.000Z",
    eventId: "submission-1",
    source: "measurement-submission"
  };
  const first = await queue.enqueue(payload);
  const duplicate = await queue.enqueue(payload);
  assert.equal(first.created, true);
  assert.equal(duplicate.created, false);
  assert.equal(first.jobId, duplicate.jobId);

  const result = await queue.process(first.jobId);
  const repeated = await queue.process(first.jobId);
  assert.equal(result.status, "published");
  assert.equal(repeated.skipped, true);
  assert.equal(publishCount, 1);
});
