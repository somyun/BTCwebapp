"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const { formKeyForSheet } = require("../lib/publisher");
const {
  MAX_READY_REVISIONS,
  createXlsxCacheService,
  revisionId
} = require("../lib/xlsx-cache");
const { FakeFirestore } = require("./helpers/fake-firestore");

class FakeBucket {
  constructor() {
    this.name = "test.firebasestorage.app";
    this.files = new Map();
  }

  file(path) {
    return {
      save: async (buffer, options) => this.files.set(path, { buffer: Buffer.from(buffer), options }),
      delete: async () => this.files.delete(path)
    };
  }
}

function makeService({ firestore = new FakeFirestore(), bucket = new FakeBucket() } = {}) {
  const requests = [];
  const service = createXlsxCacheService({
    firestore,
    bucket,
    fetchImpl: async (url) => {
      requests.push(new URL(url));
      return {
        ok: true,
        json: async () => ({
          filename: new URL(url).searchParams.get("filename"),
          base64: Buffer.from("PK\u0003\u0004fake-xlsx").toString("base64")
        })
      };
    },
    serverTimestamp: () => "2026-08-19T00:00:00.000Z",
    now: () => new Date("2026-08-19T00:00:00.000Z"),
    logger: { info() {}, warn() {}, error() {} }
  });
  return { service, firestore, bucket, requests };
}

test("queues one deterministic job and passes the optional revision guard to GAS", async () => {
  const { service, firestore, requests } = makeService();
  const sheetName = "율리24";
  const formKey = formKeyForSheet(sheetName);
  const revision = "2026-08-19T01:02:03.000Z";
  const first = await service.enqueue({ formKey, sheetName, revision });
  const duplicate = await service.enqueue({ formKey, sheetName, revision });
  assert.equal(first.created, true);
  assert.equal(duplicate.created, false);
  assert.equal(firestore.documents.get(`xlsxExports/${formKey}`).status, "preparing");

  const completed = await service.processJob(first.jobId);
  assert.equal(completed.status, "ready");
  assert.equal(requests.length, 1);
  assert.equal(requests[0].searchParams.get("fileId"),
    "19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s");
  assert.equal(requests[0].searchParams.get("sheetName"), sheetName);
  assert.equal(requests[0].searchParams.get("expectedRevision"), revision);
  const download = await service.getDownload({ formKey, expectedRevision: revision });
  assert.equal(download.status, "ready");
  assert.match(download.latest.downloadUrl, /^https:\/\/firebasestorage\.googleapis\.com\//);
});

test("keeps only the five newest ready revisions for each form", async () => {
  const { service, firestore, bucket } = makeService();
  const sheetName = "율리24";
  const formKey = formKeyForSheet(sheetName);
  const revisions = Array.from({ length: MAX_READY_REVISIONS + 1 }, (_, index) =>
    `2026-08-19T0${index}:00:00.000Z`);
  for (const revision of revisions) {
    const queued = await service.enqueue({ formKey, sheetName, revision });
    await service.processJob(queued.jobId);
  }
  const revisionSnapshot = await firestore.collection("xlsxExports").doc(formKey)
    .collection("revisions").get();
  assert.equal(revisionSnapshot.docs.length, MAX_READY_REVISIONS);
  assert.equal(bucket.files.size, MAX_READY_REVISIONS);
  assert.equal(firestore.documents.has(
    `xlsxExports/${formKey}/revisions/${revisionId(revisions[0])}`
  ), false);
  assert.equal(firestore.documents.get(`xlsxExports/${formKey}`).latestReadyRevision,
    revisions.at(-1));
});

test("a late older job cannot replace the latest ready revision", async () => {
  const { service, firestore } = makeService();
  const sheetName = "율리24";
  const formKey = formKeyForSheet(sheetName);
  const older = await service.enqueue({
    formKey, sheetName, revision: "2026-08-19T01:00:00.000Z"
  });
  const newer = await service.enqueue({
    formKey, sheetName, revision: "2026-08-19T02:00:00.000Z"
  });
  await service.processJob(newer.jobId);
  await service.processJob(older.jobId);
  assert.equal(firestore.documents.get(`xlsxExports/${formKey}`).latestReadyRevision,
    "2026-08-19T02:00:00.000Z");
});
