"use strict";

const { createHash } = require("node:crypto");

function requiredText(value, fieldName, maxLength = 200) {
  const text = String(value ?? "").trim().normalize("NFC");
  if (!text || text.length > maxLength) throw new Error(`INVALID_${fieldName.toUpperCase()}`);
  return text;
}

function optionalRevision(value) {
  if (value === null || value === undefined || value === "") return null;
  const timestamp = Date.parse(String(value));
  if (!Number.isFinite(timestamp)) throw new Error("INVALID_REVISION");
  return new Date(timestamp).toISOString();
}

function jobIdFor({ sheetName, revision, eventId, source }) {
  const identity = eventId
    ? `event:${requiredText(eventId, "eventId", 200)}`
    : `change:${source || "unknown"}:${sheetName}:${revision || "latest"}`;
  return `fp_${createHash("sha256").update(identity, "utf8").digest("hex").slice(0, 40)}`;
}

function createFormPublishQueue({ firestore, publisher, serverTimestamp, logger = console }) {
  if (!firestore || typeof firestore.runTransaction !== "function") throw new Error("FIRESTORE_REQUIRED");
  if (!publisher || typeof publisher.publishFormAndList !== "function") throw new Error("PUBLISHER_REQUIRED");
  if (typeof serverTimestamp !== "function") throw new Error("SERVER_TIMESTAMP_REQUIRED");

  async function enqueue(payload = {}) {
    const sheetName = requiredText(payload.sheetName, "sheetName", 120);
    const revision = optionalRevision(payload.revision);
    const source = requiredText(payload.source || "unspecified", "source", 80);
    const eventId = payload.eventId ? requiredText(payload.eventId, "eventId", 200) : null;
    const jobId = jobIdFor({ sheetName, revision, eventId, source });
    const reference = firestore.collection("formPublishJobs").doc(jobId);
    const result = await firestore.runTransaction(async (transaction) => {
      const existing = await transaction.get(reference);
      if (existing.exists) {
        return { created: false, jobId, status: existing.get("status") || "unknown" };
      }
      transaction.create(reference, {
        schemaVersion: 1,
        sheetName,
        revision,
        source,
        eventId,
        force: payload.force === true,
        status: "queued",
        attemptCount: 0,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      return { created: true, jobId, status: "queued" };
    });
    return result;
  }

  async function process(jobId) {
    const reference = firestore.collection("formPublishJobs").doc(requiredText(jobId, "jobId", 100));
    const snapshot = await reference.get();
    if (!snapshot.exists) throw new Error("FORM_PUBLISH_JOB_NOT_FOUND");
    const job = snapshot.data();
    if (job.status === "published") return { skipped: true, jobId, status: "published" };
    await reference.update({
      status: "processing",
      attemptCount: Number(job.attemptCount || 0) + 1,
      errorCode: null,
      startedAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });
    try {
      const result = await publisher.publishFormAndList({
        sheetName: job.sheetName,
        force: job.force === true
      });
      await reference.update({
        status: "published",
        result,
        errorCode: null,
        publishedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      logger.info("Form publish job completed", { jobId, sheetName: job.sheetName });
      return { jobId, status: "published", result };
    } catch (error) {
      const errorCode = String(error?.message || "FORM_PUBLISH_FAILED").slice(0, 200);
      await reference.update({
        status: "failed",
        errorCode,
        failedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      logger.error("Form publish job failed", { jobId, sheetName: job.sheetName, errorCode });
      throw error;
    }
  }

  return { enqueue, process };
}

module.exports = {
  createFormPublishQueue,
  jobIdFor,
  optionalRevision,
  requiredText
};
