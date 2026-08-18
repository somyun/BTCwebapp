"use strict";

const { loadPublishedForm, publicSubmissionStatus, validateAgainstPublishedForm } = require("./submission");

const LEASE_MS = 2 * 60 * 1000;

function errorCode(error) {
  return String(error?.message || "UNKNOWN_SYNC_ERROR").slice(0, 200);
}

function isRetryableError(error) {
  const code = errorCode(error);
  return /(?:429|5\d\d|ECONN|ETIMEDOUT|TIMEOUT|ABORT|PRODUCTION_GAS_HTTP_5)/i.test(code);
}

function createSynchronizer({
  firestore,
  serverTimestamp,
  sheetsGateway,
  publisher,
  now = () => new Date(),
  logger = console
}) {
  if (!firestore || typeof firestore.runTransaction !== "function") throw new Error("FIRESTORE_REQUIRED");
  if (typeof serverTimestamp !== "function") throw new Error("SERVER_TIMESTAMP_REQUIRED");
  if (!sheetsGateway || typeof sheetsGateway.syncMeasurements !== "function") throw new Error("SHEETS_GATEWAY_REQUIRED");
  if (!publisher || typeof publisher.publishDailyMeasurementCache !== "function") {
    throw new Error("PUBLISHER_REQUIRED");
  }

  async function acquire(submissionRef, owner) {
    return firestore.runTransaction(async (transaction) => {
      const snapshot = await transaction.get(submissionRef);
      if (!snapshot.exists) throw new Error("SUBMISSION_NOT_FOUND");
      const data = snapshot.data();
      if (data.status === "synced") {
        return { skip: true, status: publicSubmissionStatus(snapshot.id, data) };
      }
      const leaseExpiresAt = Date.parse(data.leaseExpiresAt || "");
      if (["syncing", "caching", "cached"].includes(data.status) && data.syncOwner !== owner &&
          Number.isFinite(leaseExpiresAt) && leaseExpiresAt > now().getTime()) {
        throw new Error("SUBMISSION_SYNC_IN_PROGRESS");
      }
      const startedAt = now();
      transaction.update(submissionRef, {
        status: data.cachedAt ? "cached" : "caching",
        syncOwner: owner,
        leaseExpiresAt: new Date(startedAt.getTime() + LEASE_MS).toISOString(),
        attemptCount: Number(data.attemptCount || 0) + 1,
        retryable: false,
        errorCode: null,
        syncStartedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      return { skip: false, data };
    });
  }

  async function syncSubmission(submissionId, owner, { throwRetryable = false } = {}) {
    const submissionRef = firestore.collection("measurementSubmissions").doc(submissionId);
    const acquired = await acquire(submissionRef, owner);
    if (acquired.skip) return acquired.status;

    try {
      const submission = acquired.data;
      const published = await loadPublishedForm(firestore, submission.formKey);
      const currentRevision = new Date(published.document.sourceRevision).toISOString();
      validateAgainstPublishedForm(
        { ...submission, formRevision: currentRevision },
        published.document,
        published.rows
      );
      let cacheResult = null;
      if (!submission.cachedAt) {
        cacheResult = await publisher.publishDailyMeasurementCache({
          formDocument: published.document,
          rows: published.rows,
          measurements: submission.measurements,
          sourceRevision: submission.acceptedAt
        });
        await submissionRef.update({
          status: "cached",
          sourceRevisionAfterCache: submission.acceptedAt,
          dailyCacheId: cacheResult.cacheId,
          dailyCacheDate: cacheResult.cacheDate,
          cacheStatus: cacheResult.status,
          cachedAt: serverTimestamp(),
          updatedAt: serverTimestamp()
        });
      }
      const sheetResult = await sheetsGateway.syncMeasurements({
        sheetName: submission.sheetName,
        measurements: submission.measurements,
        revision: submission.acceptedAt
      });
      await submissionRef.update({
        status: "synced",
        sourceRevisionAfterSync: submission.acceptedAt,
        updatedCellCount: sheetResult.updatedCellCount,
        retryable: false,
        errorCode: null,
        syncOwner: null,
        leaseExpiresAt: null,
        syncedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      const completed = await submissionRef.get();
      logger.info("Test measurement submission synced", {
        submissionId,
        sheetName: submission.sheetName,
        updatedCellCount: sheetResult.updatedCellCount
      });
      return publicSubmissionStatus(completed.id, completed.data());
    } catch (error) {
      const retryable = isRetryableError(error);
      await submissionRef.update({
        status: "failed",
        retryable,
        errorCode: errorCode(error),
        syncOwner: null,
        leaseExpiresAt: null,
        failedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      logger.error("Test measurement submission sync failed", {
        submissionId,
        retryable,
        errorCode: errorCode(error)
      });
      if (retryable && throwRetryable) throw error;
      const failed = await submissionRef.get();
      return publicSubmissionStatus(failed.id, failed.data());
    }
  }

  return { syncSubmission };
}

module.exports = {
  LEASE_MS,
  createSynchronizer,
  errorCode,
  isRetryableError
};
