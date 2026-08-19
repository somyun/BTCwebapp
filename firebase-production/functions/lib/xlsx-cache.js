"use strict";

const { createHash, randomUUID } = require("node:crypto");
const {
  PRODUCTION_GAS_API_URL,
  PRODUCTION_SPREADSHEET_ID
} = require("./publisher");

const GAS_XLSX_TIMEOUT_MS = 90_000;
const MAX_XLSX_BYTES = 20 * 1024 * 1024;
const MAX_READY_REVISIONS = 5;
const MAX_JOB_ATTEMPTS = 5;

function requireString(value, fieldName) {
  const normalized = String(value ?? "").trim().normalize("NFC");
  if (!normalized) throw new Error(`INVALID_${fieldName.toUpperCase()}`);
  return normalized;
}

function normalizeRevision(value) {
  const revision = requireString(value, "revision");
  const timestamp = Date.parse(revision);
  if (!Number.isFinite(timestamp)) throw new Error("INVALID_XLSX_REVISION");
  return new Date(timestamp).toISOString();
}

function revisionId(revision) {
  return createHash("sha256").update(normalizeRevision(revision), "utf8").digest("hex").slice(0, 32);
}

function revisionTime(revision) {
  const timestamp = Date.parse(revision);
  return Number.isFinite(timestamp) ? timestamp : Number.NEGATIVE_INFINITY;
}

function dateKey(revision) {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Seoul",
    year: "2-digit",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date(normalizeRevision(revision)));
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}${values.month}${values.day}`;
}

function filenameFor(sheetName, revision) {
  const displayName = requireString(sheetName, "sheetName").split("_")[0];
  return `${displayName}_${dateKey(revision)}.xlsx`;
}

function storagePathFor(formKey, revision) {
  return `xlsx-cache/${requireString(formKey, "formKey")}/${revisionId(revision)}.xlsx`;
}

function downloadUrlFor(bucketName, storagePath, token) {
  return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucketName)}` +
    `/o/${encodeURIComponent(storagePath)}?alt=media&token=${encodeURIComponent(token)}`;
}

function isRetryable(error) {
  return /(?:SOURCE_REVISION_MISMATCH|429|5\d\d|ECONN|ETIMEDOUT|TIMEOUT|ABORT)/i
    .test(String(error?.message || ""));
}

function decodeXlsx(base64) {
  const text = requireString(base64, "xlsxBase64").replace(/\s/g, "");
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(text)) throw new Error("INVALID_XLSX_BASE64");
  const buffer = Buffer.from(text, "base64");
  if (!buffer.length || buffer.length > MAX_XLSX_BYTES) throw new Error("INVALID_XLSX_SIZE");
  if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) throw new Error("INVALID_XLSX_SIGNATURE");
  return buffer;
}

function createXlsxCacheService({
  firestore,
  bucket,
  fetchImpl = globalThis.fetch,
  serverTimestamp,
  now = () => new Date(),
  logger = console
}) {
  if (!firestore || typeof firestore.collection !== "function") throw new Error("FIRESTORE_REQUIRED");
  if (!bucket || typeof bucket.file !== "function" || !bucket.name) throw new Error("STORAGE_BUCKET_REQUIRED");
  if (typeof fetchImpl !== "function") throw new Error("FETCH_REQUIRED");
  if (typeof serverTimestamp !== "function") throw new Error("SERVER_TIMESTAMP_REQUIRED");

  async function enqueue({ formKey, sheetName, revision, force = false }) {
    const normalizedFormKey = requireString(formKey, "formKey");
    const normalizedSheetName = requireString(sheetName, "sheetName");
    const normalizedRevision = normalizeRevision(revision);
    const stableRevisionId = revisionId(normalizedRevision);
    const jobId = force ? `${stableRevisionId}_${now().getTime()}` : stableRevisionId;
    const jobRef = firestore.collection("xlsxExportJobs").doc(`${normalizedFormKey}_${jobId}`);
    const rootRef = firestore.collection("xlsxExports").doc(normalizedFormKey);

    const result = await firestore.runTransaction(async (transaction) => {
      const [existingJob, existingRoot] = await Promise.all([
        transaction.get(jobRef),
        transaction.get(rootRef)
      ]);
      if (existingJob.exists) {
        return { status: existingJob.get("status"), jobId: existingJob.id, created: false };
      }
      transaction.create(jobRef, {
        schemaVersion: 1,
        formKey: normalizedFormKey,
        sheetName: normalizedSheetName,
        spreadsheetId: PRODUCTION_SPREADSHEET_ID,
        revision: normalizedRevision,
        revisionId: stableRevisionId,
        filename: filenameFor(normalizedSheetName, normalizedRevision),
        status: "pending",
        attemptCount: 0,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      const currentDesired = existingRoot.get("desiredRevision");
      if (!currentDesired || revisionTime(normalizedRevision) >= revisionTime(currentDesired)) {
        transaction.set(rootRef, {
          schemaVersion: 1,
          formKey: normalizedFormKey,
          sheetName: normalizedSheetName,
          desiredRevision: normalizedRevision,
          status: "preparing",
          updatedAt: serverTimestamp()
        }, { merge: true });
      }
      return { status: "pending", jobId: jobRef.id, created: true };
    });
    return result;
  }

  async function fetchGasXlsx(job) {
    const url = new URL(PRODUCTION_GAS_API_URL);
    url.searchParams.set("fileId", job.spreadsheetId);
    url.searchParams.set("sheetName", job.sheetName);
    url.searchParams.set("filename", job.filename);
    url.searchParams.set("expectedRevision", job.revision);
    const response = await fetchImpl(url, {
      method: "GET",
      redirect: "follow",
      headers: { Accept: "application/json" },
      signal: AbortSignal.timeout(GAS_XLSX_TIMEOUT_MS)
    });
    if (!response.ok) throw new Error(`GAS_XLSX_HTTP_${response.status}`);
    const payload = await response.json();
    if (payload?.error) throw new Error(`GAS_XLSX_ERROR_${String(payload.error).slice(0, 200)}`);
    if (!payload || typeof payload !== "object") throw new Error("INVALID_GAS_XLSX_RESPONSE");
    const filename = requireString(payload.filename, "xlsxFilename");
    if (filename !== job.filename) throw new Error("XLSX_FILENAME_MISMATCH");
    return decodeXlsx(payload.base64);
  }

  async function publishReady(job, buffer) {
    const storagePath = storagePathFor(job.formKey, job.revision);
    const downloadToken = randomUUID();
    const file = bucket.file(storagePath);
    await file.save(buffer, {
      resumable: false,
      validation: "crc32c",
      metadata: {
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        contentDisposition: `attachment; filename*=UTF-8''${encodeURIComponent(job.filename)}`,
        cacheControl: "public, max-age=300",
        metadata: {
          firebaseStorageDownloadTokens: downloadToken,
          formKey: job.formKey,
          revision: job.revision
        }
      }
    });

    const revisionRef = firestore.collection("xlsxExports").doc(job.formKey)
      .collection("revisions").doc(job.revisionId);
    await revisionRef.set({
      schemaVersion: 1,
      formKey: job.formKey,
      sheetName: job.sheetName,
      revision: job.revision,
      revisionId: job.revisionId,
      filename: job.filename,
      storagePath,
      downloadToken,
      size: buffer.length,
      status: "ready",
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });

    const rootRef = firestore.collection("xlsxExports").doc(job.formKey);
    await firestore.runTransaction(async (transaction) => {
      const current = await transaction.get(rootRef);
      const latestRevision = current.get("latestReadyRevision");
      const desiredRevision = current.get("desiredRevision");
      const update = {
        updatedAt: serverTimestamp()
      };
      if (!latestRevision || revisionTime(job.revision) >= revisionTime(latestRevision)) {
        Object.assign(update, {
          latestReadyRevision: job.revision,
          latestRevisionId: job.revisionId,
          filename: job.filename,
          storagePath,
          downloadToken,
          size: buffer.length
        });
      }
      if (!desiredRevision || revisionTime(job.revision) >= revisionTime(desiredRevision)) {
        update.status = "ready";
      }
      transaction.set(rootRef, update, { merge: true });
    });
    return { storagePath, downloadToken };
  }

  async function pruneOldRevisions(formKey) {
    const revisionsRef = firestore.collection("xlsxExports").doc(formKey).collection("revisions");
    const snapshot = await revisionsRef.get();
    const ready = snapshot.docs
      .filter((document) => document.get("status") === "ready")
      .sort((left, right) => revisionTime(right.get("revision")) - revisionTime(left.get("revision")));
    const expired = ready.slice(MAX_READY_REVISIONS);
    for (const document of expired) {
      const storagePath = document.get("storagePath");
      await document.ref.set({ status: "deleting", updatedAt: serverTimestamp() }, { merge: true });
      if (storagePath) await bucket.file(storagePath).delete({ ignoreNotFound: true });
      await document.ref.delete();
    }
    return expired.length;
  }

  async function processJob(jobId) {
    const jobRef = firestore.collection("xlsxExportJobs").doc(requireString(jobId, "jobId"));
    const snapshot = await jobRef.get();
    if (!snapshot.exists) throw new Error("XLSX_JOB_NOT_FOUND");
    const job = snapshot.data();
    if (job.status === "ready") return { status: "ready", jobId };
    const attemptCount = Number(job.attemptCount || 0) + 1;
    await jobRef.update({
      status: "processing",
      attemptCount,
      errorCode: null,
      startedAt: serverTimestamp(),
      updatedAt: serverTimestamp()
    });
    try {
      const buffer = await fetchGasXlsx(job);
      const stored = await publishReady(job, buffer);
      await jobRef.update({
        status: "ready",
        storagePath: stored.storagePath,
        completedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      let prunedCount = 0;
      try {
        prunedCount = await pruneOldRevisions(job.formKey);
      } catch (cleanupError) {
        logger.warn("Old XLSX cache cleanup failed", {
          formKey: job.formKey,
          error: String(cleanupError?.message || cleanupError)
        });
      }
      logger.info("XLSX cache prepared", {
        formKey: job.formKey,
        revision: job.revision,
        size: buffer.length,
        prunedCount
      });
      return { status: "ready", jobId, size: buffer.length, prunedCount };
    } catch (error) {
      const retryable = isRetryable(error) && attemptCount < MAX_JOB_ATTEMPTS;
      await jobRef.update({
        status: "failed",
        retryable,
        errorCode: String(error?.message || "UNKNOWN_XLSX_ERROR").slice(0, 200),
        failedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      const rootRef = firestore.collection("xlsxExports").doc(job.formKey);
      await firestore.runTransaction(async (transaction) => {
        const root = await transaction.get(rootRef);
        if (root.exists && root.get("desiredRevision") === job.revision) {
          transaction.set(rootRef, {
            status: retryable ? "preparing" : "failed",
            errorCode: String(error?.message || "UNKNOWN_XLSX_ERROR").slice(0, 200),
            updatedAt: serverTimestamp()
          }, { merge: true });
        }
      });
      logger.error("XLSX cache preparation failed", {
        jobId,
        formKey: job.formKey,
        attemptCount,
        retryable,
        error: String(error?.message || error)
      });
      if (retryable) throw error;
      return { status: "failed", jobId, retryable: false };
    }
  }

  async function getDownload({ formKey, expectedRevision } = {}) {
    const normalizedFormKey = requireString(formKey, "formKey");
    const root = await firestore.collection("xlsxExports").doc(normalizedFormKey).get();
    if (!root.exists) return { status: "unavailable", latest: null };
    const latestRevision = root.get("latestReadyRevision");
    const latest = latestRevision && root.get("storagePath") && root.get("downloadToken")
      ? {
          filename: root.get("filename"),
          revision: latestRevision,
          size: Number(root.get("size") || 0),
          downloadUrl: downloadUrlFor(
            bucket.name,
            root.get("storagePath"),
            root.get("downloadToken")
          )
        }
      : null;
    const expected = expectedRevision ? normalizeRevision(expectedRevision) : null;
    const readyForExpected = !expected || (latest && revisionTime(latest.revision) >= revisionTime(expected));
    const rootStatus = root.get("status");
    return {
      status: readyForExpected && latest
        ? "ready"
        : (rootStatus === "failed" ? "failed" : "preparing"),
      desiredRevision: root.get("desiredRevision") || null,
      errorCode: root.get("errorCode") || null,
      latest
    };
  }

  return {
    enqueue,
    getDownload,
    processJob,
    pruneOldRevisions
  };
}

module.exports = {
  GAS_XLSX_TIMEOUT_MS,
  MAX_JOB_ATTEMPTS,
  MAX_READY_REVISIONS,
  MAX_XLSX_BYTES,
  createXlsxCacheService,
  dateKey,
  decodeXlsx,
  downloadUrlFor,
  filenameFor,
  isRetryable,
  revisionId,
  storagePathFor
};
