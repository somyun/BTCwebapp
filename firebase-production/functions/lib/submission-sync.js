"use strict";

const { PRODUCTION_SPREADSHEET_ID, hashCanonical } = require("./publisher");

const SHEETS_API_BASE = "https://sheets.googleapis.com/v4/spreadsheets";
const SUBMISSION_SCHEMA_VERSION = 1;
const MAX_MEASUREMENTS = 250;
const MAX_REQUEST_BYTES = 128 * 1024;
const RATE_LIMIT_PER_MINUTE = 12;
const LEASE_MS = 2 * 60 * 1000;
const IDEMPOTENCY_KEY_PATTERN = /^[A-Za-z0-9_-]{20,80}$/;
const FORM_KEY_PATTERN = /^f_[a-f0-9]{32}$/;
const NUMERIC_VALUE_PATTERN = /^-?(?:\d+|\d*\.\d+)$/;

function requiredString(value, fieldName, maxLength = 200) {
  const normalized = String(value ?? "").trim().normalize("NFC");
  if (!normalized || normalized.length > maxLength) {
    throw new Error(`INVALID_${fieldName.toUpperCase()}`);
  }
  return normalized;
}

function optionalString(value, maxLength = 200) {
  if (value === null || value === undefined) return "";
  const normalized = String(value).normalize("NFC");
  if (normalized.length > maxLength) throw new Error("INVALID_OPTIONAL_STRING");
  return normalized;
}

function normalizeRevision(value) {
  const timestamp = Date.parse(requiredString(value, "formRevision", 100));
  if (!Number.isFinite(timestamp)) throw new Error("INVALID_FORM_REVISION");
  return new Date(timestamp).toISOString();
}

function comparisonText(value) {
  return String(value ?? "").trim().normalize("NFC").toLocaleLowerCase("ko-KR");
}

function measurementIdentity(entry) {
  const uniqueId = optionalString(entry?.uniqueId, 120).trim();
  return uniqueId
    ? `id:${uniqueId}`
    : `pair:${comparisonText(entry?.location)}|${comparisonText(entry?.item)}`;
}

function normalizeMeasurement(entry) {
  if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
    throw new Error("INVALID_MEASUREMENT");
  }
  const value = optionalString(entry.value, 32).trim();
  if (value && !NUMERIC_VALUE_PATTERN.test(value)) throw new Error("INVALID_MEASUREMENT_VALUE");
  return {
    uniqueId: optionalString(entry.uniqueId, 120).trim(),
    location: requiredString(entry.location, "measurementLocation", 200),
    item: requiredString(entry.item, "measurementItem", 200),
    value,
    unit: optionalString(entry.unit, 80).trim()
  };
}

function normalizeSubmissionRequest(payload) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("INVALID_SUBMISSION_REQUEST");
  }
  if (payload.schemaVersion !== SUBMISSION_SCHEMA_VERSION) {
    throw new Error("UNSUPPORTED_SUBMISSION_SCHEMA");
  }
  const idempotencyKey = requiredString(payload.idempotencyKey, "idempotencyKey", 80);
  if (!IDEMPOTENCY_KEY_PATTERN.test(idempotencyKey)) throw new Error("INVALID_IDEMPOTENCY_KEY");
  const formKey = requiredString(payload.formKey, "formKey", 34);
  if (!FORM_KEY_PATTERN.test(formKey)) throw new Error("INVALID_FORM_KEY");
  if (!Array.isArray(payload.measurements) || !payload.measurements.length ||
      payload.measurements.length > MAX_MEASUREMENTS) {
    throw new Error("INVALID_MEASUREMENT_COUNT");
  }
  const measurements = payload.measurements.map(normalizeMeasurement);
  const identities = new Set();
  for (const measurement of measurements) {
    const identity = measurementIdentity(measurement);
    if (identities.has(identity)) throw new Error("DUPLICATE_MEASUREMENT");
    identities.add(identity);
  }
  const normalized = {
    schemaVersion: SUBMISSION_SCHEMA_VERSION,
    idempotencyKey,
    formKey,
    sheetName: requiredString(payload.sheetName, "sheetName", 100),
    formRevision: normalizeRevision(payload.formRevision),
    measurements
  };
  return { ...normalized, requestHash: hashCanonical(normalized) };
}

async function loadPublishedForm(firestore, formKey) {
  const reference = firestore.collection("publicForms").doc(formKey);
  const snapshot = await reference.get();
  if (!snapshot.exists) throw new Error("PUBLISHED_FORM_NOT_FOUND");
  const document = snapshot.data();
  let rows;
  if (document.storageMode === "inline") {
    rows = document.rows;
  } else if (document.storageMode === "chunked") {
    const chunks = await reference.collection("chunks").get();
    rows = chunks.docs
      .map((entry) => entry.data())
      .sort((left, right) => left.index - right.index)
      .flatMap((chunk) => chunk.rows);
  } else {
    throw new Error("INVALID_PUBLISHED_FORM_STORAGE_MODE");
  }
  if (!Array.isArray(rows) || rows.length !== document.rowCount) {
    throw new Error("PUBLISHED_FORM_ROW_COUNT_MISMATCH");
  }
  return { document, rows };
}

function validateAgainstPublishedForm(submission, formDocument, rows) {
  if (!formDocument || formDocument.formKey !== submission.formKey ||
      formDocument.sheetName !== submission.sheetName) {
    throw new Error("PUBLISHED_FORM_IDENTITY_MISMATCH");
  }
  if (normalizeRevision(formDocument.sourceRevision) !== submission.formRevision) {
    throw new Error("STALE_SUBMISSION_FORM_REVISION");
  }
  if (!Array.isArray(rows) || rows.length !== submission.measurements.length ||
      formDocument.rowCount !== rows.length) {
    throw new Error("SUBMISSION_ROW_COUNT_MISMATCH");
  }
  const publishedByIdentity = new Map(rows.map((row) => [measurementIdentity(row), row]));
  if (publishedByIdentity.size !== rows.length) throw new Error("DUPLICATE_PUBLISHED_ROW");
  for (const measurement of submission.measurements) {
    const published = publishedByIdentity.get(measurementIdentity(measurement));
    if (!published || comparisonText(published.location) !== comparisonText(measurement.location) ||
        comparisonText(published.item) !== comparisonText(measurement.item) ||
        comparisonText(published.unit) !== comparisonText(measurement.unit)) {
      throw new Error("SUBMISSION_ROW_IDENTITY_MISMATCH");
    }
  }
  return true;
}

function timestampToIso(value) {
  if (!value) return null;
  if (typeof value.toDate === "function") return value.toDate().toISOString();
  const timestamp = Date.parse(value);
  return Number.isFinite(timestamp) ? new Date(timestamp).toISOString() : null;
}

function publicSubmissionStatus(id, data) {
  return {
    id,
    status: data.status,
    sheetName: data.sheetName,
    formKey: data.formKey,
    acceptedAt: data.acceptedAt || timestampToIso(data.createdAt),
    syncStartedAt: timestampToIso(data.syncStartedAt),
    syncedAt: timestampToIso(data.syncedAt),
    failedAt: timestampToIso(data.failedAt),
    attemptCount: Number(data.attemptCount || 0),
    retryable: Boolean(data.retryable),
    errorCode: data.errorCode || null,
    sourceRevisionAfterSync: data.sourceRevisionAfterSync || null,
    updatedCellCount: Number(data.updatedCellCount || 0)
  };
}

function minuteBucket(date) {
  return date.toISOString().slice(0, 16).replace(/[-:T]/g, "");
}

function createSubmissionService({ firestore, serverTimestamp, now = () => new Date() }) {
  async function submit(payload) {
    const submission = normalizeSubmissionRequest(payload);
    const published = await loadPublishedForm(firestore, submission.formKey);
    validateAgainstPublishedForm(submission, published.document, published.rows);
    const submissionRef = firestore.collection("measurementSubmissions").doc(submission.idempotencyKey);
    const gateRef = firestore.collection("systemConfig").doc("submissions");
    const acceptedDate = now();
    const rateRef = firestore.collection("submissionRateLimits").doc(`minute_${minuteBucket(acceptedDate)}`);
    return firestore.runTransaction(async (transaction) => {
      const existing = await transaction.get(submissionRef);
      if (existing.exists) {
        const data = existing.data();
        if (data.requestHash !== submission.requestHash) throw new Error("IDEMPOTENCY_KEY_CONFLICT");
        return { created: false, status: publicSubmissionStatus(existing.id, data) };
      }
      const gate = await transaction.get(gateRef);
      if (!gate.exists || gate.get("enabled") !== true) throw new Error("SUBMISSIONS_DISABLED");
      const rate = await transaction.get(rateRef);
      const currentCount = rate.exists ? Number(rate.get("count") || 0) : 0;
      if (currentCount >= RATE_LIMIT_PER_MINUTE) throw new Error("SUBMISSION_RATE_LIMITED");
      const acceptedAt = acceptedDate.toISOString();
      const document = {
        schemaVersion: SUBMISSION_SCHEMA_VERSION,
        requestHash: submission.requestHash,
        formKey: submission.formKey,
        sheetName: submission.sheetName,
        formRevision: submission.formRevision,
        measurements: submission.measurements,
        measurementCount: submission.measurements.length,
        status: "queued",
        attemptCount: 0,
        retryable: false,
        acceptedAt,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      };
      transaction.create(submissionRef, document);
      transaction.set(rateRef, {
        count: currentCount + 1,
        window: minuteBucket(acceptedDate),
        updatedAt: serverTimestamp()
      });
      return { created: true, status: publicSubmissionStatus(submission.idempotencyKey, document) };
    });
  }

  async function getStatus(idempotencyKey) {
    const id = requiredString(idempotencyKey, "idempotencyKey", 80);
    if (!IDEMPOTENCY_KEY_PATTERN.test(id)) throw new Error("INVALID_IDEMPOTENCY_KEY");
    const snapshot = await firestore.collection("measurementSubmissions").doc(id).get();
    if (!snapshot.exists) throw new Error("SUBMISSION_NOT_FOUND");
    return publicSubmissionStatus(snapshot.id, snapshot.data());
  }

  async function setGate(enabled) {
    await firestore.collection("systemConfig").doc("submissions").set({
      enabled: enabled === true,
      updatedAt: serverTimestamp()
    }, { merge: true });
    return { enabled: enabled === true };
  }

  return { getStatus, setGate, submit };
}

function quoteSheetName(sheetName) {
  return `'${String(sheetName).replaceAll("'", "''")}'`;
}

function formatMeasurementValue(value, decimalPlacesRaw) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const numeric = Number(text);
  if (!Number.isFinite(numeric)) throw new Error("INVALID_MEASUREMENT_VALUE");
  const decimalPlaces = Number.parseInt(String(decimalPlacesRaw ?? ""), 10);
  return Number.isInteger(decimalPlaces) && decimalPlaces >= 0 && decimalPlaces <= 20
    ? numeric.toFixed(decimalPlaces)
    : text;
}

function buildSheetsBatchUpdate({ sheetName, measurements, sheetRows, formListRows, revision }) {
  if (!Array.isArray(sheetRows) || !sheetRows.length) throw new Error("EMPTY_PRODUCTION_SHEET");
  if (!Array.isArray(formListRows) || !formListRows.length) throw new Error("EMPTY_FORM_LIST_SHEET");
  const rowsByIdentity = new Map();
  sheetRows.forEach((row, index) => {
    const key = measurementIdentity({ uniqueId: row[0], location: row[1], item: row[2], unit: row[4] });
    if (rowsByIdentity.has(key)) throw new Error("DUPLICATE_PRODUCTION_SHEET_ROW");
    rowsByIdentity.set(key, { row, sheetRowNumber: index + 2 });
  });
  const usedRows = new Set();
  const quotedSheetName = quoteSheetName(sheetName);
  const data = measurements.map((measurement) => {
    const match = rowsByIdentity.get(measurementIdentity(measurement));
    if (!match) throw new Error("MEASUREMENT_ROW_NOT_FOUND");
    if (usedRows.has(match.sheetRowNumber)) throw new Error("DUPLICATE_MEASUREMENT_TARGET");
    usedRows.add(match.sheetRowNumber);
    if (comparisonText(match.row[1]) !== comparisonText(measurement.location) ||
        comparisonText(match.row[2]) !== comparisonText(measurement.item) ||
        comparisonText(match.row[4]) !== comparisonText(measurement.unit)) {
      throw new Error("MEASUREMENT_ROW_IDENTITY_MISMATCH");
    }
    return {
      range: `${quotedSheetName}!F${match.sheetRowNumber}`,
      majorDimension: "ROWS",
      values: [[formatMeasurementValue(measurement.value, match.row[7])]]
    };
  });
  const formListIndex = formListRows.findIndex((row) =>
    String(row[0] ?? "").normalize("NFC") === sheetName && row[1] === PRODUCTION_SPREADSHEET_ID);
  if (formListIndex < 0) throw new Error("FORM_LIST_ENTRY_NOT_FOUND");
  data.push({
    range: `'FormList'!C${formListIndex + 2}`,
    majorDimension: "ROWS",
    values: [[revision]]
  });
  return { valueInputOption: "RAW", includeValuesInResponse: false, data };
}

function createSheetsGateway({ authClient, spreadsheetId = PRODUCTION_SPREADSHEET_ID }) {
  if (!authClient || typeof authClient.request !== "function") throw new Error("AUTH_CLIENT_REQUIRED");
  if (spreadsheetId !== PRODUCTION_SPREADSHEET_ID) throw new Error("PRODUCTION_SPREADSHEET_ID_MISMATCH");
  async function readRange(range) {
    const response = await authClient.request({
      method: "GET",
      url: `${SHEETS_API_BASE}/${spreadsheetId}/values/${encodeURIComponent(range)}`,
      params: {
        majorDimension: "ROWS",
        valueRenderOption: "FORMATTED_VALUE",
        dateTimeRenderOption: "FORMATTED_STRING"
      }
    });
    return response.data.values || [];
  }
  async function syncMeasurements({ sheetName, measurements, revision }) {
    const [sheetRows, formListRows] = await Promise.all([
      readRange(`${quoteSheetName(sheetName)}!C2:J`),
      readRange("'FormList'!A2:C")
    ]);
    const requestBody = buildSheetsBatchUpdate({ sheetName, measurements, sheetRows, formListRows, revision });
    const response = await authClient.request({
      method: "POST",
      url: `${SHEETS_API_BASE}/${spreadsheetId}/values:batchUpdate`,
      data: requestBody
    });
    if (Number(response.data.totalUpdatedCells) !== requestBody.data.length) {
      throw new Error(`SHEETS_UPDATED_CELL_COUNT_MISMATCH:${response.data.totalUpdatedCells || 0}:${requestBody.data.length}`);
    }
    return { updatedCellCount: response.data.totalUpdatedCells, updatedRangeCount: requestBody.data.length };
  }
  return { readRange, syncMeasurements };
}

function errorCode(error) {
  return String(error?.message || "UNKNOWN_SYNC_ERROR").slice(0, 200);
}

function isRetryableError(error) {
  return /(?:429|5\d\d|ECONN|ETIMEDOUT|TIMEOUT|ABORT)/i.test(errorCode(error));
}

function createSynchronizer({ firestore, serverTimestamp, sheetsGateway, publisher, now = () => new Date(), logger = console }) {
  async function acquire(submissionRef, owner) {
    return firestore.runTransaction(async (transaction) => {
      const snapshot = await transaction.get(submissionRef);
      if (!snapshot.exists) throw new Error("SUBMISSION_NOT_FOUND");
      const data = snapshot.data();
      if (data.status === "synced") return { skip: true, status: publicSubmissionStatus(snapshot.id, data) };
      const leaseExpiresAt = Date.parse(data.leaseExpiresAt || "");
      if (data.status === "syncing" && data.syncOwner !== owner &&
          Number.isFinite(leaseExpiresAt) && leaseExpiresAt > now().getTime()) {
        throw new Error("SUBMISSION_SYNC_IN_PROGRESS");
      }
      const startedAt = now();
      transaction.update(submissionRef, {
        status: "syncing",
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
      const currentRevision = normalizeRevision(published.document.sourceRevision);
      const validationRevision = currentRevision === submission.acceptedAt
        ? submission.acceptedAt
        : submission.formRevision;
      validateAgainstPublishedForm(
        { ...submission, formRevision: validationRevision },
        published.document,
        published.rows
      );
      const sheetResult = await sheetsGateway.syncMeasurements({
        sheetName: submission.sheetName,
        measurements: submission.measurements,
        revision: submission.acceptedAt
      });
      const cacheResult = await publisher.publishSubmissionSnapshot({
        formDocument: published.document,
        rows: published.rows,
        measurements: submission.measurements,
        sourceRevision: submission.acceptedAt
      });
      await submissionRef.update({
        status: "synced",
        sourceRevisionAfterSync: submission.acceptedAt,
        updatedCellCount: sheetResult.updatedCellCount,
        cacheFormStatus: cacheResult.form.status,
        cacheListStatus: cacheResult.list.status,
        retryable: false,
        errorCode: null,
        syncOwner: null,
        leaseExpiresAt: null,
        syncedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      const completed = await submissionRef.get();
      logger.info("Production measurement submission synced", {
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
      logger.error("Production measurement submission sync failed", {
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
  MAX_REQUEST_BYTES,
  buildSheetsBatchUpdate,
  createSheetsGateway,
  createSubmissionService,
  createSynchronizer,
  measurementIdentity,
  normalizeSubmissionRequest,
  publicSubmissionStatus,
  quoteSheetName,
  validateAgainstPublishedForm
};
