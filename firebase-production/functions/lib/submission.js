"use strict";

const { hashCanonical } = require("./publisher");

const SUBMISSION_SCHEMA_VERSION = 1;
const MAX_MEASUREMENTS = 250;
const MAX_REQUEST_BYTES = 128 * 1024;
const IDEMPOTENCY_KEY_PATTERN = /^[A-Za-z0-9_-]{20,80}$/;
const FORM_KEY_PATTERN = /^f_[a-f0-9]{32}$/;
const NUMERIC_VALUE_PATTERN = /^-?(?:\d+|\d*\.\d+)$/;
const RATE_LIMIT_PER_MINUTE = 12;

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
  const text = requiredString(value, "formRevision", 100);
  const timestamp = Date.parse(text);
  if (!Number.isFinite(timestamp)) throw new Error("INVALID_FORM_REVISION");
  return new Date(timestamp).toISOString();
}

function measurementIdentity(measurement) {
  if (measurement.uniqueId) return `id:${measurement.uniqueId}`;
  return `pair:${measurement.location.toLocaleLowerCase("ko-KR")}|${measurement.item.toLocaleLowerCase("ko-KR")}`;
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
  return {
    ...normalized,
    requestHash: hashCanonical(normalized)
  };
}

function normalizePublishedRow(entry) {
  return {
    uniqueId: optionalString(entry?.uniqueId, 120).trim(),
    location: requiredString(entry?.location, "publishedLocation", 200),
    item: requiredString(entry?.item, "publishedItem", 200),
    unit: optionalString(entry?.unit, 80).trim()
  };
}

function validateAgainstPublishedForm(submission, formDocument, rows) {
  if (!formDocument || formDocument.formKey !== submission.formKey ||
      formDocument.sheetName !== submission.sheetName) {
    throw new Error("PUBLISHED_FORM_IDENTITY_MISMATCH");
  }
  const sourceTimestamp = Date.parse(formDocument.sourceRevision);
  if (!Number.isFinite(sourceTimestamp) || new Date(sourceTimestamp).toISOString() !== submission.formRevision) {
    throw new Error("STALE_SUBMISSION_FORM_REVISION");
  }
  if (!Array.isArray(rows) || rows.length !== submission.measurements.length ||
      formDocument.rowCount !== rows.length) {
    throw new Error("SUBMISSION_ROW_COUNT_MISMATCH");
  }

  const publishedByIdentity = new Map();
  rows.map(normalizePublishedRow).forEach((published) => {
    const identity = measurementIdentity(published);
    if (publishedByIdentity.has(identity)) throw new Error("DUPLICATE_PUBLISHED_ROW");
    publishedByIdentity.set(identity, published);
  });
  for (const measurement of submission.measurements) {
    const published = publishedByIdentity.get(measurementIdentity(measurement));
    if (!published || published.uniqueId !== measurement.uniqueId ||
        published.location !== measurement.location ||
        published.item !== measurement.item ||
        published.unit !== measurement.unit) {
      throw new Error("SUBMISSION_ROW_IDENTITY_MISMATCH");
    }
  }
  return true;
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
    cachedAt: timestampToIso(data.cachedAt),
    syncStartedAt: timestampToIso(data.syncStartedAt),
    syncedAt: timestampToIso(data.syncedAt),
    failedAt: timestampToIso(data.failedAt),
    attemptCount: Number(data.attemptCount || 0),
    retryable: Boolean(data.retryable),
    errorCode: data.errorCode || null,
    dailyCacheDate: data.dailyCacheDate || null,
    sourceRevisionAfterCache: data.sourceRevisionAfterCache || null,
    sourceRevisionAfterSync: data.sourceRevisionAfterSync || null,
    updatedCellCount: Number(data.updatedCellCount || 0)
  };
}

function minuteBucket(date) {
  return date.toISOString().slice(0, 16).replace(/[-:T]/g, "");
}

function createSubmissionService({ firestore, serverTimestamp, now = () => new Date() }) {
  if (!firestore || typeof firestore.runTransaction !== "function") throw new Error("FIRESTORE_REQUIRED");
  if (typeof serverTimestamp !== "function") throw new Error("SERVER_TIMESTAMP_REQUIRED");

  async function submit(payload) {
    const submission = normalizeSubmissionRequest(payload);
    const submissionRef = firestore.collection("measurementSubmissions").doc(submission.idempotencyKey);
    const existingSnapshot = await submissionRef.get();
    if (existingSnapshot.exists) {
      const existingData = existingSnapshot.data();
      if (existingData.requestHash !== submission.requestHash) throw new Error("IDEMPOTENCY_KEY_CONFLICT");
      return { created: false, status: publicSubmissionStatus(existingSnapshot.id, existingData) };
    }

    const published = await loadPublishedForm(firestore, submission.formKey);
    validateAgainstPublishedForm(submission, published.document, published.rows);

    const gateRef = firestore.collection("systemConfig").doc("submissions");
    const acceptedDate = now();
    const rateRef = firestore.collection("submissionRateLimits").doc(`minute_${minuteBucket(acceptedDate)}`);

    const result = await firestore.runTransaction(async (transaction) => {
      const existing = await transaction.get(submissionRef);
      if (existing.exists) {
        const data = existing.data();
        if (data.requestHash !== submission.requestHash) throw new Error("IDEMPOTENCY_KEY_CONFLICT");
        return { created: false, status: publicSubmissionStatus(existing.id, data) };
      }

      const gate = await transaction.get(gateRef);
      if (!gate.exists || gate.get("enabled") !== true) throw new Error("SUBMISSIONS_DISABLED");
      const allowedFormKeys = gate.get("allowedFormKeys");
      const allowedSheetNames = gate.get("allowedSheetNames");
      if (Array.isArray(allowedFormKeys) && allowedFormKeys.length > 0 &&
          !allowedFormKeys.includes(submission.formKey)) {
        throw new Error("SUBMISSION_FORM_NOT_ALLOWED");
      }
      if (Array.isArray(allowedSheetNames) && allowedSheetNames.length > 0 &&
          !allowedSheetNames.includes(submission.sheetName)) {
        throw new Error("SUBMISSION_FORM_NOT_ALLOWED");
      }
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
    return result;
  }

  async function getStatus(idempotencyKey) {
    const id = requiredString(idempotencyKey, "idempotencyKey", 80);
    if (!IDEMPOTENCY_KEY_PATTERN.test(id)) throw new Error("INVALID_IDEMPOTENCY_KEY");
    const snapshot = await firestore.collection("measurementSubmissions").doc(id).get();
    if (!snapshot.exists) throw new Error("SUBMISSION_NOT_FOUND");
    return publicSubmissionStatus(snapshot.id, snapshot.data());
  }

  async function setGate(enabled, { allowedFormKeys = [], allowedSheetNames = [] } = {}) {
    if (!Array.isArray(allowedFormKeys) || !Array.isArray(allowedSheetNames)) {
      throw new Error("INVALID_SUBMISSION_ALLOWLIST");
    }
    const normalizedFormKeys = [...new Set(allowedFormKeys.map((value) =>
      requiredString(value, "allowedFormKey", 80)))];
    const normalizedSheetNames = [...new Set(allowedSheetNames.map((value) =>
      requiredString(value, "allowedSheetName", 120)))];
    if (normalizedFormKeys.some((value) => !FORM_KEY_PATTERN.test(value))) {
      throw new Error("INVALID_ALLOWED_FORM_KEY");
    }
    await firestore.collection("systemConfig").doc("submissions").set({
      enabled: enabled === true,
      allowedFormKeys: normalizedFormKeys,
      allowedSheetNames: normalizedSheetNames,
      updatedAt: serverTimestamp()
    }, { merge: true });
    return {
      enabled: enabled === true,
      allowedFormKeys: normalizedFormKeys,
      allowedSheetNames: normalizedSheetNames
    };
  }

  return { getStatus, setGate, submit };
}

module.exports = {
  FORM_KEY_PATTERN,
  IDEMPOTENCY_KEY_PATTERN,
  MAX_MEASUREMENTS,
  MAX_REQUEST_BYTES,
  RATE_LIMIT_PER_MINUTE,
  SUBMISSION_SCHEMA_VERSION,
  createSubmissionService,
  loadPublishedForm,
  measurementIdentity,
  normalizeSubmissionRequest,
  publicSubmissionStatus,
  validateAgainstPublishedForm
};
