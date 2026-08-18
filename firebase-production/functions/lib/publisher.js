"use strict";

const { createHash } = require("node:crypto");

const PRODUCTION_PROJECT_ID = "btcwebapp-551bd";
const PRODUCTION_SPREADSHEET_ID = "19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s";
const PRODUCTION_GAS_API_URL = "https://script.google.com/macros/s/AKfycbzuWS4Q5kTzDRH4IBpeXBa69KngElRdArtTCzTV0NDQsB3y4oABBIzrTLuPOZH5KOPP/exec";
const SCHEMA_VERSION = 1;
const GAS_TIMEOUT_MS = 45_000;
const MAX_INLINE_DOCUMENT_BYTES = 850 * 1024;
const MAX_CHUNK_DOCUMENT_BYTES = 700 * 1024;

function stableStringify(value) {
  if (Array.isArray(value)) {
    return `[${value.map((entry) => stableStringify(entry)).join(",")}]`;
  }
  if (value && typeof value === "object") {
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value);
}

function hashCanonical(value) {
  return createHash("sha256").update(stableStringify(value), "utf8").digest("hex");
}

function serializedBytes(value) {
  return Buffer.byteLength(stableStringify(value), "utf8");
}

function isOlderRevision(candidate, current) {
  const candidateTime = Date.parse(candidate);
  const currentTime = Date.parse(current);
  return Number.isFinite(candidateTime) && Number.isFinite(currentTime) && candidateTime < currentTime;
}

function formListRegresses(candidate, current) {
  if (!Array.isArray(candidate?.items) || !Array.isArray(current?.items)) return false;
  const currentRevisions = new Map(current.items.map((item) => [item.formKey, item.lastModifiedDate]));
  return candidate.items.some((item) => currentRevisions.has(item.formKey) &&
    isOlderRevision(item.lastModifiedDate, currentRevisions.get(item.formKey)));
}

function requireNonEmptyString(value, fieldName) {
  const normalized = String(value ?? "").trim().normalize("NFC");
  if (!normalized) {
    throw new Error(`INVALID_${fieldName.toUpperCase()}`);
  }
  return normalized;
}

function normalizeOptionalString(value) {
  if (value === null || value === undefined) return null;
  return String(value).normalize("NFC");
}

function normalizeRevision(value) {
  const text = requireNonEmptyString(value, "lastModifiedDate");
  const timestamp = Date.parse(text);
  if (!Number.isFinite(timestamp)) {
    throw new Error("INVALID_LASTMODIFIEDDATE");
  }
  return new Date(timestamp).toISOString();
}

function formKeyForSheet(sheetName) {
  const normalized = requireNonEmptyString(sheetName, "sheetName");
  return `f_${createHash("sha256").update(normalized, "utf8").digest("hex").slice(0, 32)}`;
}

function normalizeFormList(payload) {
  if (!Array.isArray(payload)) throw new Error("INVALID_FORM_LIST_RESPONSE");
  if (!payload.length) throw new Error("EMPTY_FORM_LIST_REJECTED");

  const seenSheetNames = new Set();
  const seenFormKeys = new Set();
  const items = payload.map((entry) => {
    if (!entry || typeof entry !== "object") throw new Error("INVALID_FORM_LIST_ENTRY");
    if (entry.spreadsheetId !== PRODUCTION_SPREADSHEET_ID) {
      throw new Error("PRODUCTION_SPREADSHEET_ID_MISMATCH");
    }

    const sheetName = requireNonEmptyString(entry.sheetName, "sheetName");
    const formKey = formKeyForSheet(sheetName);
    if (seenSheetNames.has(sheetName) || seenFormKeys.has(formKey)) {
      throw new Error("DUPLICATE_FORM_ENTRY");
    }
    seenSheetNames.add(sheetName);
    seenFormKeys.add(formKey);

    return {
      formKey,
      sheetName,
      displayName: sheetName,
      lastModifiedDate: normalizeRevision(entry.lastModifiedDate)
    };
  });

  const sourceRevision = items.length
    ? items.map((item) => item.lastModifiedDate).sort().at(-1)
    : new Date(0).toISOString();
  const contentHash = hashCanonical({ schemaVersion: SCHEMA_VERSION, sourceRevision, items });

  return {
    schemaVersion: SCHEMA_VERSION,
    sourceRevision,
    contentHash,
    itemCount: items.length,
    items
  };
}

function normalizeValidation(validation) {
  if (validation === null || validation === undefined) return null;
  if (typeof validation !== "object") throw new Error("INVALID_VALIDATION");
  return {
    minValue: normalizeOptionalString(validation.minValue),
    maxValue: normalizeOptionalString(validation.maxValue)
  };
}

function normalizeRecentInfo(recentInfo) {
  if (recentInfo === null || recentInfo === undefined) return null;
  if (typeof recentInfo !== "object") throw new Error("INVALID_RECENT_INFO");
  return {
    value: normalizeOptionalString(recentInfo.value),
    date: normalizeOptionalString(recentInfo.date)
  };
}

function normalizeFormRows(payload) {
  if (!Array.isArray(payload)) throw new Error("INVALID_FORM_RESPONSE");
  if (!payload.length) throw new Error("EMPTY_FORM_RESPONSE_REJECTED");

  const seenUniqueIds = new Set();
  return payload.map((entry) => {
    if (!entry || typeof entry !== "object") throw new Error("INVALID_FORM_ROW");
    const uniqueId = normalizeOptionalString(entry.uniqueId) ?? "";
    if (uniqueId && seenUniqueIds.has(uniqueId)) throw new Error("DUPLICATE_UNIQUE_ID");
    if (uniqueId) seenUniqueIds.add(uniqueId);

    return {
      uniqueId,
      location: normalizeOptionalString(entry.location) ?? "",
      item: normalizeOptionalString(entry.item) ?? "",
      value: normalizeOptionalString(entry.value) ?? "",
      unit: normalizeOptionalString(entry.unit) ?? "",
      validation: normalizeValidation(entry.validation),
      recentInfo: normalizeRecentInfo(entry.recentInfo)
    };
  });
}

function buildFormDocument(formListItem, payload) {
  const rows = normalizeFormRows(payload);
  const base = {
    schemaVersion: SCHEMA_VERSION,
    formKey: formListItem.formKey,
    sheetName: formListItem.sheetName,
    sourceRevision: formListItem.lastModifiedDate,
    rowCount: rows.length,
    rows
  };
  return {
    ...base,
    contentHash: hashCanonical(base)
  };
}

function buildStoragePlan(formDocument) {
  if (serializedBytes(formDocument) <= MAX_INLINE_DOCUMENT_BYTES) {
    return {
      root: { ...formDocument, storageMode: "inline", chunkCount: 0 },
      chunks: []
    };
  }

  const chunks = [];
  let currentRows = [];
  for (const row of formDocument.rows) {
    const candidate = [...currentRows, row];
    if (serializedBytes({ rows: candidate }) > MAX_CHUNK_DOCUMENT_BYTES) {
      if (!currentRows.length) throw new Error("FORM_ROW_TOO_LARGE");
      chunks.push(currentRows);
      currentRows = [row];
    } else {
      currentRows = candidate;
    }
  }
  if (currentRows.length) chunks.push(currentRows);

  const chunkDocuments = chunks.map((rows, index) => ({
    schemaVersion: SCHEMA_VERSION,
    formKey: formDocument.formKey,
    contentHash: formDocument.contentHash,
    index,
    rowCount: rows.length,
    rows
  }));

  const { rows, ...manifest } = formDocument;
  return {
    root: {
      ...manifest,
      storageMode: "chunked",
      chunkCount: chunkDocuments.length
    },
    chunks: chunkDocuments
  };
}

async function mapWithConcurrency(values, limit, mapper) {
  const results = new Array(values.length);
  let nextIndex = 0;
  async function worker() {
    while (nextIndex < values.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await mapper(values[index], index);
    }
  }
  const workers = Array.from({ length: Math.min(limit, values.length) }, () => worker());
  await Promise.all(workers);
  return results;
}

function seoulDateKey(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return null;
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Seoul",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function dailyMeasurementCacheId(formKey, value) {
  return `${requireNonEmptyString(formKey, "formKey")}_${seoulDateKey(value)}`;
}

function createPublisher({ firestore, fetchImpl = globalThis.fetch, serverTimestamp, logger = console }) {
  if (!firestore || typeof firestore.collection !== "function") throw new Error("FIRESTORE_REQUIRED");
  if (typeof fetchImpl !== "function") throw new Error("FETCH_REQUIRED");
  if (typeof serverTimestamp !== "function") throw new Error("SERVER_TIMESTAMP_REQUIRED");

  async function fetchGasJson(action, parameters = {}) {
    const url = new URL(PRODUCTION_GAS_API_URL);
    url.searchParams.set("action", action);
    Object.entries(parameters).forEach(([key, value]) => url.searchParams.set(key, String(value)));

    const response = await fetchImpl(url, {
      method: "GET",
      redirect: "follow",
      headers: { Accept: "application/json" },
      signal: AbortSignal.timeout(GAS_TIMEOUT_MS)
    });
    if (!response.ok) throw new Error(`PRODUCTION_GAS_HTTP_${response.status}`);
    const payload = await response.json();
    if (payload && !Array.isArray(payload) && payload.error) {
      throw new Error(`PRODUCTION_GAS_ERROR_${String(payload.error).slice(0, 200)}`);
    }
    return payload;
  }

  async function readSourceFormList() {
    return normalizeFormList(await fetchGasJson("getFormList"));
  }

  async function deleteChunks(rootRef, keepIds = new Set()) {
    const snapshot = await rootRef.collection("chunks").get();
    await Promise.all(snapshot.docs
      .filter((document) => !keepIds.has(document.id))
      .map((document) => document.ref.delete()));
  }

  async function persistFormList(formList, force = false) {
    const reference = firestore.collection("publicCache").doc("formList");
    const existing = await reference.get();
    if (existing.exists && formListRegresses(formList, existing.data())) {
      return {
        status: "stale_skipped",
        contentHash: existing.get("contentHash"),
        itemCount: existing.get("itemCount")
      };
    }
    if (!force && existing.exists && existing.get("contentHash") === formList.contentHash) {
      return { status: "unchanged", contentHash: formList.contentHash, itemCount: formList.itemCount };
    }
    await reference.set({ ...formList, publishedAt: serverTimestamp() });
    return { status: "published", contentHash: formList.contentHash, itemCount: formList.itemCount };
  }

  async function persistForm(formDocument, force = false) {
    const rootRef = firestore.collection("publicForms").doc(formDocument.formKey);
    const existing = await rootRef.get();
    if (existing.exists && isOlderRevision(formDocument.sourceRevision, existing.get("sourceRevision"))) {
      return {
        formKey: formDocument.formKey,
        sheetName: formDocument.sheetName,
        status: "stale_skipped",
        rowCount: existing.get("rowCount"),
        contentHash: existing.get("contentHash")
      };
    }
    if (!force && existing.exists && existing.get("contentHash") === formDocument.contentHash) {
      return {
        formKey: formDocument.formKey,
        sheetName: formDocument.sheetName,
        status: "unchanged",
        rowCount: formDocument.rowCount,
        contentHash: formDocument.contentHash
      };
    }

    const plan = buildStoragePlan(formDocument);
    const keepChunkIds = new Set();
    for (const chunk of plan.chunks) {
      const chunkId = String(chunk.index).padStart(4, "0");
      keepChunkIds.add(chunkId);
      await rootRef.collection("chunks").doc(chunkId).set(chunk);
    }
    await rootRef.set({ ...plan.root, publishedAt: serverTimestamp() });
    await deleteChunks(rootRef, keepChunkIds);

    return {
      formKey: formDocument.formKey,
      sheetName: formDocument.sheetName,
      status: "published",
      storageMode: plan.root.storageMode,
      chunkCount: plan.root.chunkCount,
      rowCount: formDocument.rowCount,
      contentHash: formDocument.contentHash
    };
  }

  async function publishFormList({ force = false } = {}) {
    const formList = await readSourceFormList();
    return persistFormList(formList, force);
  }

  async function publishForm({ sheetName, force = false } = {}) {
    const requestedSheetName = requireNonEmptyString(sheetName, "sheetName");
    const formList = await readSourceFormList();
    const formListItem = formList.items.find((item) => item.sheetName === requestedSheetName);
    if (!formListItem) throw new Error("FORM_NOT_FOUND_IN_PRODUCTION_FORM_LIST");
    const payload = await fetchGasJson("getFormDataForWeb", { sheetName: requestedSheetName });
    return persistForm(buildFormDocument(formListItem, payload), force);
  }

  async function publishDailyMeasurementCache({ formDocument, rows, measurements, sourceRevision }) {
    if (!formDocument || !Array.isArray(rows) || !Array.isArray(measurements) ||
        rows.length !== measurements.length || rows.length !== formDocument.rowCount) {
      throw new Error("INVALID_DAILY_MEASUREMENT_CACHE");
    }
    const normalizedRevision = normalizeRevision(sourceRevision);
    const valuesByIdentity = new Map(measurements.map((measurement) => [
      measurement.uniqueId
        ? `id:${measurement.uniqueId}`
        : `pair:${measurement.location.toLocaleLowerCase("ko-KR")}|${measurement.item.toLocaleLowerCase("ko-KR")}`,
      measurement
    ]));
    const updatedRows = rows.map((row) => {
      const identity = row.uniqueId
        ? `id:${row.uniqueId}`
        : `pair:${row.location.toLocaleLowerCase("ko-KR")}|${row.item.toLocaleLowerCase("ko-KR")}`;
      const measurement = valuesByIdentity.get(identity);
      if (!measurement || row.uniqueId !== measurement.uniqueId ||
          row.location !== measurement.location || row.item !== measurement.item ||
          row.unit !== measurement.unit) {
        throw new Error("DAILY_MEASUREMENT_CACHE_IDENTITY_MISMATCH");
      }
      return {
        uniqueId: row.uniqueId,
        location: row.location,
        item: row.item,
        value: normalizeOptionalString(measurement.value) ?? "",
        unit: row.unit
      };
    });
    const cacheDate = seoulDateKey(normalizedRevision);
    const cacheId = dailyMeasurementCacheId(formDocument.formKey, normalizedRevision);
    const reference = firestore.collection("dailyMeasurementCaches").doc(cacheId);
    const existing = await reference.get();
    if (existing.exists && isOlderRevision(normalizedRevision, existing.get("sourceRevision"))) {
      return {
        status: "stale_skipped",
        cacheId,
        cacheDate,
        sourceRevision: existing.get("sourceRevision")
      };
    }
    const base = {
      schemaVersion: SCHEMA_VERSION,
      formKey: formDocument.formKey,
      sheetName: formDocument.sheetName,
      cacheDate,
      sourceRevision: normalizedRevision,
      measurementCount: updatedRows.length,
      measurements: updatedRows
    };
    await reference.set({
      ...base,
      contentHash: hashCanonical(base),
      cachedAt: serverTimestamp()
    });
    return {
      status: "published",
      cacheId,
      cacheDate,
      sourceRevision: normalizedRevision,
      measurementCount: updatedRows.length
    };
  }

  async function deleteOrphanedForms(validFormKeys) {
    const snapshot = await firestore.collection("publicForms").get();
    const orphans = snapshot.docs.filter((document) => !validFormKeys.has(document.id));
    for (const orphan of orphans) {
      await deleteChunks(orphan.ref);
      await orphan.ref.delete();
    }
    return orphans.length;
  }

  async function deleteExpiredDailyMeasurementCaches({ currentDate = seoulDateKey(new Date()) } = {}) {
    const snapshot = await firestore.collection("dailyMeasurementCaches").get();
    const expired = snapshot.docs.filter((document) => {
      const cacheDate = String(document.get("cacheDate") || "");
      return /^\d{4}-\d{2}-\d{2}$/.test(cacheDate) && cacheDate < currentDate;
    });
    await Promise.all(expired.map((document) => document.ref.delete()));
    return { currentDate, deletedCount: expired.length };
  }

  async function publishAllChangedForms({ force = false } = {}) {
    const startedAt = Date.now();
    const formList = await readSourceFormList();
    const forms = await mapWithConcurrency(formList.items, 2, async (formListItem) => {
      const payload = await fetchGasJson("getFormDataForWeb", { sheetName: formListItem.sheetName });
      return persistForm(buildFormDocument(formListItem, payload), force);
    });
    const list = await persistFormList(formList, force);
    const removedOrphanCount = await deleteOrphanedForms(new Set(formList.items.map((item) => item.formKey)));
    const result = {
      list,
      forms,
      removedOrphanCount,
      durationMs: Date.now() - startedAt
    };
    logger.info("Firestore test cache publish completed", {
      itemCount: formList.itemCount,
      publishedFormCount: forms.filter((form) => form.status === "published").length,
      removedOrphanCount,
      durationMs: result.durationMs
    });
    return result;
  }

  return {
    publishFormList,
    publishForm,
    publishAllChangedForms,
    publishDailyMeasurementCache,
    deleteExpiredDailyMeasurementCaches
  };
}

module.exports = {
  GAS_TIMEOUT_MS,
  MAX_CHUNK_DOCUMENT_BYTES,
  MAX_INLINE_DOCUMENT_BYTES,
  SCHEMA_VERSION,
  PRODUCTION_GAS_API_URL,
  PRODUCTION_PROJECT_ID,
  PRODUCTION_SPREADSHEET_ID,
  buildFormDocument,
  buildStoragePlan,
  createPublisher,
  dailyMeasurementCacheId,
  formKeyForSheet,
  hashCanonical,
  isOlderRevision,
  normalizeFormList,
  normalizeFormRows,
  serializedBytes,
  stableStringify
};
