"use strict";

const {
  PRODUCTION_GAS_API_URL,
  buildFormDocument,
  normalizeFormList,
  stableStringify
} = require("../firebase-production/functions/lib/publisher");

const PROJECT_ID = "btcwebapp-551bd";
const API_KEY = "AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM";
const FIRESTORE_BASE = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}` +
  "/databases/(default)/documents";

function decodeFirestoreValue(value) {
  if ("nullValue" in value) return null;
  if ("stringValue" in value) return value.stringValue;
  if ("integerValue" in value) return Number(value.integerValue);
  if ("doubleValue" in value) return Number(value.doubleValue);
  if ("booleanValue" in value) return Boolean(value.booleanValue);
  if ("timestampValue" in value) return value.timestampValue;
  if ("arrayValue" in value) return (value.arrayValue.values || []).map(decodeFirestoreValue);
  if ("mapValue" in value) return decodeFirestoreFields(value.mapValue.fields || {});
  throw new Error(`UNSUPPORTED_FIRESTORE_VALUE:${Object.keys(value).join(",")}`);
}

function decodeFirestoreFields(fields) {
  return Object.fromEntries(
    Object.entries(fields).map(([key, value]) => [key, decodeFirestoreValue(value)])
  );
}

async function fetchJson(url) {
  const response = await fetch(url, {
    headers: { Accept: "application/json" },
    signal: AbortSignal.timeout(60_000)
  });
  if (!response.ok) throw new Error(`HTTP_${response.status}:${url}`);
  return response.json();
}

async function fetchGas(action, parameters = {}) {
  const url = new URL(PRODUCTION_GAS_API_URL);
  url.searchParams.set("action", action);
  Object.entries(parameters).forEach(([key, value]) => url.searchParams.set(key, String(value)));
  return fetchJson(url);
}

async function fetchFirestoreDocument(path) {
  const url = `${FIRESTORE_BASE}/${path}?key=${encodeURIComponent(API_KEY)}`;
  const document = await fetchJson(url);
  return decodeFirestoreFields(document.fields || {});
}

async function fetchPublishedRows(formKey, document) {
  if (document.storageMode === "inline") return document.rows;
  const url = `${FIRESTORE_BASE}/publicForms/${formKey}/chunks` +
    `?key=${encodeURIComponent(API_KEY)}&pageSize=1000`;
  const response = await fetchJson(url);
  return (response.documents || [])
    .map((entry) => decodeFirestoreFields(entry.fields || {}))
    .sort((left, right) => left.index - right.index)
    .flatMap((chunk) => chunk.rows);
}

async function main() {
  const gasList = await fetchGas("getFormList");
  const expectedList = normalizeFormList(gasList);
  const actualList = await fetchFirestoreDocument("publicCache/formList");
  if (actualList.contentHash !== expectedList.contentHash ||
      actualList.itemCount !== expectedList.itemCount) {
    throw new Error("FORM_LIST_MISMATCH");
  }

  const formResults = [];
  for (const item of expectedList.items) {
    const gasRows = await fetchGas("getFormDataForWeb", { sheetName: item.sheetName });
    const expectedDocument = buildFormDocument(item, gasRows);
    const actualDocument = await fetchFirestoreDocument(`publicForms/${item.formKey}`);
    const actualRows = await fetchPublishedRows(item.formKey, actualDocument);
    if (actualDocument.contentHash !== expectedDocument.contentHash) {
      throw new Error(`FORM_HASH_MISMATCH:${item.sheetName}`);
    }
    if (stableStringify(actualRows) !== stableStringify(expectedDocument.rows)) {
      throw new Error(`FORM_ROWS_MISMATCH:${item.sheetName}`);
    }
    formResults.push({
      sheetName: item.sheetName,
      formKey: item.formKey,
      rowCount: actualRows.length,
      status: "matched"
    });
  }

  console.log(JSON.stringify({
    ok: true,
    list: { itemCount: actualList.itemCount, contentHash: actualList.contentHash },
    forms: formResults
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
