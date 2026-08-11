"use strict";

const { PRODUCTION_SPREADSHEET_ID } = require("./publisher");

const SHEETS_API_BASE = "https://sheets.googleapis.com/v4/spreadsheets";

function quoteSheetName(sheetName) {
  return `'${String(sheetName).replaceAll("'", "''")}'`;
}

function comparisonText(value) {
  return String(value ?? "").trim().normalize("NFC").toLocaleLowerCase("ko-KR");
}

function measurementKey(entry) {
  const uniqueId = String(entry.uniqueId ?? "").trim().normalize("NFC");
  if (uniqueId) return `id:${uniqueId}`;
  return `pair:${comparisonText(entry.location)}|${comparisonText(entry.item)}`;
}

function formatMeasurementValue(value, decimalPlacesRaw) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const numeric = Number(text);
  if (!Number.isFinite(numeric)) throw new Error("INVALID_MEASUREMENT_VALUE");
  const decimalPlaces = Number.parseInt(String(decimalPlacesRaw ?? ""), 10);
  if (Number.isInteger(decimalPlaces) && decimalPlaces >= 0 && decimalPlaces <= 20) {
    return numeric.toFixed(decimalPlaces);
  }
  return text;
}

function buildSheetsBatchUpdate({ sheetName, measurements, sheetRows, formListRows, revision }) {
  if (!Array.isArray(sheetRows) || !sheetRows.length) throw new Error("EMPTY_PRODUCTION_SHEET");
  if (!Array.isArray(formListRows) || !formListRows.length) throw new Error("EMPTY_FORM_LIST_SHEET");

  const rowsByIdentity = new Map();
  sheetRows.forEach((row, index) => {
    const entry = {
      uniqueId: row[0],
      location: row[1],
      item: row[2],
      unit: row[4]
    };
    const key = measurementKey(entry);
    if (rowsByIdentity.has(key)) throw new Error("DUPLICATE_PRODUCTION_SHEET_ROW");
    rowsByIdentity.set(key, { row, sheetRowNumber: index + 2 });
  });

  const usedRows = new Set();
  const quotedSheetName = quoteSheetName(sheetName);
  const data = measurements.map((measurement) => {
    const match = rowsByIdentity.get(measurementKey(measurement));
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

  return {
    valueInputOption: "RAW",
    includeValuesInResponse: false,
    data
  };
}

function createSheetsGateway({ authClient, spreadsheetId = PRODUCTION_SPREADSHEET_ID }) {
  if (!authClient || typeof authClient.request !== "function") throw new Error("AUTH_CLIENT_REQUIRED");
  if (spreadsheetId !== PRODUCTION_SPREADSHEET_ID) throw new Error("PRODUCTION_SPREADSHEET_ID_MISMATCH");

  async function readRange(range) {
    const encodedRange = encodeURIComponent(range);
    const response = await authClient.request({
      method: "GET",
      url: `${SHEETS_API_BASE}/${spreadsheetId}/values/${encodedRange}`,
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
    const requestBody = buildSheetsBatchUpdate({
      sheetName,
      measurements,
      sheetRows,
      formListRows,
      revision
    });
    const response = await authClient.request({
      method: "POST",
      url: `${SHEETS_API_BASE}/${spreadsheetId}/values:batchUpdate`,
      data: requestBody
    });
    const expectedCells = requestBody.data.length;
    if (Number(response.data.totalUpdatedCells) !== expectedCells) {
      throw new Error(`SHEETS_UPDATED_CELL_COUNT_MISMATCH:${response.data.totalUpdatedCells || 0}:${expectedCells}`);
    }
    return {
      updatedCellCount: response.data.totalUpdatedCells,
      updatedRangeCount: requestBody.data.length
    };
  }

  return { readRange, syncMeasurements };
}

module.exports = {
  SHEETS_API_BASE,
  buildSheetsBatchUpdate,
  comparisonText,
  createSheetsGateway,
  formatMeasurementValue,
  measurementKey,
  quoteSheetName
};
