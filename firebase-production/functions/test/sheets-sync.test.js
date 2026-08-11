"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const { PRODUCTION_SPREADSHEET_ID } = require("../lib/publisher");
const {
  buildFormListDateFormatRequest,
  buildSheetsBatchUpdate,
  quoteSheetName,
  revisionToSheetsSerial
} = require("../lib/sheets-sync");

const input = {
  sheetName: "율리'24",
  revision: "2026-07-29T01:02:03.000Z",
  measurements: [
    { uniqueId: "u1", location: "24KV SIS", item: "전압", value: "24", unit: "kV" },
    { uniqueId: "u2", location: "24KV SIS", item: "전류", value: "", unit: "A" }
  ],
  sheetRows: [
    ["u1", "24KV SIS", "전압", "old", "kV", "", "", "2"],
    ["u2", "24KV SIS", "전류", "old", "A", "", "", "1"]
  ],
  formListRows: [
    ["율리'24", PRODUCTION_SPREADSHEET_ID, "old-revision"]
  ]
};

test("builds bounded F-column updates plus one deterministic FormList revision", () => {
  const update = buildSheetsBatchUpdate(input);
  assert.equal(update.valueInputOption, "RAW");
  assert.deepEqual(update.data, [
    { range: "'율리''24'!F2", majorDimension: "ROWS", values: [["24.00"]] },
    { range: "'율리''24'!F3", majorDimension: "ROWS", values: [[""]] },
    { range: "'FormList'!C2", majorDimension: "ROWS", values: [[revisionToSheetsSerial(input.revision)]] }
  ]);
});

test("stores FormList revisions as Seoul date values with the legacy display format", () => {
  const serial = revisionToSheetsSerial("2026-07-29T01:02:03.000Z");
  const representedMilliseconds = Math.round((serial - 25569) * 24 * 60 * 60 * 1000);
  assert.equal(new Date(representedMilliseconds).toISOString(), "2026-07-29T10:02:03.000Z");
  assert.deepEqual(buildFormListDateFormatRequest(0), {
    requests: [{
      repeatCell: {
        range: {
          sheetId: 0,
          startRowIndex: 1,
          endRowIndex: 2,
          startColumnIndex: 2,
          endColumnIndex: 3
        },
        cell: {
          userEnteredFormat: {
            numberFormat: { type: "DATE_TIME", pattern: "yyyy-mm-dd h:mm:ss" }
          }
        },
        fields: "userEnteredFormat.numberFormat"
      }
    }]
  });
});

test("rejects malformed FormList revisions", () => {
  assert.throws(() => revisionToSheetsSerial("not-a-date"), { message: "INVALID_SOURCE_REVISION" });
});

test("rebuilding the same update is exactly idempotent", () => {
  assert.deepEqual(buildSheetsBatchUpdate(input), buildSheetsBatchUpdate(structuredClone(input)));
});

test("rejects a measurement that cannot be mapped to the test sheet", () => {
  const missing = structuredClone(input);
  missing.measurements[0].uniqueId = "missing";
  assert.throws(() => buildSheetsBatchUpdate(missing), { message: "MEASUREMENT_ROW_NOT_FOUND" });
});

test("quotes apostrophes in A1 sheet names", () => {
  assert.equal(quoteSheetName("A'B"), "'A''B'");
});
