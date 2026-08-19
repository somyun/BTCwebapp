'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

function read(relativePath) {
    return fs.readFileSync(path.join(__dirname, '..', relativePath), 'utf8');
}

test('legacy AHK-compatible GAS XLSX parameters remain unchanged', () => {
    const gas = read('apps-script/Code.js');
    assert.match(gas, /e\.parameter\.fileId && e\.parameter\.sheetName/);
    assert.match(gas, /const fileId = e\.parameter\.fileId/);
    assert.match(gas, /const filename = decodeURIComponent\(e\.parameter\.filename/);
    assert.match(gas, /const expectedRevision = e\.parameter\.expectedRevision/);
    assert.match(gas, /if \(expectedRevision\)/);
});

test('the web frontend obtains XLSX downloads from Firebase instead of GAS', () => {
    const script = read('script.js');
    assert.match(script, /GET_XLSX_DOWNLOAD_URL/);
    assert.match(script, /fetchXlsxDownload\(currentSheetInfo\.formKey\)/);
    assert.doesNotMatch(script, /GAS_API_URL\}\?fileId/);
});

test('Firebase exposes separate cache worker, download, and backfill functions', () => {
    const functions = read('firebase-production/functions/index.js');
    assert.match(functions, /exports\.getXlsxDownload = publicEndpoint/);
    assert.match(functions, /exports\.prepareXlsxExportJob = onDocumentCreated/);
    assert.match(functions, /exports\.prepareXlsxExport = manualAdmin/);
    assert.match(functions, /exports\.prepareAllXlsxExports = manualAdmin/);
});
