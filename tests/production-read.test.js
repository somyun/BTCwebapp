'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const adapter = require('../read-adapter.js');

const source = fs.readFileSync(path.join(__dirname, '..', 'production-read.js'), 'utf8');

function loadRuntime(search = '', fetchImpl = async () => {
    throw new Error('UNEXPECTED_FETCH');
}, hostname = 'localhost') {
    const window = {
        BWAReadAdapter: adapter,
        location: { search, hostname },
        fetch: fetchImpl,
        setTimeout,
        clearTimeout
    };
    vm.runInNewContext(source, {
        window,
        URL,
        URLSearchParams,
        AbortController,
        Date,
        Map,
        Set,
        Object,
        Promise,
        console
    });
    return window;
}

test('production default remains GAS and makes no Firestore request', async () => {
    let fetchCount = 0;
    const window = loadRuntime('', async () => {
        fetchCount += 1;
        throw new Error('UNEXPECTED_FETCH');
    });
    const gasItems = [{
        sheetName: 'FORM_A',
        spreadsheetId: window.BWAProductionRead.CONFIG.spreadsheetId,
        lastModifiedDate: '2026-07-29T00:00:00.000Z'
    }];
    const result = await window.BWAProductionRead.loadFormList(async () => gasItems);
    assert.equal(result.servedBy, 'gas');
    assert.deepEqual(result.items, gasItems);
    assert.equal(fetchCount, 0);
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'gas');
});

test('Firestore mode falls back to GAS when production cache is unavailable', async () => {
    const window = loadRuntime('?readSource=firestore', async () => ({
        ok: false,
        status: 404,
        json: async () => ({})
    }));
    const gasItems = [{ sheetName: 'FORM_A' }];
    const result = await window.BWAProductionRead.loadFormList(async () => gasItems);
    assert.equal(result.servedBy, 'gas-fallback');
    assert.deepEqual(result.items, gasItems);
    assert.equal(window.BWA_PRODUCTION_READ_STATE.fallbackCount, 1);
    assert.equal(window.BWA_PRODUCTION_READ_STATE.lastFallback.operation, 'formList');
});

test('invalid read source cannot bypass the GAS default', async () => {
    const window = loadRuntime('?readSource=invalid');
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'gas');
});

test('hosted production ignores query-string read source overrides', async () => {
    const window = loadRuntime('?readSource=firestore', undefined, 'somyun.github.io');
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'gas');
});

test('adapter rejects a GAS list from a different spreadsheet', () => {
    assert.throws(() => adapter.normalizeGasFormList([{
        sheetName: 'FORM_A',
        spreadsheetId: 'wrong',
        lastModifiedDate: '2026-07-29T00:00:00.000Z'
    }], 'expected'), /SOURCE_SPREADSHEET_ID_MISMATCH/);
});
