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

function firestoreFormListResponse() {
    const revision = '2026-07-29T00:00:00.000Z';
    return {
        ok: true,
        status: 200,
        json: async () => ({
            fields: {
                schemaVersion: { integerValue: '1' },
                itemCount: { integerValue: '1' },
                sourceRevision: { stringValue: revision },
                items: {
                    arrayValue: {
                        values: [{
                            mapValue: {
                                fields: {
                                    formKey: { stringValue: `f_${'a'.repeat(32)}` },
                                    sheetName: { stringValue: 'FORM_A' },
                                    displayName: { stringValue: 'FORM_A' },
                                    lastModifiedDate: { stringValue: revision }
                                }
                            }
                        }]
                    }
                }
            }
        })
    };
}

test('production default serves the form list from Firestore when available', async () => {
    let gasFetchCount = 0;
    const window = loadRuntime('', async () => firestoreFormListResponse());
    const result = await window.BWAProductionRead.loadFormList(async () => {
        gasFetchCount += 1;
        return [];
    });

    assert.equal(result.servedBy, 'firestore');
    assert.equal(result.items.length, 1);
    assert.equal(result.items[0].sheetName, 'FORM_A');
    assert.equal(result.items[0].displayName, 'FORM_A');
    assert.equal(gasFetchCount, 0);
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'firestore');
});

test('production default shows a Firestore error without calling GAS when unavailable', async () => {
    let fetchCount = 0;
    const window = loadRuntime('', async () => {
        fetchCount += 1;
        return {
            ok: false,
            status: 503,
            json: async () => ({})
        };
    });
    let gasFetchCount = 0;
    await assert.rejects(window.BWAProductionRead.loadFormList(async () => {
        gasFetchCount += 1;
        return [];
    }), /HTTP_503/);
    assert.equal(fetchCount, 1);
    assert.equal(gasFetchCount, 0);
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'firestore');
});

test('Firestore mode fails closed when production cache is unavailable', async () => {
    const window = loadRuntime('?readSource=firestore', async () => ({
        ok: false,
        status: 404,
        json: async () => ({})
    }));
    let gasFetchCount = 0;
    await assert.rejects(window.BWAProductionRead.loadFormList(async () => {
        gasFetchCount += 1;
        return [];
    }), /HTTP_404/);
    assert.equal(gasFetchCount, 0);
});

test('invalid read source cannot bypass the Firestore default', async () => {
    const window = loadRuntime('?readSource=invalid');
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'firestore');
});

test('hosted production ignores query-string read source overrides', async () => {
    const window = loadRuntime('?readSource=gas', undefined, 'somyun.github.io');
    assert.equal(window.BWA_PRODUCTION_READ_STATE.source, 'firestore');
});

test('production runtime contains no GAS fallback or read-source switch', () => {
    assert.doesNotMatch(source, /gasFallback|readSource|shadow|gas-fallback/);
});

test('adapter rejects a GAS list from a different spreadsheet', () => {
    assert.throws(() => adapter.normalizeGasFormList([{
        sheetName: 'FORM_A',
        spreadsheetId: 'wrong',
        lastModifiedDate: '2026-07-29T00:00:00.000Z'
    }], 'expected'), /SOURCE_SPREADSHEET_ID_MISMATCH/);
});
