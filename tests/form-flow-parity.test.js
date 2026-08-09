'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.join(__dirname, '..');
const read = (...segments) => fs.readFileSync(path.join(root, ...segments), 'utf8');
const production = read('script.js');
const productionRead = read('production-read.js');
const auth = read('auth.js');
const testbed = read('testbed', 'bwa_test_publish', 'testbed.js');
const testIndex = read('testbed', 'bwa_test_publish', 'index.html');
const testGas = read('apps-script-test', 'Code.js');
const testAuthBackend = read('apps-script-test', 'auth_code.js');

test('production form reads fail closed and both clients render retry actions', () => {
    assert.match(productionRead, /gasFallback:\s*false/);
    assert.match(production, /function renderListFailure\(error\)[\s\S]*?목록 다시 시도/);
    assert.match(testbed, /function renderListFailure\(error\)[\s\S]*?목록 다시 시도/);
    assert.match(production, /function renderFormError\(message, retryHandler\)[\s\S]*?다시 시도/);
    assert.match(testbed, /function renderFormError\(message, retryHandler\)[\s\S]*?다시 시도/);
    assert.match(production, /catch \(error\)[\s\S]*?폼 로딩 오류:[\s\S]*?renderFormError\(message, loadSelectedForm\)/);
    assert.match(testbed, /catch \(error\)[\s\S]*?폼 로딩 오류:[\s\S]*?renderFormError\(message, loadSelectedForm\)/);
    assert.match(production, /폼 생성 실패:[\s\S]*?processFileUpload\(fileData, userChoice\)/);
    assert.match(testbed, /폼 생성 실패:[\s\S]*?processFileUpload\(fileData, userChoice\)/);
});

test('production save uses Functions, waits for Sheet sync, and only then prepares XLSX', () => {
    assert.match(production, /cloudfunctions\.net\/submitMeasurements/);
    assert.match(production, /cloudfunctions\.net\/getMeasurementSubmission/);
    assert.match(production, /await pollSubmission\(payload\.idempotencyKey\)/);
    const syncIndex = production.indexOf("updateSubmissionStatus(\n            `Google Sheet 동기화 완료");
    const xlsxIndex = production.indexOf('await prepareXlsxAfterSync();', syncIndex);
    assert.ok(syncIndex >= 0 && xlsxIndex > syncIndex, 'XLSX must be prepared after Sheet sync');
    assert.doesNotMatch(production, /callApi\(['"]saveMeasurementsToSheet['"]/);
});

test('bwa_test map entry uses the shared email-auth gate with test-only credentials', () => {
    assert.match(auth, /window\.BWA_AUTH_CONFIG \|\| window\.BWA_TEST_CONFIG/);
    assert.match(testbed, /async function openMapView\(\)[\s\S]*?BWAAuth\.requireMapAccess\(\)/);
    assert.match(testIndex, /firebase-auth-compat\.js/);
    assert.match(testGas, /case 'requestMapAuthCode':[\s\S]*?case 'verifyMapAuthCode':/);
    assert.match(testGas, /allowedScriptPropertyKeys[\s\S]*?'MAP_AUTH_CODE_PEPPER'/);
    assert.match(testAuthBackend, /SERVICE_ACCOUNT_KEY\.project_id !== TEST_FIREBASE_PROJECT_ID/);
    assert.match(testAuthBackend, /TEST_FIREBASE_SERVICE_ACCOUNT_REQUIRED/);
});

test('the deliberate local-setting isolation remains test-prefixed', () => {
    assert.match(testbed, /const STORAGE_PREFIX = 'bwa_test:';/);
    assert.match(testbed, /window\.localStorage\.setItem\(storageKey\(key\)/);
    assert.doesNotMatch(production, /const STORAGE_PREFIX = 'bwa_test:';/);
});
