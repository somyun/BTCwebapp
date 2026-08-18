'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '..', 'script.js'), 'utf8');

test('today cache values prefill measurement inputs', () => {
    assert.match(source, /const prefillTodayValues = currentSheetInfo\?\.hasTodayMeasurementCache === true/);
    assert.match(source, /input\.value = prefillTodayValues \? String\(data\.value \?\? ''\) : ''/);
});

test('a successful save restores the reusable save button', () => {
    assert.match(source, /setSaveButtonBusy\(false, '측정값 저장'\)/);
    assert.doesNotMatch(source, /setSaveButtonBusy\(true, '저장 완료'\)/);
});

test('save toast uses three concise stages and XLSX preparation stays separate', () => {
    assert.match(source, /showStatus\('저장 중\.\.\.', 'loading'\)/);
    assert.match(source, /showStatus\('측정값을 반영하고 있습니다\.\.\.', 'loading'\)/);
    assert.match(source, /setSaveButtonBusy\(false, '측정값 저장'\);\s*showStatus\('성공적으로 저장하였습니다!', 'success', 3000\)/);
    assert.match(source, /button\.textContent = `\$\{date\}\.xlsx 준비중\.\.`/);
    assert.doesNotMatch(source, /showStatus\('Firebase 캐시 저장 완료/);
    assert.doesNotMatch(source, /측정값이 Google Sheet에 저장되었습니다/);
});

test('Firebase cache success unlocks saving before Sheets follow-up completes', () => {
    assert.match(source, /currentSheetInfo\.hasTodayMeasurementCache = true/);
    assert.match(source, /setSaveButtonBusy\(false, '측정값 저장'\);\s*showStatus\('성공적으로 저장하였습니다!'/);
    assert.match(source, /void monitorSheetSync\(payload\.idempotencyKey, cached,/);
    assert.match(source, /Firebase에는 저장했지만 Google Sheets 반영에 실패했습니다/);
});
