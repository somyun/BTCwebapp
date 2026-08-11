'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.join(__dirname, '..');
const productionSource = fs.readFileSync(path.join(root, 'script.js'), 'utf8');
const productionStyles = fs.readFileSync(path.join(root, 'style.css'), 'utf8');
const testbedRoot = path.join(root, 'testbed', 'bwa_test_publish');
const testbedSource = fs.readFileSync(path.join(testbedRoot, 'testbed.js'), 'utf8');
const testbedStyles = fs.readFileSync(path.join(testbedRoot, 'style.css'), 'utf8');

const productionHomeOnlyIds = [
    'favoritesSection',
    'formMessage',
    'openMapBtn'
];

const testbedHomeOnlyIds = [
    ...productionHomeOnlyIds.slice(0, 2),
    'mainToggleContainer',
    'openMapBtn'
];

function homeOnlyIds(source) {
    const declaration = source.match(/const HOME_ONLY_ELEMENT_IDS = \[([\s\S]*?)\];/);
    assert.ok(declaration, 'HOME_ONLY_ELEMENT_IDS declaration is required');
    return [...declaration[1].matchAll(/'([^']+)'/g)].map((match) => match[1]);
}

test('production removes the legacy notification control from its home-only registry', () => {
    assert.deepEqual(homeOnlyIds(productionSource), productionHomeOnlyIds);
    assert.deepEqual(homeOnlyIds(testbedSource), testbedHomeOnlyIds);
    for (const source of [productionSource, testbedSource]) {
        assert.match(source, /function setHomeOnlyElementsVisible\(isVisible\)/);
        assert.match(source, /classList\.toggle\('home-only-hidden', !isVisible\)/);
        assert.match(source, /setHomeOnlyElementsVisible\(false\)/);
        assert.match(source, /setHomeOnlyElementsVisible\(true\)/);
    }
});

test('the shared hidden state removes every home-only control from layout', () => {
    assert.match(productionStyles, /\.home-only-hidden\s*{\s*display: none !important;/);
    assert.match(testbedStyles, /\.home-only-hidden\s*{\s*display: none !important;/);
    assert.doesNotMatch(productionSource, /favoritesSection'\)\.classList\.(?:add|remove)\('hidden'\)/);
    assert.doesNotMatch(testbedSource, /favoritesSection'\)\?\.classList\.(?:add|remove)\('hidden'\)/);
});

test('production invalidates stale form responses when selection or history changes', () => {
    assert.match(productionSource, /let selectedFormRequestId = 0;/);
    assert.match(productionSource, /const requestId = \+\+selectedFormRequestId;/);
    assert.match(productionSource, /if \(requestId !== selectedFormRequestId\) return;/);
    assert.match(productionSource, /if \(!sheetName\) \{[\s\S]*?selectedFormRequestId \+= 1;/);
});

test('bwa_test installs the same form-to-home browser history behavior', () => {
    assert.match(testbedSource, /function addHomeStateToHistory\(\)[\s\S]*?history\.pushState\(\{ page: 'home' \}/);
    assert.match(testbedSource, /new URL\(window\.location\.href\)[\s\S]*?searchParams\.set\('page', 'home'\)/);
    assert.match(testbedSource, /window\.addEventListener\('popstate',[\s\S]*?formSelect\.value = ''[\s\S]*?loadSelectedForm\(\)/);
    assert.ok((testbedSource.match(/addHomeStateToHistory\(\);/g) || []).length >= 3);
});
