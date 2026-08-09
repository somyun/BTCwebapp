'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const mapSource = fs.readFileSync(path.join(root, 'map.js'), 'utf8');
const styles = fs.readFileSync(path.join(root, 'style.css'), 'utf8');

test('search button is the rightmost primary control and swaps to a cancellable search field', () => {
    const primary = html.slice(html.indexOf('id="mapPrimaryControls"'), html.indexOf('id="mapSearchPanel"'));
    assert.ok(primary.indexOf('id="mapSearchBtn"') > primary.indexOf('id="detailZoomBtn"'));
    assert.match(html, /id="mapSearchBtn"[^>]*aria-label="도면 문자 검색"/);
    assert.match(html, /id="mapSearchBackBtn"[^>]*>[\s\S]*?&lt;[\s\S]*?<\/button>/);
    assert.match(html, /id="mapSearchInput"[^>]*type="search"[^>]*placeholder="도면 문자 검색"/);
    assert.match(mapSource, /primary\.hidden = true/);
    assert.match(mapSource, /panel\.hidden = false/);
    assert.match(mapSource, /primary\.hidden = false/);
    assert.match(mapSource, /panel\.hidden = true/);
});

test('search indexes authenticated CAD labels and supports multiple matches', () => {
    assert.match(mapSource, /manifest\.layers\.map\(async \(layerInfo\)/);
    assert.match(mapSource, /const layer = await loadLayer\(layerInfo\)/);
    assert.match(mapSource, /Array\.isArray\(layer\.labels\)/);
    assert.match(mapSource, /position: \[Number\(position\[0\]\), Number\(position\[1\]\)\]/);
    assert.match(mapSource, /matches\.slice\(0, SEARCH_RESULT_LIMIT\)/);
});

test('selecting a result moves the map and displays a Kakao marker', () => {
    assert.match(mapSource, /new window\.kakao\.maps\.Marker\(\{/);
    assert.match(mapSource, /position: \[Number\(position\[0\]\), Number\(position\[1\]\)\]/);
    assert.match(mapSource, /map\.panTo\(position\)/);
    assert.match(mapSource, /if \(map\.getLevel\(\) > 2\) map\.setLevel\(2/);
    assert.match(mapSource, /button\.addEventListener\('click', \(\) => selectSearchResult\(result\)\)/);
});

test('search results are responsive and independently scrollable', () => {
    assert.match(styles, /\.map-search-results\s*{[\s\S]*?max-height:[\s\S]*?overflow: hidden;/);
    assert.match(styles, /\.map-search-list\s*{[\s\S]*?overflow-y: auto;[\s\S]*?scroll-padding-bottom: 12px;/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-type-controls\.search-active[\s\S]*?right: 70px;/);
});

test('CAD data is loaded only through the authenticated Storage client', () => {
    const storageSource = fs.readFileSync(path.join(root, 'cad-storage.js'), 'utf8');
    assert.match(mapSource, /BWACadStorage\.readJson\(CAD_MANIFEST_PATH/);
    assert.match(mapSource, /BWACadStorage\.readJson\(layer\.file/);
    assert.doesNotMatch(mapSource, /cad-data\/hopo|CAD_MANIFEST_URL/);
    assert.match(storageSource, /getBytes\(ref\(storage, `\$\{CAD_ROOT\}\/\$\{path\}`\)/);
    assert.match(storageSource, /token\.claims\.humetro !== true/);
});
