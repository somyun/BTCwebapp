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

test('search indexes every CAD label and supports multiple matches', () => {
    const manifest = JSON.parse(fs.readFileSync(path.join(root, 'cad-data', 'hopo', 'manifest.json'), 'utf8'));
    const labels = manifest.layers.flatMap((layer) => {
        const data = JSON.parse(fs.readFileSync(path.join(root, 'cad-data', 'hopo', layer.file), 'utf8'));
        return (data.labels || []).map((label) => ({ ...label, layerName: layer.name }));
    });
    const matches = labels.filter((label) => String(label.text).toLocaleLowerCase('ko-KR').includes('cf'));
    assert.equal(labels.length, 2396);
    assert.ok(matches.length > 1);
    assert.ok(matches.every((label) => label.layerName && label.position.length === 2));
    assert.match(mapSource, /manifest\.layers\.map\(async \(layerInfo\)/);
    assert.match(mapSource, /Array\.isArray\(layer\.labels\)/);
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
