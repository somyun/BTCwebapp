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
    assert.ok(primary.indexOf('id="mapSearchBtn"') > primary.indexOf('id="displaySettingsBtn"'));
    assert.doesNotMatch(primary, /id="detailZoomBtn"/);
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

test('selecting a result moves the map and displays a NAVER marker', () => {
    assert.match(mapSource, /new window\.naver\.maps\.Marker\(\{/);
    assert.match(mapSource, /position: \[Number\(position\[0\]\), Number\(position\[1\]\)\]/);
    assert.match(mapSource, /map\.panTo\(position\)/);
    assert.match(mapSource, /if \(map\.getZoom\(\) < 20\) map\.setZoom\(20, true\)/);
    assert.match(mapSource, /button\.addEventListener\('click', \(\) => selectSearchResult\(result, index\)\)/);
    assert.match(mapSource, /content: SEARCH_MARKER_ICON/);
    assert.match(styles, /\.cad-search-pin\s*{[\s\S]*?rotate\(var\(--map-counter-rotation, 0deg\)\)/);
});

test('selected search results can move to the previous or next match', () => {
    assert.match(html, /id="mapSearchPrevBtn"[^>]*aria-label="이전 검색 결과"/);
    assert.match(html, /id="mapSearchNextBtn"[^>]*aria-label="다음 검색 결과"/);
    assert.match(html, /id="mapSearchNavigationText"[^>]*title="검색 결과 목록 다시 보기"[^>]*>검색텍스트</);
    assert.match(html, /id="mapSearchNavigationLayer"[^>]*>레이어</);
    assert.match(html, /id="mapSearchNavigationStatus"[^>]*>0\/0</);
    assert.match(mapSource, /function moveSearchSelection\(offset\)/);
    assert.match(mapSource, /moveSearchSelection\(-1\)/);
    assert.match(mapSource, /moveSearchSelection\(1\)/);
    assert.match(mapSource, /status\.textContent = `\$\{selectedSearchResultIndex \+ 1\}\/\$\{activeSearchResults\.length\}`/);
    assert.match(mapSource, /text\.textContent = selectedResult\.text/);
    assert.match(mapSource, /layer\.textContent = selectedResult\.layerName/);
    assert.match(mapSource, /function reopenSearchResults\(\)/);
    assert.match(mapSource, /getElement\('mapSearchNavigationText'\)\?\.addEventListener\('click', reopenSearchResults\)/);
    assert.doesNotMatch(mapSource, /setLocationStatus\(`검색 위치:/);
    assert.match(styles, /\.map-search-navigation-arrow svg\s*{[\s\S]*?width: 30px;[\s\S]*?height: 30px;/);
});

test('search results are responsive and independently scrollable', () => {
    assert.match(styles, /\.map-search-results\s*{[\s\S]*?max-height:[\s\S]*?overflow: hidden;/);
    assert.match(styles, /\.map-search-list\s*{[\s\S]*?overflow-y: auto;[\s\S]*?scroll-padding-bottom: 12px;/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-type-controls\.search-active[\s\S]*?right: 70px;/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-type-controls\.search-active\.result-selected\s*{[\s\S]*?display: flex;[\s\S]*?flex-wrap: nowrap;/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-type-controls\.result-selected \.map-search-navigation\s*{[\s\S]*?position: static;/);
});

test('CAD data resolves beside the deployed map script for test-site parity', () => {
    assert.match(mapSource, /document\.currentScript\?\.src/);
    assert.match(mapSource, /new URL\('cad-data\/hopo\/manifest\.json', MAP_SCRIPT_BASE_URL\)\.href/);
});
