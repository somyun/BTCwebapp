'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const mapSource = fs.readFileSync(path.join(root, 'map.js'), 'utf8');
const styles = fs.readFileSync(path.join(root, 'style.css'), 'utf8');

test('map controls expose display settings and an icon-only location button', () => {
    assert.match(html, /id="mapTypeToggleBtn"[^>]*aria-pressed="true"/);
    assert.doesNotMatch(html, /id="roadmapBtn"|id="skyviewBtn"|id="recenterMapBtn"/);
    assert.match(html, /id="displaySettingsBtn"[^>]*aria-controls="cadLayerPanel"/);
    assert.match(html, /id="currentLocationBtn"[^>]*class="map-current-location"/);
    assert.match(html, /class="location-dot"/);
    assert.match(html, /id="cadLayerPanel"[^>]*hidden/);
});

test('display settings button toggles the layer panel accessibly', () => {
    assert.match(mapSource, /panel\.hidden = !willOpen/);
    assert.match(mapSource, /setAttribute\('aria-expanded', String\(willOpen\)\)/);
    assert.match(styles, /\.cad-layer-panel\[hidden\]\s*{\s*display: none;/);
    assert.match(mapSource, /document\.addEventListener\('pointerdown'/);
    assert.match(mapSource, /panel\.contains\(event\.target\) \|\| button\.contains\(event\.target\)/);
});

test('mobile pinch updates the CAD canvas transform during the gesture', () => {
    assert.match(mapSource, /addEventListener\('touchmove', updatePinch/);
    assert.match(mapSource, /translate3d\(\$\{translateX\}px, \$\{translateY\}px, 0\) scale\(\$\{scale\}\)/);
    assert.match(styles, /\.cad-map-overlay\.pinching\s*{[\s\S]*?transition: none;/);
});

test('detail zoom cycles beyond the Kakao tile limit and keeps controls fixed', () => {
    assert.match(html, /id="mapZoomStage" class="map-zoom-stage"/);
    assert.match(html, /id="detailZoomBtn"[^>]*aria-pressed="false"/);
    assert.match(html, /id="zoomInBtn"[^>]*aria-label="지도 확대"/);
    assert.match(html, /id="zoomOutBtn"[^>]*aria-label="지도 축소"/);
    assert.match(mapSource, /DETAIL_ZOOM_STEPS = \[1, 2, 4, 8\]/);
    assert.match(mapSource, /map\.setDraggable\(!active\)/);
    assert.match(mapSource, /stage\.style\.transform = active/);
    assert.match(styles, /\.map-zoom-controls\s*{[\s\S]*?top: 50%;[\s\S]*?right: 12px;/);
});

test('device orientation automatically controls landscape layout without an app button', () => {
    assert.doesNotMatch(html, /id="orientationModeBtn"|class="map-orientation-button"/);
    assert.doesNotMatch(html, /id="rotateMapBtn"|id="mapCompassBtn"/);
    assert.doesNotMatch(mapSource, /detailRotation|rotateMapStep|resetMapRotation/);
    assert.doesNotMatch(mapSource, /screen\.orientation\.lock|toggleOrientationMode/);
    assert.match(mapSource, /syncOrientationFromDevice\(\)/);
    assert.match(mapSource, /angle === 270[\s\S]*?return 90;[\s\S]*?angle === 90[\s\S]*?return -90;/);
    assert.doesNotMatch(styles, /\.map-orientation-button/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-type-controls\s*{[\s\S]*?right: auto;[\s\S]*?left: 12px;/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-zoom-controls\s*{[\s\S]*?right: 12px;/);
});

test('CAD overlay keeps vector labels upright by mode and constant in detail zoom', () => {
    assert.match(html, /<svg id="cadOverlay"/);
    assert.match(mapSource, /createSvgElement\('path'/);
    assert.match(mapSource, /createSvgElement\('text'/);
    assert.match(mapSource, /class: 'cad-map-label'/);
    assert.match(mapSource, /detectedLabelRotation\(viewportLandscape\)/);
    assert.match(mapSource, /String\(1 \/ detailScale\)/);
    assert.match(mapSource, /'vector-effect': 'non-scaling-stroke'/);
    assert.match(styles, /\.cad-map-label\s*{[\s\S]*?rotate\(var\(--cad-label-rotation\)\) scale\(var\(--cad-label-inverse-scale\)\)/);
});
