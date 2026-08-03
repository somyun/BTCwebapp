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
    assert.match(html, /id="displaySettingsBtn"[^>]*aria-controls="cadLayerPanel"/);
    assert.match(html, /id="currentLocationBtn"[^>]*class="map-current-location"/);
    assert.match(html, /class="location-dot"/);
    assert.match(html, /id="cadLayerPanel"[^>]*hidden/);
});

test('display settings button toggles the layer panel accessibly', () => {
    assert.match(mapSource, /panel\.hidden = !willOpen/);
    assert.match(mapSource, /setAttribute\('aria-expanded', String\(willOpen\)\)/);
    assert.match(styles, /\.cad-layer-panel\[hidden\]\s*{\s*display: none;/);
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
