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
    assert.match(html, /id="currentLocationBtn"[\s\S]*?title="현재위치"[\s\S]*?aria-pressed="false"/);
    assert.match(html, /class="location-dot"/);
    assert.match(html, /id="cadLayerPanel"[^>]*hidden/);
});

test('depot map entry is on the home page below notifications and labels default to visible', () => {
    const sidebar = html.slice(html.indexOf('<nav id="sidenav"'), html.indexOf('</nav>'));
    const homeView = html.slice(html.indexOf('<div id="homeView"'), html.indexOf('<section id="mapView"'));
    assert.doesNotMatch(sidebar, /id="openMapBtn"/);
    assert.match(homeView, /id="openMapBtn"/);
    assert.doesNotMatch(html, /id="notificationToggleMain"/);
    assert.match(html, /id="openMapBtn"[^>]*>[\s\S]*?호포기지 도면보기[\s\S]*?<\/button>/);
    assert.match(html, /id="cadLabelToggle"[^>]*type="checkbox"[^>]*checked/);
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
    assert.match(mapSource, /screenPointToStagePoint\(midpoint\)/);
    assert.match(mapSource, /screenPointToCanvasPoint\(midpoint\)/);
    assert.match(mapSource, /translate3d\(\$\{translateX\}px, \$\{translateY\}px, 0\) scale\(\$\{scale\}\)/);
    assert.match(styles, /\.cad-map-overlay\.pinching\s*{[\s\S]*?transition: none;/);
});

test('NAVER native zoom reaches level 21 without an additional zoom control', () => {
    assert.match(html, /id="mapZoomStage" class="map-zoom-stage"/);
    assert.match(html, /id="zoomInBtn"[^>]*aria-label="지도 확대"/);
    assert.match(html, /id="zoomOutBtn"[^>]*aria-label="지도 축소"/);
    assert.doesNotMatch(html, /id="detailZoomBtn"|추가확대/);
    assert.match(mapSource, /NAVER_MAX_ZOOM = 21/);
    assert.match(mapSource, /map\.setZoom\(Math\.min\(NAVER_MAX_ZOOM, map\.getZoom\(\) \+ 1\), true\)/);
    assert.match(mapSource, /map\.setOptions\('draggable', !active\)/);
    assert.match(mapSource, /if \(customTransformActive\(\) && event\.touches\.length === 1\)/);
    assert.doesNotMatch(mapSource, /detailScale|DETAIL_ZOOM_STEPS|detailTransformActive/);
    assert.match(mapSource, /stage\.style\.transform = active/);
    assert.match(styles, /\.map-zoom-controls\s*{[\s\S]*?top: 50%;[\s\S]*?right: 12px;/);
});

test('team layers are included by default and the landscape list uses remaining height', () => {
    assert.match(mapSource, /'teamA', 'teamB', 'teamC', 'teamD'/);
    assert.match(mapSource, /'TEAM_A', 'TEAM_B', 'TEAM_C', 'TEAM_D'/);
    assert.match(styles, /@media \(orientation: landscape\) and \(max-height: 720px\)/);
    assert.match(styles, /\.map-view\.landscape-mode \.cad-layer-panel:not\(\[hidden\]\)[\s\S]*?top: 60px;[\s\S]*?bottom: 8px;/);
    assert.match(styles, /\.map-view\.landscape-mode \.cad-layer-list[\s\S]*?flex: 1 1 auto;[\s\S]*?max-height: none;[\s\S]*?padding-bottom: 12px;/);
});

test('device orientation automatically rotates the map without an app button', () => {
    assert.doesNotMatch(html, /id="orientationModeBtn"|class="map-orientation-button"/);
    assert.doesNotMatch(html, /id="rotateMapBtn"|id="mapCompassBtn"/);
    assert.doesNotMatch(mapSource, /detailRotation|rotateMapStep|resetMapRotation/);
    assert.doesNotMatch(mapSource, /screen\.orientation\.lock|toggleOrientationMode/);
    assert.match(mapSource, /syncOrientationFromDevice\(\)/);
    assert.match(mapSource, /detectedMapRotation\(viewportLandscape\)/);
    assert.match(mapSource, /rotate\(\$\{mapRotationDegrees\}deg\)/);
    assert.doesNotMatch(styles, /\.map-orientation-button/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-type-controls\s*{[\s\S]*?right: auto;[\s\S]*?left: 12px;/);
    assert.match(styles, /\.map-view\.landscape-mode \.map-zoom-controls\s*{[\s\S]*?right: 12px;/);
});

test('CAD overlay keeps vector labels upright and stable at native zoom levels', () => {
    assert.match(html, /<svg id="cadOverlay"/);
    assert.match(mapSource, /createSvgElement\('path'/);
    assert.match(mapSource, /createSvgElement\('text'/);
    assert.match(mapSource, /class: 'cad-map-label'/);
    assert.match(mapSource, /transform: `rotate\(\$\{-mapRotationDegrees\} \$\{x\} \$\{y\}\)`/);
    assert.match(mapSource, /label\.setAttribute\('transform', `rotate\(\$\{-mapRotationDegrees\}/);
    assert.match(mapSource, /relayoutMap\(\);[\s\S]*?map\.setCenter\(center\);/);
    assert.doesNotMatch(mapSource, /LABEL_DETAIL_SCALE_COMPENSATION|detailScale/);
    assert.match(mapSource, /'vector-effect': 'non-scaling-stroke'/);
    assert.match(styles, /\.cad-map-label\s*{[\s\S]*?font-size: var\(--cad-label-font-size, 14px\)/);
});

test('current location uses a circular marker and toggles real-time tracking', () => {
    assert.match(mapSource, /navigator\.geolocation\.watchPosition\(updateTrackedPosition/);
    assert.match(mapSource, /navigator\.geolocation\.clearWatch\(locationWatchId\)/);
    assert.match(mapSource, /content: CURRENT_POSITION_ICON/);
    assert.match(styles, /\.current-position-marker\s*{[\s\S]*?border-radius: 50%;[\s\S]*?background: #1677ff;/);
    assert.match(styles, /\.map-current-location\.active\s*{/);
});

test('display settings offer small, medium, and large CAD label sizes', () => {
    assert.match(html, /id="cadLabelToggle"[^>]*>[\s\S]*?<span>문자<\/span>/);
    assert.doesNotMatch(html, /문자도 표시/);
    assert.match(html, /data-cad-label-size="small"[^>]*>작게<\/button>/);
    assert.match(html, /data-cad-label-size="medium"[^>]*>중간<\/button>/);
    assert.match(html, /data-cad-label-size="large"[^>]*>크게<\/button>/);
    assert.match(mapSource, /LABEL_FONT_SIZES = Object\.freeze\(\{ small: 14, medium: 17, large: 20 \}\)/);
    assert.match(mapSource, /canvas\.style\.setProperty\('--cad-label-font-size', fontSize\)/);
});

test('CAD SVG uses NAVER OverlayView so native dragging moves map and drawing together', () => {
    assert.match(mapSource, /class CadSvgOverlay extends window\.naver\.maps\.OverlayView/);
    assert.match(mapSource, /this\.getPanes\(\)\.overlayLayer\.appendChild\(canvas\)/);
    assert.match(mapSource, /cadOverlayView\.setMap\(map\)/);
});
