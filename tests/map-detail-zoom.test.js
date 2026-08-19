'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function classList() {
    const values = new Set();
    return {
        add: (...names) => names.forEach((name) => values.add(name)),
        remove: (...names) => names.forEach((name) => values.delete(name)),
        contains: (name) => values.has(name),
        toggle(name, force) {
            const enabled = force === undefined ? !values.has(name) : force;
            if (enabled) values.add(name); else values.delete(name);
            return enabled;
        }
    };
}

function element(overrides = {}) {
    const listeners = {};
    return {
        style: {}, dataset: {}, classList: classList(), clientWidth: 1000, clientHeight: 600,
        addEventListener: (name, handler) => { listeners[name] = handler; }, listeners,
        setAttribute(name, value) { this[name] = String(value); }, replaceChildren() {}, ...overrides
    };
}

async function createHarness() {
    const renderedLabel = element({ getAttribute(name) { return name === 'x' ? '125' : '240'; } });
    const overlayPane = element({
        appendChild(child) { child.parentNode = this; }
    });
    const elements = {
        mapView: element({ getBoundingClientRect: () => ({ left: 0, top: 0, width: 1000, height: 600 }) }),
        mapZoomStage: element({ setPointerCapture() {} }), naverMap: element(),
        cadOverlay: element({
            querySelectorAll: (selector) => selector === '.cad-map-label' ? [renderedLabel] : [],
            getBoundingClientRect: () => ({ left: 0, top: 0 })
        }),
        mapLoading: element(), cadLayerStatus: element(), cadLayerMessage: element(), cadLayerList: element(),
        cadOpacity: element({ value: '80' }), cadLabelToggle: element({ checked: true }),
        mapTypeToggleBtn: element(), currentLocationBtn: element(), displaySettingsBtn: element(),
        zoomInBtn: element(), zoomOutBtn: element(), cadLayerPanel: element({ hidden: true }),
        cadCoreLayersBtn: element(), cadAllLayersBtn: element(), cadNoLayersBtn: element(), mapLocationStatus: element()
    };

    const mapEvents = {};
    class LatLng { constructor(lat, lng) { this.lat = lat; this.lng = lng; } }
    class Point { constructor(x, y) { this.x = x; this.y = y; } }
    class Size { constructor(width, height) { this.width = width; this.height = height; } }
    class LatLngBounds { constructor(sw, ne) { this.sw = sw; this.ne = ne; } }
    class Overlay {
        constructor(options) { Object.assign(this, options); }
        setMap(map) { this.map = map; }
        setPosition(position) { this.position = position; }
        setCenter(center) { this.center = center; }
        setRadius(radius) { this.radius = radius; }
    }
    class FakeMap {
        constructor(_surface, options) {
            this.zoom = options.zoom; this.center = options.center; this.draggable = true; this.sizeCalls = 0;
        }
        setMapTypeId(type) { this.mapTypeId = type; }
        fitBounds() {}
        setZoom(zoom) { this.zoom = zoom; mapEvents.zoom_changed?.(); }
        getZoom() { return this.zoom; }
        getCenter() { return this.center; }
        setCenter(center) { this.center = center; mapEvents.center_changed?.(); }
        panTo(center) { this.setCenter(center); }
        setOptions(key, value) { if (key === 'draggable') this.draggable = value; }
        setSize() { this.sizeCalls += 1; }
        getPanes() { return { overlayLayer: overlayPane }; }
        getProjection() {
            return {
                fromCoordToOffset: (latLng) => ({ x: 500 + ((latLng.lng - 129) * 1000), y: 300 - ((latLng.lat - 35) * 1000) }),
                fromOffsetToCoord: (point) => new LatLng(35 - ((point.y - 300) / 1000), 129 + ((point.x - 500) / 1000))
            };
        }
    }
    class OverlayView {
        setMap(map) {
            this.map = map;
            if (map) {
                this.onAdd();
                this.draw();
            } else {
                this.onRemove();
            }
        }
        getPanes() { return this.map.getPanes(); }
    }

    let mapInstance;
    const windowListeners = {};
    const window = {
        naver: { maps: {
            Map: class extends FakeMap { constructor(...args) { super(...args); mapInstance = this; } },
            LatLng, Point, Size, LatLngBounds, Marker: Overlay, Circle: Overlay, OverlayView,
            MapTypeId: { SATELLITE: 'satellite', NORMAL: 'normal' },
            Event: { addListener: (_map, name, handler) => { mapEvents[name] = handler; } }
        } },
        requestAnimationFrame: (callback) => { callback(); return 1; }, setTimeout, clearTimeout,
        addEventListener: (name, handler) => { windowListeners[name] = handler; }, navigator: {},
        innerWidth: 600, innerHeight: 1000,
        screen: { orientation: { angle: 0, addEventListener() {} } }
    };
    const documentListeners = {};
    const document = {
        currentScript: { src: 'https://example.test/map.js' },
        getElementById: (id) => elements[id] || null,
        addEventListener: (name, handler) => { documentListeners[name] = handler; },
        querySelector: () => ({ content: 'test-ncp-key-id' }), querySelectorAll: () => [],
        createElement: () => element(), createElementNS: () => element(),
        createDocumentFragment: () => ({ appendChild() {} }), head: element()
    };
    const context = vm.createContext({
        window, document, navigator: window.navigator, console, Math, Number, String, Boolean, Map, Set, URL, Promise,
        fetch: async () => ({
            ok: true, url: 'https://example.test/cad-data/hopo/manifest.json',
            json: async () => ({ center_wgs84: [129, 35], bounds_wgs84: [128.99, 34.99, 129.01, 35.01], layers: [] })
        })
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'map.js'), 'utf8'), context);
    await window.BWAMap.initialize();
    return { elements, overlayPane, documentListeners, renderedLabel, window, windowListeners, getMap: () => mapInstance };
}

test('CAD SVG is mounted in the NAVER overlay pane so it shares native drag transforms', async () => {
    const { elements, overlayPane } = await createHarness();
    assert.equal(elements.cadOverlay.parentNode, overlayPane);
});

test('zoom controls use NAVER native zoom through level 21 without extra scaling', async () => {
    const { elements, getMap } = await createHarness();
    getMap().setZoom(20);
    elements.zoomInBtn.listeners.click();
    assert.equal(getMap().getZoom(), 21);
    assert.equal(elements.zoomInBtn.disabled, true);
    assert.doesNotMatch(elements.mapZoomStage.style.transform || '', /scale/);
    elements.zoomOutBtn.listeners.click();
    assert.equal(getMap().getZoom(), 20);
});

test('one map type button toggles between satellite and roadmap', async () => {
    const { elements, getMap } = await createHarness();
    assert.equal(elements.mapTypeToggleBtn.textContent, '위성지도');
    elements.mapTypeToggleBtn.listeners.click();
    assert.equal(elements.mapTypeToggleBtn.textContent, '일반지도');
    assert.equal(getMap().mapTypeId, 'normal');
    elements.mapTypeToggleBtn.listeners.click();
    assert.equal(getMap().mapTypeId, 'satellite');
});

test('an outside pointer closes the layer panel', async () => {
    const { elements, documentListeners } = await createHarness();
    elements.cadLayerPanel.contains = () => false;
    elements.displaySettingsBtn.contains = () => false;
    elements.displaySettingsBtn.listeners.click();
    documentListeners.pointerdown({ target: {} });
    assert.equal(elements.cadLayerPanel.hidden, true);
});

test('device orientation rotates the map and keeps labels readable', async () => {
    const { elements, renderedLabel, window, windowListeners, getMap } = await createHarness();
    window.innerWidth = 1000; window.innerHeight = 600; window.screen.orientation.angle = 90;
    await windowListeners.resize();
    assert.match(elements.mapZoomStage.style.transform, /rotate\(-90deg\)/);
    assert.equal(renderedLabel.transform, 'rotate(90 125 240)');
    assert.equal(getMap().draggable, false);
    assert.ok(getMap().sizeCalls >= 2);
    window.innerWidth = 600; window.innerHeight = 1000;
    await windowListeners.resize();
    assert.equal(elements.mapZoomStage.style.transform, '');
    assert.equal(getMap().draggable, true);
});

test('dragging while device-rotated follows the visible screen direction', async () => {
    const { elements, window, windowListeners, getMap } = await createHarness();
    window.innerWidth = 1000; window.innerHeight = 600; window.screen.orientation.angle = 90;
    await windowListeners.resize();
    const before = getMap().getCenter();
    elements.naverMap.listeners.touchstart({ touches: [{ clientX: 100, clientY: 100 }] });
    elements.naverMap.listeners.touchmove({ touches: [{ clientX: 140, clientY: 100 }] });
    const after = getMap().getCenter();
    assert.ok(after.lat > before.lat);
    assert.equal(after.lng, before.lng);
});

test('pinching while device-rotated keeps the CAD preview aligned', async () => {
    const { elements, window, windowListeners } = await createHarness();
    window.innerWidth = 1000; window.innerHeight = 600; window.screen.orientation.angle = 90;
    await windowListeners.resize();
    elements.naverMap.listeners.touchstart({ touches: [{ clientX: 100, clientY: 100 }, { clientX: 200, clientY: 100 }] });
    elements.naverMap.listeners.touchmove({ touches: [{ clientX: 70, clientY: 100 }, { clientX: 270, clientY: 100 }] });
    assert.equal(elements.cadOverlay.style.transform, 'translate3d(0px, 20px, 0) scale(2)');
    assert.doesNotMatch(elements.mapZoomStage.style.transform, /scale/);
});
