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
        style: {},
        dataset: {},
        classList: classList(),
        addEventListener: (name, handler) => { listeners[name] = handler; },
        listeners,
        setAttribute(name, value) { this[name] = String(value); },
        replaceChildren() {},
        ...overrides
    };
}

async function createHarness() {
    const renderedLabel = element({
        getAttribute(name) { return name === 'x' ? '125' : '240'; }
    });
    const elements = {
        mapView: element({
            clientWidth: 1000,
            clientHeight: 600,
            getBoundingClientRect: () => ({ left: 0, top: 0, width: 1000, height: 600 })
        }),
        mapZoomStage: element({ setPointerCapture() {} }),
        kakaoMap: element(),
        cadOverlay: element({
            width: 0,
            height: 0,
            querySelectorAll: (selector) => selector === '.cad-map-label' ? [renderedLabel] : [],
            getBoundingClientRect: () => ({ left: 0, top: 0 }),
            getContext: () => ({
                clearRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
                fillRect() {}, fillText() {}
            })
        }),
        mapLoading: element(),
        cadLayerStatus: element(),
        cadLayerMessage: element(),
        cadLayerList: element(),
        cadOpacity: element({ value: '80' }),
        cadLabelToggle: element({ checked: false }),
        mapTypeToggleBtn: element(),
        currentLocationBtn: element(),
        displaySettingsBtn: element(),
        detailZoomBtn: element(),
        zoomInBtn: element(),
        zoomOutBtn: element(),
        cadLayerPanel: element({ hidden: true }),
        cadCoreLayersBtn: element(),
        cadAllLayersBtn: element(),
        cadNoLayersBtn: element(),
        mapLocationStatus: element()
    };

    const mapEvents = {};
    class FakeMap {
        constructor() {
            this.level = 4;
            this.draggable = true;
            this.zoomable = true;
            this.center = new LatLng(35, 129);
            this.relayoutCalls = 0;
        }
        setMapTypeId() {}
        setBounds() {}
        setLevel(level) {
            this.level = level;
            mapEvents.zoom_changed?.();
        }
        getLevel() { return this.level; }
        getCenter() { return this.center; }
        setCenter(center) { this.center = center; mapEvents.center_changed?.(); }
        setDraggable(value) { this.draggable = value; }
        setZoomable(value) { this.zoomable = value; }
        relayout() { this.relayoutCalls += 1; }
        getProjection() {
            return {
                containerPointFromCoords: (latLng) => ({
                    x: 500 + ((latLng.lng - 129) * 1000),
                    y: 300 - ((latLng.lat - 35) * 1000)
                }),
                coordsFromContainerPoint: (point) => new LatLng(
                    35 - ((point.y - 300) / 1000),
                    129 + ((point.x - 500) / 1000)
                )
            };
        }
    }
    class LatLng {
        constructor(lat, lng) { this.lat = lat; this.lng = lng; }
    }
    class Point {
        constructor(x, y) { this.x = x; this.y = y; }
    }
    class LatLngBounds { extend() {} }

    let mapInstance;
    const windowListeners = {};
    const orientationListeners = {};
    const window = {
        kakao: {
            maps: {
                load: (callback) => callback(),
                Map: class extends FakeMap {
                    constructor(...args) { super(...args); mapInstance = this; }
                },
                LatLng,
                Point,
                LatLngBounds,
                MapTypeId: { SKYVIEW: 'skyview', ROADMAP: 'roadmap' },
                event: {
                    addListener: (_map, name, handler) => { mapEvents[name] = handler; },
                    trigger: (_map, name) => mapEvents[name]?.()
                }
            }
        },
        requestAnimationFrame: (callback) => { callback(); return 1; },
        setTimeout,
        clearTimeout,
        addEventListener: (name, handler) => { windowListeners[name] = handler; },
        navigator: {},
        innerWidth: 600,
        innerHeight: 1000,
        screen: {
            orientation: {
                angle: 0,
                addEventListener: (name, handler) => { orientationListeners[name] = handler; }
            }
        }
    };
    const documentListeners = {};
    const document = {
        getElementById: (id) => elements[id] || null,
        addEventListener: (name, handler) => { documentListeners[name] = handler; },
        querySelectorAll: () => [],
        createElement: () => element(),
        createElementNS: () => element(),
        createDocumentFragment: () => ({ appendChild() {} }),
        head: element()
    };
    const context = vm.createContext({
        window,
        document,
        navigator: window.navigator,
        console,
        Math,
        Number,
        String,
        Boolean,
        Map,
        Set,
        URL,
        Promise,
        fetch: async () => ({
            ok: true,
            url: 'https://example.test/cad-data/hopo/manifest.json',
            json: async () => ({
                center_wgs84: [129, 35],
                bounds_wgs84: [128.99, 34.99, 129.01, 35.01],
                layers: []
            })
        })
    });
    const source = fs.readFileSync(path.join(__dirname, '..', 'map.js'), 'utf8');
    vm.runInContext(source, context);
    await window.BWAMap.initialize();
    return { elements, documentListeners, orientationListeners, renderedLabel, window, windowListeners, getMap: () => mapInstance };
}

test('detail button cycles 2x, 4x, 8x and restores normal map interaction', async () => {
    const { elements, getMap } = await createHarness();
    const click = elements.detailZoomBtn.listeners.click;

    click();
    assert.equal(getMap().getLevel(), 0);
    assert.match(elements.mapZoomStage.style.transform, /scale\(2\)/);
    assert.equal(getMap().draggable, false);

    click();
    assert.match(elements.mapZoomStage.style.transform, /scale\(4\)/);
    click();
    assert.match(elements.mapZoomStage.style.transform, /scale\(8\)/);
    click();
    assert.equal(elements.mapZoomStage.style.transform, '');
    assert.equal(getMap().draggable, true);
    assert.equal(getMap().zoomable, true);
});

test('right-side zoom controls enter and leave detail zoom at the tile limit', async () => {
    const { elements, getMap } = await createHarness();
    getMap().setLevel(0);
    elements.zoomInBtn.listeners.click();
    assert.match(elements.mapZoomStage.style.transform, /scale\(2\)/);
    elements.zoomOutBtn.listeners.click();
    assert.equal(elements.mapZoomStage.style.transform, '');
});

test('one map type button toggles between satellite and roadmap', async () => {
    const { elements } = await createHarness();
    assert.equal(elements.mapTypeToggleBtn.textContent, '위성지도');
    elements.mapTypeToggleBtn.listeners.click();
    assert.equal(elements.mapTypeToggleBtn.textContent, '일반지도');
    assert.equal(elements.mapTypeToggleBtn['aria-pressed'], 'false');
    elements.mapTypeToggleBtn.listeners.click();
    assert.equal(elements.mapTypeToggleBtn.textContent, '위성지도');
});

test('an outside pointer closes the layer panel', async () => {
    const { elements, documentListeners } = await createHarness();
    elements.cadLayerPanel.contains = () => false;
    elements.displaySettingsBtn.contains = () => false;
    elements.displaySettingsBtn.listeners.click();
    assert.equal(elements.cadLayerPanel.hidden, false);
    documentListeners.pointerdown({ target: {} });
    assert.equal(elements.cadLayerPanel.hidden, true);
    assert.equal(elements.displaySettingsBtn['aria-expanded'], 'false');
});

test('device orientation rotates the map toward the hardware top and keeps labels readable', async () => {
    const { elements, renderedLabel, window, windowListeners, getMap } = await createHarness();

    window.innerWidth = 1000;
    window.innerHeight = 600;
    window.screen.orientation.angle = 90;
    await windowListeners.resize();
    assert.equal(elements.mapView.classList.contains('landscape-mode'), true);
    assert.match(elements.mapZoomStage.style.transform, /rotate\(-90deg\)/);
    assert.equal(elements.mapZoomStage.style.width, '600px');
    assert.equal(elements.mapZoomStage.style.height, '1000px');
    assert.equal(renderedLabel.transform, 'rotate(90 125 240)');
    assert.equal(getMap().draggable, false);
    assert.equal(getMap().zoomable, true);
    assert.ok(getMap().relayoutCalls >= 2);

    window.screen.orientation.angle = 270;
    await windowListeners.resize();
    assert.match(elements.mapZoomStage.style.transform, /rotate\(90deg\)/);
    assert.equal(renderedLabel.transform, 'rotate(-90 125 240)');

    window.innerWidth = 600;
    window.innerHeight = 1000;
    await windowListeners.resize();
    assert.equal(elements.mapView.classList.contains('landscape-mode'), false);
    assert.equal(elements.mapZoomStage.style.transform, '');
    assert.equal(renderedLabel.transform, 'rotate(0 125 240)');
    assert.equal(getMap().draggable, true);
});

test('dragging while device-rotated follows the visible screen direction', async () => {
    const { elements, window, windowListeners, getMap } = await createHarness();
    window.innerWidth = 1000;
    window.innerHeight = 600;
    window.screen.orientation.angle = 90;
    await windowListeners.resize();

    const beforeClockwise = getMap().getCenter();
    elements.kakaoMap.listeners.touchstart({
        touches: [{ clientX: 100, clientY: 100 }]
    });
    elements.kakaoMap.listeners.touchmove({
        touches: [{ clientX: 140, clientY: 100 }]
    });
    const afterClockwise = getMap().getCenter();
    assert.ok(afterClockwise.lat > beforeClockwise.lat);
    assert.equal(afterClockwise.lng, beforeClockwise.lng);

    elements.kakaoMap.listeners.touchend({ touches: [] });
    window.screen.orientation.angle = 270;
    await windowListeners.resize();
    const beforeCounterClockwise = getMap().getCenter();
    elements.kakaoMap.listeners.touchstart({
        touches: [{ clientX: 100, clientY: 100 }]
    });
    elements.kakaoMap.listeners.touchmove({
        touches: [{ clientX: 140, clientY: 100 }]
    });
    const afterCounterClockwise = getMap().getCenter();
    assert.ok(afterCounterClockwise.lat < beforeCounterClockwise.lat);
    assert.equal(afterCounterClockwise.lng, beforeCounterClockwise.lng);
});

test('pinching while device-rotated stays in Kakao map zoom instead of detail zoom', async () => {
    const { elements, window, windowListeners, getMap } = await createHarness();
    window.innerWidth = 1000;
    window.innerHeight = 600;
    window.screen.orientation.angle = 90;
    await windowListeners.resize();
    elements.kakaoMap.listeners.touchstart({
        touches: [
            { clientX: 100, clientY: 100 },
            { clientX: 200, clientY: 100 }
        ]
    });
    elements.kakaoMap.listeners.touchmove({
        touches: [
            { clientX: 50, clientY: 100 },
            { clientX: 250, clientY: 100 }
        ]
    });
    assert.match(elements.cadOverlay.style.transform, /scale\(2\)/);
    assert.doesNotMatch(elements.mapZoomStage.style.transform, /scale\(2\)/);
    assert.equal(elements.detailZoomBtn.textContent, '상세확대');
    assert.equal(getMap().draggable, false);
    assert.equal(getMap().zoomable, true);
});

test('detail zoom applies the inverse scale used to keep labels the same size', async () => {
    const { elements } = await createHarness();
    elements.detailZoomBtn.listeners.click();
    assert.equal(elements.cadOverlay.style['--cad-label-inverse-scale'], '0.5');
    elements.detailZoomBtn.listeners.click();
    assert.equal(elements.cadOverlay.style['--cad-label-inverse-scale'], '0.25');
});
