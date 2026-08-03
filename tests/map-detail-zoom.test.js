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
        roadmapBtn: element(),
        skyviewBtn: element(),
        recenterMapBtn: element(),
        currentLocationBtn: element(),
        displaySettingsBtn: element(),
        detailZoomBtn: element(),
        orientationModeBtn: element(),
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
        relayout() {}
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
        addEventListener() {},
        navigator: {},
        innerWidth: 600,
        innerHeight: 1000,
        screen: {
            orientation: {
                lock: async () => {}
            }
        }
    };
    const document = {
        getElementById: (id) => elements[id] || null,
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
    return { elements, getMap: () => mapInstance };
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

test('orientation button toggles landscape label mode without rotating the map', async () => {
    const { elements, getMap } = await createHarness();

    await elements.orientationModeBtn.listeners.click();
    assert.equal(elements.mapView.classList.contains('landscape-mode'), true);
    assert.equal(elements.orientationModeBtn['aria-pressed'], 'true');
    assert.equal(elements.cadOverlay.style['--cad-label-rotation'], '-90deg');
    assert.doesNotMatch(elements.mapZoomStage.style.transform || '', /rotate/);
    assert.equal(getMap().draggable, true);

    await elements.orientationModeBtn.listeners.click();
    assert.equal(elements.mapView.classList.contains('landscape-mode'), false);
    assert.equal(elements.cadOverlay.style['--cad-label-rotation'], '0deg');
});

test('detail zoom applies the inverse scale used to keep labels the same size', async () => {
    const { elements } = await createHarness();
    elements.detailZoomBtn.listeners.click();
    assert.equal(elements.cadOverlay.style['--cad-label-inverse-scale'], '0.5');
    elements.detailZoomBtn.listeners.click();
    assert.equal(elements.cadOverlay.style['--cad-label-inverse-scale'], '0.25');
});
