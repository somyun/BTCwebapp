(function () {
    'use strict';

    const KAKAO_JAVASCRIPT_KEY = '708065ee6e872ac3f158928a61d3252e';
    const CAD_MANIFEST_URL = './cad-data/hopo/manifest.json';
    const CORE_LAYER_NAMES = new Set(['0', 'SIMPLE', 'CABLE', '신설', '신설1', 'WALL', '전주']);
    const OVERLAY_PADDING_RATIO = 0.015;

    const layerCache = new Map();
    const selectedLayers = new Set();

    let sdkPromise = null;
    let manifestPromise = null;
    let manifest = null;
    let manifestUrl = null;
    let map = null;
    let canvas = null;
    let context = null;
    let controlsBound = false;
    let layerListBuilt = false;
    let initialSelectionApplied = false;
    let initialBoundsApplied = false;
    let currentMapType = 'skyview';
    let positionFrame = 0;
    let idleTimer = 0;
    let rasterDirty = true;
    let renderedLevel = null;
    let currentPositionMarker = null;
    let currentPositionAccuracy = null;

    function getElement(id) {
        return document.getElementById(id);
    }

    function setLayerStatus(text, tone = 'ready') {
        const badge = getElement('cadLayerStatus');
        if (!badge) return;
        badge.textContent = text;
        badge.dataset.tone = tone;
    }

    function setLayerMessage(text, isError = false) {
        const message = getElement('cadLayerMessage');
        if (!message) return;
        message.textContent = text;
        message.classList.toggle('error', isError);
    }

    function loadKakaoMapSdk() {
        if (window.kakao && window.kakao.maps) {
            return new Promise((resolve) => window.kakao.maps.load(resolve));
        }

        if (sdkPromise) return sdkPromise;

        sdkPromise = new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${KAKAO_JAVASCRIPT_KEY}&autoload=false`;
            script.async = true;
            script.addEventListener('load', () => {
                if (!window.kakao || !window.kakao.maps) {
                    reject(new Error('카카오 지도 SDK를 시작하지 못했습니다.'));
                    return;
                }
                window.kakao.maps.load(resolve);
            });
            script.addEventListener('error', () => {
                reject(new Error('카카오 지도 SDK를 내려받지 못했습니다.'));
            });
            document.head.appendChild(script);
        });

        return sdkPromise;
    }

    function loadManifest() {
        if (manifestPromise) return manifestPromise;

        manifestPromise = fetch(CAD_MANIFEST_URL, { cache: 'no-cache' }).then(async (response) => {
            if (!response.ok) throw new Error('호포 도면 정보를 불러오지 못했습니다.');
            manifestUrl = response.url;
            return response.json();
        });
        return manifestPromise;
    }

    function loadLayer(layer) {
        if (layerCache.has(layer.id)) return layerCache.get(layer.id);

        const url = new URL(layer.file, manifestUrl).href;
        const request = fetch(url, { cache: 'force-cache' }).then((response) => {
            if (!response.ok) throw new Error(`${layer.name} 레이어를 불러오지 못했습니다.`);
            return response.json();
        });
        layerCache.set(layer.id, request);
        return request;
    }

    function setMapType(type) {
        if (!map) return;

        currentMapType = type === 'roadmap' ? 'roadmap' : 'skyview';
        map.setMapTypeId(currentMapType === 'skyview'
            ? window.kakao.maps.MapTypeId.SKYVIEW
            : window.kakao.maps.MapTypeId.ROADMAP);

        getElement('roadmapBtn')?.classList.toggle('active', currentMapType === 'roadmap');
        getElement('skyviewBtn')?.classList.toggle('active', currentMapType === 'skyview');
        rasterDirty = true;
        renderRaster();
    }

    function fitToDepot() {
        if (!map || !manifest) return;

        const [west, south, east, north] = manifest.bounds_wgs84;
        const bounds = new window.kakao.maps.LatLngBounds();
        bounds.extend(new window.kakao.maps.LatLng(south, west));
        bounds.extend(new window.kakao.maps.LatLng(north, east));
        map.setBounds(bounds, 42, 42, 42, 42);
    }

    function setLocationStatus(text, tone = '') {
        const status = getElement('mapLocationStatus');
        if (!status) return;
        status.textContent = text;
        status.dataset.tone = tone;
        status.classList.toggle('visible', Boolean(text));
        window.clearTimeout(setLocationStatus.timer);
        if (text) {
            setLocationStatus.timer = window.setTimeout(() => {
                status.classList.remove('visible');
            }, 5000);
        }
    }

    function geolocationErrorMessage(error) {
        if (error?.code === 1) return '위치 권한이 차단되었습니다. 브라우저 설정에서 위치 권한을 허용해 주세요.';
        if (error?.code === 2) return '현재 위치를 확인할 수 없습니다. GPS 또는 네트워크 상태를 확인해 주세요.';
        if (error?.code === 3) return '위치 확인 시간이 초과되었습니다. 다시 시도해 주세요.';
        return '현재 위치를 가져오지 못했습니다.';
    }

    function showCurrentPosition() {
        const button = getElement('currentLocationBtn');
        if (!map || !navigator.geolocation || !button) {
            setLocationStatus('이 브라우저에서는 현재 위치 기능을 사용할 수 없습니다.', 'error');
            return;
        }

        button.disabled = true;
        button.textContent = '위치 확인 중…';
        setLocationStatus('현재 위치 권한을 확인하고 있습니다.', 'loading');

        navigator.geolocation.getCurrentPosition((position) => {
            const latitude = position.coords.latitude;
            const longitude = position.coords.longitude;
            const accuracy = Math.max(1, position.coords.accuracy || 1);
            const latLng = new window.kakao.maps.LatLng(latitude, longitude);

            if (!currentPositionAccuracy) {
                currentPositionAccuracy = new window.kakao.maps.Circle({
                    map,
                    center: latLng,
                    radius: accuracy,
                    strokeWeight: 1,
                    strokeColor: '#1677ff',
                    strokeOpacity: 0.75,
                    fillColor: '#1677ff',
                    fillOpacity: 0.14
                });
            } else {
                currentPositionAccuracy.setMap(map);
                currentPositionAccuracy.setPosition(latLng);
                currentPositionAccuracy.setRadius(accuracy);
            }

            if (!currentPositionMarker) {
                currentPositionMarker = new window.kakao.maps.Marker({
                    map,
                    position: latLng,
                    title: '현재 위치'
                });
            } else {
                currentPositionMarker.setMap(map);
                currentPositionMarker.setPosition(latLng);
            }

            map.setLevel(3);
            map.panTo(latLng);
            setLocationStatus(`현재 위치로 이동했습니다. 정확도 약 ${Math.round(accuracy)}m`, 'ready');
            button.disabled = false;
            button.textContent = '현재 위치';
        }, (error) => {
            setLocationStatus(geolocationErrorMessage(error), 'error');
            button.disabled = false;
            button.textContent = '현재 위치';
        }, {
            enableHighAccuracy: true,
            timeout: 12000,
            maximumAge: 15000
        });
    }

    function overlayGeoBounds() {
        const [west, south, east, north] = manifest.bounds_wgs84;
        const lonPadding = (east - west) * OVERLAY_PADDING_RATIO;
        const latPadding = (north - south) * OVERLAY_PADDING_RATIO;
        return [
            west - lonPadding,
            south - latPadding,
            east + lonPadding,
            north + latPadding
        ];
    }

    function overlayScreenBox() {
        const [west, south, east, north] = overlayGeoBounds();
        const projection = map.getProjection();
        const corners = [
            [west, north],
            [east, north],
            [east, south],
            [west, south]
        ].map(([longitude, latitude]) => projection.containerPointFromCoords(
            new window.kakao.maps.LatLng(latitude, longitude)
        ));
        const xs = corners.map((point) => point.x);
        const ys = corners.map((point) => point.y);
        const left = Math.min(...xs);
        const right = Math.max(...xs);
        const top = Math.min(...ys);
        const bottom = Math.max(...ys);

        return {
            left,
            top,
            width: Math.max(1, right - left),
            height: Math.max(1, bottom - top)
        };
    }

    function updateOverlayPosition() {
        if (!map || !manifest || !canvas) return;
        const box = overlayScreenBox();
        canvas.style.left = `${box.left}px`;
        canvas.style.top = `${box.top}px`;
        canvas.style.width = `${box.width}px`;
        canvas.style.height = `${box.height}px`;
    }

    function queuePositionUpdate() {
        if (positionFrame) return;
        positionFrame = window.requestAnimationFrame(() => {
            positionFrame = 0;
            updateOverlayPosition();
        });
    }

    function chooseRasterSize(box) {
        const aspect = Math.max(0.05, Math.min(2, box.width / Math.max(box.height, 1)));
        const deviceScale = window.devicePixelRatio || 1;
        let height = Math.max(2048, Math.min(8192, Math.round(box.height * deviceScale * 1.35)));
        let width = Math.max(512, Math.round(height * aspect));

        if (width > 4096) {
            const factor = 4096 / width;
            width = 4096;
            height = Math.round(height * factor);
        }
        return { width, height };
    }

    function displayColor(color) {
        const normalized = String(color || '').toLowerCase();
        if (normalized === '#000000' || normalized === '#ffffff') {
            return currentMapType === 'skyview' ? '#f8fafc' : '#0f172a';
        }
        return color || '#00d9e8';
    }

    async function renderRaster(force = false) {
        if (!map || !manifest || !canvas || !context || (!rasterDirty && !force)) return;

        const box = overlayScreenBox();
        updateOverlayPosition();
        const size = chooseRasterSize(box);
        if (canvas.width !== size.width || canvas.height !== size.height) {
            canvas.width = size.width;
            canvas.height = size.height;
        }

        context.clearRect(0, 0, canvas.width, canvas.height);
        context.lineJoin = 'round';
        context.lineCap = 'round';

        const projection = map.getProjection();
        const scaleX = canvas.width / Math.max(box.width, 1);
        const scaleY = canvas.height / Math.max(box.height, 1);
        const screenScale = Math.min(scaleX, scaleY);
        const toPixel = ([longitude, latitude]) => {
            const point = projection.containerPointFromCoords(
                new window.kakao.maps.LatLng(latitude, longitude)
            );
            return [
                (point.x - box.left) * scaleX,
                (point.y - box.top) * scaleY
            ];
        };
        const showLabels = Boolean(getElement('cadLabelToggle')?.checked);

        for (const layerInfo of manifest.layers) {
            if (!selectedLayers.has(layerInfo.id)) continue;
            const layer = await loadLayer(layerInfo);
            context.strokeStyle = displayColor(layer.color);
            context.fillStyle = context.strokeStyle;
            context.lineWidth = Math.max(1, screenScale);
            context.beginPath();

            for (const path of layer.paths) {
                let first = true;
                for (const coordinate of path) {
                    const [x, y] = toPixel(coordinate);
                    if (first) {
                        context.moveTo(x, y);
                        first = false;
                    } else {
                        context.lineTo(x, y);
                    }
                }
            }
            context.stroke();

            const pointSize = Math.max(2, screenScale * 2);
            for (const coordinate of layer.points) {
                const [x, y] = toPixel(coordinate);
                context.fillRect(x - pointSize / 2, y - pointSize / 2, pointSize, pointSize);
            }

            if (showLabels) {
                context.font = `${Math.max(10, 11 * screenScale)}px "Malgun Gothic", sans-serif`;
                for (const label of layer.labels) {
                    const [x, y] = toPixel(label.position);
                    context.fillText(label.text, x, y);
                }
            }
        }

        rasterDirty = false;
        renderedLevel = map.getLevel();
    }

    function updateSelectedStatus() {
        setLayerStatus(`${selectedLayers.size}개 선택`, 'ready');
        setLayerMessage('레이어를 선택해 도면을 켜고 끌 수 있습니다.');
    }

    async function applySelection(predicate) {
        selectedLayers.clear();
        const loads = [];

        for (const input of document.querySelectorAll('input[data-cad-layer]')) {
            const layer = manifest.layers.find((item) => item.id === input.dataset.cadLayer);
            const checked = Boolean(layer && predicate(layer));
            input.checked = checked;
            if (checked) {
                selectedLayers.add(layer.id);
                loads.push(loadLayer(layer));
            }
        }

        setLayerStatus(`${selectedLayers.size}개 불러오는 중`, 'loading');
        await Promise.all(loads);
        rasterDirty = true;
        await renderRaster();
        updateSelectedStatus();
    }

    function buildLayerList() {
        if (layerListBuilt) return;
        const root = getElement('cadLayerList');
        if (!root) return;
        root.replaceChildren();

        for (const layer of manifest.layers) {
            const label = document.createElement('label');
            label.className = 'cad-layer-item';

            const input = document.createElement('input');
            input.type = 'checkbox';
            input.dataset.cadLayer = layer.id;

            const swatch = document.createElement('span');
            swatch.className = 'cad-layer-swatch';
            swatch.style.backgroundColor = layer.color;

            const name = document.createElement('span');
            name.className = 'cad-layer-name';
            name.textContent = layer.name;

            const count = document.createElement('span');
            count.className = 'cad-layer-count';
            count.textContent = layer.path_count.toLocaleString();

            label.append(input, swatch, name, count);
            root.appendChild(label);

            input.addEventListener('change', async () => {
                input.disabled = true;
                try {
                    if (input.checked) {
                        selectedLayers.add(layer.id);
                        setLayerStatus(`${layer.name} 불러오는 중`, 'loading');
                        await loadLayer(layer);
                    } else {
                        selectedLayers.delete(layer.id);
                    }
                    rasterDirty = true;
                    await renderRaster();
                    updateSelectedStatus();
                } catch (error) {
                    input.checked = false;
                    selectedLayers.delete(layer.id);
                    setLayerStatus('오류', 'error');
                    setLayerMessage(error.message, true);
                } finally {
                    input.disabled = false;
                }
            });
        }
        layerListBuilt = true;
    }

    function onZoomStart() {
        window.clearTimeout(idleTimer);
        canvas?.classList.add('zooming');
        queuePositionUpdate();
    }

    function onMapIdle() {
        queuePositionUpdate();
        window.clearTimeout(idleTimer);
        idleTimer = window.setTimeout(async () => {
            canvas?.classList.remove('zooming');
            if (rasterDirty || renderedLevel !== map.getLevel()) {
                await renderRaster(true);
            } else {
                updateOverlayPosition();
            }
        }, 330);
    }

    function bindControls() {
        if (controlsBound) return;
        getElement('roadmapBtn')?.addEventListener('click', () => setMapType('roadmap'));
        getElement('skyviewBtn')?.addEventListener('click', () => setMapType('skyview'));
        getElement('recenterMapBtn')?.addEventListener('click', fitToDepot);
        getElement('currentLocationBtn')?.addEventListener('click', showCurrentPosition);
        getElement('cadCoreLayersBtn')?.addEventListener('click', () => applySelection(
            (layer) => CORE_LAYER_NAMES.has(layer.name)
        ));
        getElement('cadAllLayersBtn')?.addEventListener('click', () => applySelection(() => true));
        getElement('cadNoLayersBtn')?.addEventListener('click', () => applySelection(() => false));
        getElement('cadLabelToggle')?.addEventListener('change', async () => {
            rasterDirty = true;
            await renderRaster();
        });
        getElement('cadOpacity')?.addEventListener('input', (event) => {
            if (canvas) canvas.style.opacity = String(Number(event.target.value) / 100);
        });
        window.addEventListener('resize', async () => {
            if (!map) return;
            window.kakao.maps.event.trigger(map, 'resize');
            updateOverlayPosition();
            await renderRaster(true);
        });
        controlsBound = true;
    }

    function bindMapEvents() {
        window.kakao.maps.event.addListener(map, 'zoom_start', onZoomStart);
        window.kakao.maps.event.addListener(map, 'zoom_changed', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'center_changed', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'bounds_changed', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'dragstart', () => canvas?.classList.remove('zooming'));
        window.kakao.maps.event.addListener(map, 'drag', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'idle', onMapIdle);
    }

    async function initialize() {
        const loading = getElement('mapLoading');

        try {
            const results = await Promise.all([loadKakaoMapSdk(), loadManifest()]);
            manifest = results[1];
            canvas = getElement('cadOverlay');
            context = canvas?.getContext('2d') || null;
            if (!canvas || !context) throw new Error('도면 표시 화면을 준비하지 못했습니다.');

            if (!map) {
                const [longitude, latitude] = manifest.center_wgs84;
                map = new window.kakao.maps.Map(getElement('kakaoMap'), {
                    center: new window.kakao.maps.LatLng(latitude, longitude),
                    level: 4
                });
                const zoomControl = new window.kakao.maps.ZoomControl();
                map.addControl(zoomControl, window.kakao.maps.ControlPosition.RIGHT);
                bindMapEvents();
            }

            buildLayerList();
            bindControls();
            map.relayout();
            setMapType(currentMapType);

            if (!initialBoundsApplied) {
                fitToDepot();
                initialBoundsApplied = true;
            } else {
                queuePositionUpdate();
            }

            if (!initialSelectionApplied) {
                initialSelectionApplied = true;
                await applySelection((layer) => CORE_LAYER_NAMES.has(layer.name));
            } else {
                await renderRaster(true);
            }

            canvas.style.opacity = String(Number(getElement('cadOpacity')?.value || 80) / 100);
            loading?.classList.add('hidden');
        } catch (error) {
            console.error('Kakao CAD map initialization failed:', error);
            setLayerStatus('오류', 'error');
            setLayerMessage(error.message, true);
            if (loading) {
                loading.textContent = error.message || '지도를 불러오지 못했습니다.';
                loading.classList.add('error');
            }
        }
    }

    window.BWAMap = {
        initialize,
        relayout() {
            if (!map) return;
            map.relayout();
            queuePositionUpdate();
        }
    };
}());
