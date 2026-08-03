(function () {
    'use strict';

    const KAKAO_JAVASCRIPT_KEY = '708065ee6e872ac3f158928a61d3252e';
    const CAD_MANIFEST_URL = './cad-data/hopo/manifest.json';
    const CORE_LAYER_NAMES = new Set(['0', 'SIMPLE', 'CABLE', '신설', '신설1', 'WALL', '전주']);
    const OVERLAY_PADDING_RATIO = 0.015;
    const DETAIL_ZOOM_STEPS = [1, 2, 4, 8];

    const layerCache = new Map();
    const selectedLayers = new Set();

    let sdkPromise = null;
    let manifestPromise = null;
    let manifest = null;
    let manifestUrl = null;
    let map = null;
    let canvas = null;
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
    let pinchGesture = null;
    let pinchFinishTimer = 0;
    let detailScale = 1;
    let detailOffset = { x: 0, y: 0 };
    let detailPanGesture = null;
    let detailTransitionTimer = 0;
    let detailInteractionLocked = false;

    function getElement(id) {
        return document.getElementById(id);
    }

    function minimumMapLevel() {
        return currentMapType === 'skyview' ? 0 : 1;
    }

    function customTransformActive() {
        return detailScale > 1.001;
    }

    function clampDetailOffset(scale = detailScale, offset = detailOffset) {
        const view = getElement('mapView');
        if (!view || scale <= 1) return { x: 0, y: 0 };
        const maxX = ((scale - 1) * view.clientWidth) / 2;
        const maxY = ((scale - 1) * view.clientHeight) / 2;
        return {
            x: Math.max(-maxX, Math.min(maxX, offset.x)),
            y: Math.max(-maxY, Math.min(maxY, offset.y))
        };
    }

    function updateZoomControls() {
        const detailButton = getElement('detailZoomBtn');
        const detailActive = detailScale > 1.001;
        if (detailButton) {
            const displayScale = Number.isInteger(detailScale)
                ? detailScale.toFixed(0)
                : detailScale.toFixed(1);
            detailButton.textContent = detailActive ? `상세 ${displayScale}×` : '상세확대';
            detailButton.classList.toggle('active', detailActive);
            detailButton.setAttribute('aria-pressed', String(detailActive));
        }

        const zoomInButton = getElement('zoomInBtn');
        const zoomOutButton = getElement('zoomOutBtn');
        if (zoomInButton) zoomInButton.disabled = detailScale >= DETAIL_ZOOM_STEPS[DETAIL_ZOOM_STEPS.length - 1] - 0.001;
        if (zoomOutButton) zoomOutButton.disabled = detailScale <= 1.001 && map?.getLevel() >= 14;
    }

    function applyDetailTransform(animate = false) {
        const stage = getElement('mapZoomStage');
        if (!stage) return;

        const scale = detailScale;
        detailOffset = clampDetailOffset();
        const active = customTransformActive();
        stage.classList.toggle('detail-mode', active);
        stage.classList.toggle('detail-transition', animate);
        stage.style.transform = active
            ? `translate3d(${detailOffset.x}px, ${detailOffset.y}px, 0) scale(${scale})`
            : '';

        if (canvas) {
            const inverseScale = String(1 / detailScale);
            if (canvas.style.setProperty) {
                canvas.style.setProperty('--cad-label-inverse-scale', inverseScale);
            } else {
                canvas.style['--cad-label-inverse-scale'] = inverseScale;
            }
        }

        window.clearTimeout(detailTransitionTimer);
        if (animate) {
            detailTransitionTimer = window.setTimeout(() => {
                stage.classList.remove('detail-transition');
            }, 240);
        }

        if (map && active !== detailInteractionLocked) {
            map.setDraggable(!active);
            map.setZoomable(!active);
            detailInteractionLocked = active;
        }
        updateZoomControls();
    }

    function setDetailScale(nextScale, anchor = null, animate = false) {
        const view = getElement('mapView');
        const oldScale = detailScale;
        const boundedScale = Math.max(1, Math.min(DETAIL_ZOOM_STEPS[DETAIL_ZOOM_STEPS.length - 1], nextScale));

        if (view && anchor && oldScale > 0) {
            const rect = view.getBoundingClientRect();
            const centerX = rect.left + (rect.width / 2);
            const centerY = rect.top + (rect.height / 2);
            const ratio = boundedScale / oldScale;
            detailOffset = {
                x: anchor.x - centerX - (ratio * (anchor.x - centerX - detailOffset.x)),
                y: anchor.y - centerY - (ratio * (anchor.y - centerY - detailOffset.y))
            };
        }

        detailScale = boundedScale;
        if (detailScale <= 1.001) {
            detailScale = 1;
            detailOffset = { x: 0, y: 0 };
        }
        applyDetailTransform(animate);
    }

    function resetDetailZoom(animate = false) {
        setDetailScale(1, null, animate);
    }

    function resetViewTransform(animate = false) {
        detailScale = 1;
        detailOffset = { x: 0, y: 0 };
        applyDetailTransform(animate);
    }

    function screenDeltaToMapDelta(deltaX, deltaY) {
        const scale = detailScale;
        return {
            x: deltaX / scale,
            y: deltaY / scale
        };
    }

    function normalizedScreenAngle() {
        const screenAngle = Number(window.screen?.orientation?.angle);
        if (Number.isFinite(screenAngle)) return ((screenAngle % 360) + 360) % 360;

        const legacyAngle = Number(window.orientation);
        if (Number.isFinite(legacyAngle)) return ((legacyAngle % 360) + 360) % 360;
        return null;
    }

    function detectedLabelRotation(viewportLandscape) {
        if (!viewportLandscape) return 0;
        const angle = normalizedScreenAngle();
        if (angle === 270) return 90;
        if (angle === 90) return -90;
        return -90;
    }

    function syncOrientationFromDevice() {
        const view = getElement('mapView');
        const viewportLandscape = window.innerWidth > window.innerHeight;
        view?.classList.toggle('landscape-mode', viewportLandscape);
        if (canvas) {
            const labelRotation = `${detectedLabelRotation(viewportLandscape)}deg`;
            if (canvas.style.setProperty) {
                canvas.style.setProperty('--cad-label-rotation', labelRotation);
            } else {
                canvas.style['--cad-label-rotation'] = labelRotation;
            }
        }
    }

    function panTransformedMap(deltaX, deltaY) {
        if (!map || (!deltaX && !deltaY)) return;
        const projection = map.getProjection();
        const centerPoint = projection.containerPointFromCoords(map.getCenter());
        const delta = screenDeltaToMapDelta(deltaX, deltaY);
        const targetPoint = new window.kakao.maps.Point(
            centerPoint.x - delta.x,
            centerPoint.y - delta.y
        );
        map.setCenter(projection.coordsFromContainerPoint(targetPoint));
    }

    function commitDetailOffsetToMap() {
        if (Math.abs(detailOffset.x) < 0.01 && Math.abs(detailOffset.y) < 0.01) return;
        const offset = { ...detailOffset };
        detailOffset = { x: 0, y: 0 };
        panTransformedMap(offset.x, offset.y);
        applyDetailTransform();
    }

    function cycleDetailZoom() {
        if (!map) return;
        if (map.getLevel() !== minimumMapLevel()) map.setLevel(minimumMapLevel());
        const next = detailScale < 1.5 ? 2 : detailScale < 3 ? 4 : detailScale < 6 ? 8 : 1;
        setDetailScale(next, null, true);
    }

    function zoomIn() {
        if (!map) return;
        const level = map.getLevel();
        if (detailScale > 1.001 || level <= minimumMapLevel()) {
            const next = detailScale < 1.5 ? 2 : detailScale < 3 ? 4 : 8;
            setDetailScale(next, null, true);
            return;
        }
        map.setLevel(level - 1, { animate: true });
    }

    function zoomOut() {
        if (!map) return;
        if (detailScale > 1.001) {
            const next = detailScale > 6 ? 4 : detailScale > 3 ? 2 : 1;
            setDetailScale(next, null, true);
            return;
        }
        map.setLevel(Math.min(14, map.getLevel() + 1), { animate: true });
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

        resetViewTransform();
        currentMapType = type === 'roadmap' ? 'roadmap' : 'skyview';
        map.setMapTypeId(currentMapType === 'skyview'
            ? window.kakao.maps.MapTypeId.SKYVIEW
            : window.kakao.maps.MapTypeId.ROADMAP);

        const toggleButton = getElement('mapTypeToggleBtn');
        if (toggleButton) {
            const isSkyview = currentMapType === 'skyview';
            toggleButton.textContent = isSkyview ? '위성지도' : '일반지도';
            toggleButton.classList.toggle('active', isSkyview);
            toggleButton.setAttribute('aria-pressed', String(isSkyview));
            toggleButton.title = isSkyview ? '일반지도로 전환' : '위성지도로 전환';
        }
        rasterDirty = true;
        renderRaster();
    }

    function toggleMapType() {
        setMapType(currentMapType === 'skyview' ? 'roadmap' : 'skyview');
    }

    function fitToDepot() {
        if (!map || !manifest) return;

        resetViewTransform();
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

        resetViewTransform();
        button.disabled = true;
        button.setAttribute('aria-label', '현재 위치 확인 중');
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
            button.setAttribute('aria-label', '현재 위치로 이동');
        }, (error) => {
            setLocationStatus(geolocationErrorMessage(error), 'error');
            button.disabled = false;
            button.setAttribute('aria-label', '현재 위치로 이동');
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
        if (canvas?.classList.contains('pinching')) return;
        if (positionFrame) return;
        positionFrame = window.requestAnimationFrame(() => {
            positionFrame = 0;
            updateOverlayPosition();
        });
    }

    function touchDistance(first, second) {
        return Math.hypot(second.clientX - first.clientX, second.clientY - first.clientY);
    }

    function touchMidpoint(first, second) {
        return {
            x: (first.clientX + second.clientX) / 2,
            y: (first.clientY + second.clientY) / 2
        };
    }

    function beginPinch(event) {
        if (!canvas) return;

        if (customTransformActive() && event.touches.length === 1) {
            const touch = event.touches[0];
            detailPanGesture = {
                x: touch.clientX,
                y: touch.clientY
            };
            getElement('mapZoomStage')?.classList.add('dragging');
            return;
        }

        if (event.touches.length !== 2) return;
        const first = event.touches[0];
        const second = event.touches[1];
        const midpoint = touchMidpoint(first, second);

        window.clearTimeout(pinchFinishTimer);
        detailPanGesture = null;
        if (customTransformActive()) {
            pinchGesture = {
                mode: 'detail',
                distance: Math.max(1, touchDistance(first, second)),
                midpoint,
                scale: detailScale,
                offset: { ...detailOffset }
            };
            getElement('mapZoomStage')?.classList.add('dragging');
            return;
        }

        const canvasRect = canvas.getBoundingClientRect();
        pinchGesture = {
            mode: 'map',
            distance: Math.max(1, touchDistance(first, second)),
            midpoint
        };
        canvas.classList.add('pinching');
        canvas.classList.remove('zooming');
        canvas.style.transformOrigin = `${midpoint.x - canvasRect.left}px ${midpoint.y - canvasRect.top}px`;
        canvas.style.transform = 'translate3d(0, 0, 0) scale(1)';
    }

    function updatePinch(event) {
        if (detailPanGesture && event.touches.length === 1 && customTransformActive()) {
            const touch = event.touches[0];
            panTransformedMap(
                touch.clientX - detailPanGesture.x,
                touch.clientY - detailPanGesture.y
            );
            detailPanGesture.x = touch.clientX;
            detailPanGesture.y = touch.clientY;
            return;
        }

        if (!pinchGesture || event.touches.length !== 2 || !canvas) return;
        const first = event.touches[0];
        const second = event.touches[1];
        const midpoint = touchMidpoint(first, second);
        const distanceRatio = touchDistance(first, second) / pinchGesture.distance;

        if (pinchGesture.mode === 'detail') {
            detailScale = Math.max(1, Math.min(DETAIL_ZOOM_STEPS[DETAIL_ZOOM_STEPS.length - 1], pinchGesture.scale * distanceRatio));
            const viewRect = getElement('mapView')?.getBoundingClientRect();
            if (viewRect) {
                const centerX = viewRect.left + (viewRect.width / 2);
                const centerY = viewRect.top + (viewRect.height / 2);
                const ratio = detailScale / pinchGesture.scale;
                detailOffset = {
                    x: midpoint.x - centerX - (ratio * (pinchGesture.midpoint.x - centerX - pinchGesture.offset.x)),
                    y: midpoint.y - centerY - (ratio * (pinchGesture.midpoint.y - centerY - pinchGesture.offset.y))
                };
            }
            applyDetailTransform();
            return;
        }

        const scale = Math.max(0.25, Math.min(4, distanceRatio));
        const translateX = midpoint.x - pinchGesture.midpoint.x;
        const translateY = midpoint.y - pinchGesture.midpoint.y;
        canvas.style.transform = `translate3d(${translateX}px, ${translateY}px, 0) scale(${scale})`;
    }

    function clearPinchTransform() {
        if (!canvas) return;
        pinchGesture = null;
        canvas.classList.remove('pinching');
        canvas.style.transform = '';
        canvas.style.transformOrigin = '';
        updateOverlayPosition();
    }

    function endPinch(event) {
        if (detailPanGesture && event.touches.length === 0) {
            detailPanGesture = null;
            getElement('mapZoomStage')?.classList.remove('dragging');
        }
        if (!pinchGesture || event.touches.length >= 2) return;

        if (pinchGesture.mode === 'detail') {
            pinchGesture = null;
            if (event.touches.length === 1) {
                const touch = event.touches[0];
                detailPanGesture = {
                    x: touch.clientX,
                    y: touch.clientY
                };
            } else {
                getElement('mapZoomStage')?.classList.remove('dragging');
            }
            commitDetailOffsetToMap();
            if (detailScale < 1.15) resetDetailZoom(true);
            else applyDetailTransform();
            return;
        }

        pinchGesture = null;
        window.clearTimeout(pinchFinishTimer);
        // Give Kakao Maps a moment to commit the new projection, then hand control
        // back quickly so a remaining finger can continue panning without lag.
        pinchFinishTimer = window.setTimeout(clearPinchTransform, 80);
    }

    function beginDetailPointerPan(event) {
        if (!customTransformActive() || event.pointerType === 'touch' || event.button !== 0) return;
        const stage = getElement('mapZoomStage');
        detailPanGesture = {
            pointerId: event.pointerId,
            x: event.clientX,
            y: event.clientY
        };
        stage?.setPointerCapture?.(event.pointerId);
        stage?.classList.add('dragging');
        event.preventDefault();
    }

    function updateDetailPointerPan(event) {
        if (!detailPanGesture || detailPanGesture.pointerId !== event.pointerId) return;
        panTransformedMap(
            event.clientX - detailPanGesture.x,
            event.clientY - detailPanGesture.y
        );
        detailPanGesture.x = event.clientX;
        detailPanGesture.y = event.clientY;
        event.preventDefault();
    }

    function endDetailPointerPan(event) {
        if (!detailPanGesture || detailPanGesture.pointerId !== event.pointerId) return;
        detailPanGesture = null;
        getElement('mapZoomStage')?.classList.remove('dragging');
    }

    function detailWheelZoom(event) {
        if (!customTransformActive()) return;
        event.preventDefault();
        const factor = Math.exp(-event.deltaY * 0.0015);
        setDetailScale(detailScale * factor, { x: event.clientX, y: event.clientY });
    }

    function displayColor(color) {
        const normalized = String(color || '').toLowerCase();
        if (normalized === '#000000' || normalized === '#ffffff') {
            return currentMapType === 'skyview' ? '#f8fafc' : '#0f172a';
        }
        return color || '#00d9e8';
    }

    function createSvgElement(name, attributes = {}) {
        const element = document.createElementNS('http://www.w3.org/2000/svg', name);
        for (const [key, value] of Object.entries(attributes)) {
            element.setAttribute(key, String(value));
        }
        return element;
    }

    async function renderRaster(force = false) {
        if (!map || !manifest || !canvas || (!rasterDirty && !force)) return;

        const box = overlayScreenBox();
        updateOverlayPosition();
        const width = Math.max(1, box.width);
        const height = Math.max(1, box.height);
        canvas.setAttribute('viewBox', `0 0 ${width} ${height}`);

        const projection = map.getProjection();
        const toPixel = ([longitude, latitude]) => {
            const point = projection.containerPointFromCoords(
                new window.kakao.maps.LatLng(latitude, longitude)
            );
            return [
                point.x - box.left,
                point.y - box.top
            ];
        };
        const showLabels = Boolean(getElement('cadLabelToggle')?.checked);
        const fragment = document.createDocumentFragment();

        for (const layerInfo of manifest.layers) {
            if (!selectedLayers.has(layerInfo.id)) continue;
            const layer = await loadLayer(layerInfo);
            const color = displayColor(layer.color);
            const commands = [];

            for (const path of layer.paths) {
                for (let index = 0; index < path.length; index += 1) {
                    const coordinate = path[index];
                    const [x, y] = toPixel(coordinate);
                    commands.push(`${index === 0 ? 'M' : 'L'}${x.toFixed(2)} ${y.toFixed(2)}`);
                }
            }
            if (commands.length) {
                fragment.appendChild(createSvgElement('path', {
                    d: commands.join(' '),
                    fill: 'none',
                    stroke: color,
                    'stroke-width': 1.2,
                    'stroke-linecap': 'round',
                    'stroke-linejoin': 'round',
                    'vector-effect': 'non-scaling-stroke'
                }));
            }

            for (const coordinate of layer.points) {
                const [x, y] = toPixel(coordinate);
                fragment.appendChild(createSvgElement('rect', {
                    x: x - 1.5,
                    y: y - 1.5,
                    width: 3,
                    height: 3,
                    fill: color
                }));
            }

            if (showLabels) {
                for (const label of layer.labels) {
                    const [x, y] = toPixel(label.position);
                    const text = createSvgElement('text', {
                        x,
                        y,
                        class: 'cad-map-label',
                        fill: color,
                        'font-size': 11,
                        'font-family': 'Malgun Gothic, sans-serif',
                        style: `--cad-label-origin: ${x}px ${y}px`
                    });
                    text.textContent = label.text;
                    fragment.appendChild(text);
                }
            }
        }

        canvas.replaceChildren(fragment);
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
        if (!canvas?.classList.contains('pinching')) canvas?.classList.add('zooming');
        queuePositionUpdate();
    }

    function onMapIdle() {
        if (canvas?.classList.contains('pinching')) clearPinchTransform();
        queuePositionUpdate();
        window.clearTimeout(idleTimer);
        idleTimer = window.setTimeout(async () => {
            canvas?.classList.remove('zooming');
            if (rasterDirty || renderedLevel !== map.getLevel()) {
                await renderRaster(true);
            } else {
                updateOverlayPosition();
            }
            updateZoomControls();
        }, 330);
    }

    function onMapZoomChanged() {
        queuePositionUpdate();
        updateZoomControls();
    }

    function bindControls() {
        if (controlsBound) return;
        getElement('mapTypeToggleBtn')?.addEventListener('click', toggleMapType);
        getElement('currentLocationBtn')?.addEventListener('click', showCurrentPosition);
        getElement('detailZoomBtn')?.addEventListener('click', cycleDetailZoom);
        getElement('zoomInBtn')?.addEventListener('click', zoomIn);
        getElement('zoomOutBtn')?.addEventListener('click', zoomOut);
        getElement('displaySettingsBtn')?.addEventListener('click', () => {
            const panel = getElement('cadLayerPanel');
            const button = getElement('displaySettingsBtn');
            if (!panel || !button) return;
            const willOpen = panel.hidden;
            panel.hidden = !willOpen;
            button.classList.toggle('active', willOpen);
            button.setAttribute('aria-expanded', String(willOpen));
        });
        document.addEventListener('pointerdown', (event) => {
            const panel = getElement('cadLayerPanel');
            const button = getElement('displaySettingsBtn');
            if (!panel || !button || panel.hidden) return;
            if (panel.contains(event.target) || button.contains(event.target)) return;
            panel.hidden = true;
            button.classList.remove('active');
            button.setAttribute('aria-expanded', 'false');
        });
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
            syncOrientationFromDevice();
            window.kakao.maps.event.trigger(map, 'resize');
            updateOverlayPosition();
            applyDetailTransform();
            await renderRaster(true);
        });
        window.screen?.orientation?.addEventListener?.('change', () => {
            window.setTimeout(syncOrientationFromDevice, 0);
        });
        controlsBound = true;
    }

    function bindMapEvents() {
        window.kakao.maps.event.addListener(map, 'zoom_start', onZoomStart);
        window.kakao.maps.event.addListener(map, 'zoom_changed', onMapZoomChanged);
        window.kakao.maps.event.addListener(map, 'center_changed', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'bounds_changed', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'dragstart', () => canvas?.classList.remove('zooming'));
        window.kakao.maps.event.addListener(map, 'drag', queuePositionUpdate);
        window.kakao.maps.event.addListener(map, 'idle', onMapIdle);

        const mapSurface = getElement('kakaoMap');
        const touchOptions = { passive: true, capture: true };
        mapSurface?.addEventListener('touchstart', beginPinch, touchOptions);
        mapSurface?.addEventListener('touchmove', updatePinch, touchOptions);
        mapSurface?.addEventListener('touchend', endPinch, touchOptions);
        mapSurface?.addEventListener('touchcancel', endPinch, touchOptions);

        const stage = getElement('mapZoomStage');
        stage?.addEventListener('pointerdown', beginDetailPointerPan);
        stage?.addEventListener('pointermove', updateDetailPointerPan);
        stage?.addEventListener('pointerup', endDetailPointerPan);
        stage?.addEventListener('pointercancel', endDetailPointerPan);
        stage?.addEventListener('wheel', detailWheelZoom, { passive: false });
    }

    async function initialize() {
        const loading = getElement('mapLoading');

        try {
            const results = await Promise.all([loadKakaoMapSdk(), loadManifest()]);
            manifest = results[1];
            canvas = getElement('cadOverlay');
            syncOrientationFromDevice();
            if (!canvas) throw new Error('도면 표시 화면을 준비하지 못했습니다.');

            if (!map) {
                const [longitude, latitude] = manifest.center_wgs84;
                map = new window.kakao.maps.Map(getElement('kakaoMap'), {
                    center: new window.kakao.maps.LatLng(latitude, longitude),
                    level: 4
                });
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
