(function () {
    'use strict';

    const NAVER_MAPS_KEY_META_NAME = 'naver-maps-ncp-key-id';
    const NAVER_MIN_ZOOM = 0;
    const NAVER_MAX_ZOOM = 21;
    const MAP_SCRIPT_BASE_URL = new URL(
        '.',
        document.currentScript?.src || window.location?.href || 'http://localhost/'
    );
    const CAD_MANIFEST_URL = new URL('cad-data/hopo/manifest.json', MAP_SCRIPT_BASE_URL).href;
    const CORE_LAYER_NAMES = new Set([
        '0', 'SIMPLE', 'CABLE', '신설', '신설1', 'WALL', '전주',
        'teamA', 'teamB', 'teamC', 'teamD',
        'TEAM_A', 'TEAM_B', 'TEAM_C', 'TEAM_D'
    ]);
    const OVERLAY_PADDING_RATIO = 0.015;
    const SEARCH_RESULT_LIMIT = 100;
    const LABEL_FONT_SIZES = Object.freeze({ small: 14, medium: 17, large: 20 });
    const SEARCH_MARKER_ICON = '<span class="cad-search-pin" aria-hidden="true"><svg viewBox="0 0 30 38"><path d="M15 1.5C7.55 1.5 1.5 7.55 1.5 15c0 9.7 13.5 21.5 13.5 21.5S28.5 24.7 28.5 15C28.5 7.55 22.45 1.5 15 1.5Z"/><circle cx="15" cy="15" r="5.2"/></svg></span>';
    const CURRENT_POSITION_ICON = '<span class="current-position-marker" aria-hidden="true"></span>';

    const layerCache = new Map();
    const selectedLayers = new Set();

    let sdkPromise = null;
    let manifestPromise = null;
    let manifest = null;
    let manifestUrl = null;
    let map = null;
    let canvas = null;
    let cadOverlayView = null;
    let controlsBound = false;
    let layerListBuilt = false;
    let initialSelectionApplied = false;
    let initialBoundsApplied = false;
    let currentMapType = 'skyview';
    let positionFrame = 0;
    let idleTimer = 0;
    let rasterDirty = true;
    let renderedZoom = null;
    let renderRevision = 0;
    let renderRequestSerial = 0;
    let relayoutFrame = 0;
    let currentPositionMarker = null;
    let currentPositionAccuracy = null;
    let locationWatchId = null;
    let hasLocationFix = false;
    let mapRotationDegrees = 0;
    let transformedPanGesture = null;
    let mapInteractionState = '';
    let searchIndexPromise = null;
    let searchMarker = null;
    let searchDebounceTimer = 0;
    let searchRequestSerial = 0;
    let activeSearchResults = [];
    let selectedSearchResultIndex = -1;
    let currentLabelSize = 'small';

    function getElement(id) {
        return document.getElementById(id);
    }

    function customTransformActive() {
        return mapRotationDegrees !== 0;
    }

    function updateZoomControls() {
        const zoomInButton = getElement('zoomInBtn');
        const zoomOutButton = getElement('zoomOutBtn');
        if (zoomInButton) zoomInButton.disabled = map?.getZoom() >= NAVER_MAX_ZOOM;
        if (zoomOutButton) zoomOutButton.disabled = map?.getZoom() <= NAVER_MIN_ZOOM;
    }

    function applyMapTransform() {
        const stage = getElement('mapZoomStage');
        if (!stage) return;

        const view = getElement('mapView');
        if (view && mapRotationDegrees !== 0) {
            const width = Math.max(1, view.clientWidth);
            const height = Math.max(1, view.clientHeight);
            stage.style.width = `${height}px`;
            stage.style.height = `${width}px`;
            stage.style.left = `${(width - height) / 2}px`;
            stage.style.top = `${(height - width) / 2}px`;
            stage.style.right = 'auto';
            stage.style.bottom = 'auto';
        } else {
            stage.style.width = '';
            stage.style.height = '';
            stage.style.left = '';
            stage.style.top = '';
            stage.style.right = '';
            stage.style.bottom = '';
        }
        const active = customTransformActive();
        stage.classList.toggle('map-rotation-active', active);
        if (stage.style.setProperty) {
            stage.style.setProperty('--map-counter-rotation', `${-mapRotationDegrees}deg`);
        } else {
            stage.style['--map-counter-rotation'] = `${-mapRotationDegrees}deg`;
        }
        stage.style.transform = active
            ? `rotate(${mapRotationDegrees}deg)`
            : '';

        const nextInteractionState = String(active);
        if (map && nextInteractionState !== mapInteractionState) {
            map.setOptions('draggable', !active);
            mapInteractionState = nextInteractionState;
        }
        updateZoomControls();
    }

    function resetViewTransform() {
        applyMapTransform();
    }

    function screenDeltaToMapDelta(deltaX, deltaY) {
        return screenVectorToStageVector(deltaX, deltaY);
    }

    function screenVectorToStageVector(deltaX, deltaY) {
        const radians = mapRotationDegrees * Math.PI / 180;
        const cosine = Math.cos(radians);
        const sine = Math.sin(radians);
        const x = (cosine * deltaX) + (sine * deltaY);
        const y = (-sine * deltaX) + (cosine * deltaY);
        return {
            x: Math.abs(x) < 1e-9 ? 0 : x,
            y: Math.abs(y) < 1e-9 ? 0 : y
        };
    }

    function normalizedScreenAngle() {
        const screenAngle = Number(window.screen?.orientation?.angle);
        if (Number.isFinite(screenAngle)) return ((screenAngle % 360) + 360) % 360;

        const legacyAngle = Number(window.orientation);
        if (Number.isFinite(legacyAngle)) return ((legacyAngle % 360) + 360) % 360;
        return null;
    }

    function detectedMapRotation(viewportLandscape) {
        if (!viewportLandscape) return 0;
        const angle = normalizedScreenAngle();
        if (angle === 270) return 90;
        if (angle === 90) return -90;
        return -90;
    }

    function updateRenderedLabelRotations() {
        const labels = canvas?.querySelectorAll?.('.cad-map-label') || [];
        for (const label of labels) {
            const x = label.getAttribute('x');
            const y = label.getAttribute('y');
            label.setAttribute('transform', `rotate(${-mapRotationDegrees} ${x} ${y})`);
        }
    }

    function syncOrientationFromDevice() {
        const view = getElement('mapView');
        const viewportLandscape = window.innerWidth > window.innerHeight;
        view?.classList.toggle('landscape-mode', viewportLandscape);
        const nextRotation = detectedMapRotation(viewportLandscape);
        const rotationChanged = nextRotation !== mapRotationDegrees;
        mapRotationDegrees = nextRotation;
        if (rotationChanged) {
            applyMapTransform();
            updateRenderedLabelRotations();
        }
    }

    function panTransformedMap(deltaX, deltaY) {
        if (!map || (!deltaX && !deltaY)) return;
        const projection = map.getProjection();
        const centerPoint = projection.fromCoordToOffset(map.getCenter());
        const delta = screenDeltaToMapDelta(deltaX, deltaY);
        const targetPoint = new window.naver.maps.Point(
            centerPoint.x - delta.x,
            centerPoint.y - delta.y
        );
        map.setCenter(projection.fromOffsetToCoord(targetPoint));
    }

    function zoomIn() {
        if (!map) return;
        map.setZoom(Math.min(NAVER_MAX_ZOOM, map.getZoom() + 1), true);
    }

    function zoomOut() {
        if (!map) return;
        map.setZoom(Math.max(NAVER_MIN_ZOOM, map.getZoom() - 1), true);
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

    function naverMapsKeyId() {
        return String(document.querySelector(`meta[name="${NAVER_MAPS_KEY_META_NAME}"]`)?.content || '').trim();
    }

    function loadNaverMapSdk() {
        if (window.naver?.maps?.Map) return Promise.resolve();

        if (sdkPromise) return sdkPromise;

        sdkPromise = new Promise((resolve, reject) => {
            const keyId = naverMapsKeyId();
            if (!keyId || keyId === 'YOUR_NCP_KEY_ID') {
                reject(new Error('네이버 지도 ncpKeyId를 index.html에 설정해 주세요.'));
                return;
            }

            const previousAuthFailure = window.navermap_authFailure;
            window.navermap_authFailure = () => {
                previousAuthFailure?.();
                reject(new Error('네이버 지도 인증에 실패했습니다. ncpKeyId와 Web 서비스 URL을 확인해 주세요.'));
            };
            const script = document.createElement('script');
            script.src = `https://oapi.map.naver.com/openapi/v3/maps.js?ncpKeyId=${encodeURIComponent(keyId)}`;
            script.async = true;
            script.addEventListener('load', () => {
                if (!window.naver?.maps?.Map) {
                    reject(new Error('네이버 지도 SDK를 시작하지 못했습니다.'));
                    return;
                }
                resolve();
            });
            script.addEventListener('error', () => {
                reject(new Error('네이버 지도 SDK를 내려받지 못했습니다.'));
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

    function normalizedSearchText(value) {
        return String(value || '').normalize('NFKC').trim().toLocaleLowerCase('ko-KR');
    }

    function setSearchStatus(text, isError = false) {
        const status = getElement('mapSearchStatus');
        if (!status) return;
        status.textContent = text;
        status.classList.toggle('error', isError);
    }

    function clearSearchResults() {
        activeSearchResults = [];
        selectedSearchResultIndex = -1;
        getElement('mapSearchList')?.replaceChildren();
        setSearchNavigationVisible(false);
    }

    function buildSearchIndex() {
        if (searchIndexPromise) return searchIndexPromise;

        searchIndexPromise = Promise.all(manifest.layers.map(async (layerInfo) => {
            try {
                const layer = await loadLayer(layerInfo);
                return (Array.isArray(layer.labels) ? layer.labels : []).flatMap((label) => {
                    const text = String(label.text || '').trim();
                    const position = Array.isArray(label.position) ? label.position : null;
                    const normalized = normalizedSearchText(text);
                    if (!normalized || !position || position.length < 2) return [];
                    return [{
                        text,
                        normalized,
                        layerName: layerInfo.name,
                        position: [Number(position[0]), Number(position[1])]
                    }];
                });
            } catch (error) {
                console.warn(`Search index skipped layer ${layerInfo.name}:`, error);
                return [];
            }
        })).then((groups) => groups.flat());

        return searchIndexPromise;
    }

    function matchingSearchResults(index, query) {
        const normalizedQuery = normalizedSearchText(query);
        if (!normalizedQuery) return { total: 0, results: [] };

        const matches = [];
        for (const entry of index) {
            const matchAt = entry.normalized.indexOf(normalizedQuery);
            if (matchAt < 0) continue;
            matches.push({
                ...entry,
                score: entry.normalized === normalizedQuery ? 0 : matchAt === 0 ? 1 : 2
            });
        }
        matches.sort((first, second) => first.score - second.score
            || first.text.localeCompare(second.text, 'ko-KR', { numeric: true })
            || first.layerName.localeCompare(second.layerName, 'ko-KR'));
        return {
            total: matches.length,
            results: matches.slice(0, SEARCH_RESULT_LIMIT)
        };
    }

    function updateSearchNavigation() {
        const text = getElement('mapSearchNavigationText');
        const layer = getElement('mapSearchNavigationLayer');
        const status = getElement('mapSearchNavigationStatus');
        const previous = getElement('mapSearchPrevBtn');
        const next = getElement('mapSearchNextBtn');
        const hasSelection = selectedSearchResultIndex >= 0 && activeSearchResults.length > 0;
        setSearchNavigationVisible(hasSelection);
        if (!hasSelection) return;
        const selectedResult = activeSearchResults[selectedSearchResultIndex];
        if (text) {
            text.textContent = selectedResult.text;
            text.title = selectedResult.text;
        }
        if (layer) {
            layer.textContent = selectedResult.layerName;
            layer.title = selectedResult.layerName;
        }
        if (status) status.textContent = `${selectedSearchResultIndex + 1}/${activeSearchResults.length}`;
        if (previous) previous.disabled = selectedSearchResultIndex === 0;
        if (next) next.disabled = selectedSearchResultIndex === activeSearchResults.length - 1;
    }

    function setSearchNavigationVisible(visible) {
        const navigation = getElement('mapSearchNavigation');
        const controls = getElement('mapTopControls');
        if (navigation) navigation.hidden = !visible;
        controls?.classList.toggle('result-selected', visible);
    }

    function reopenSearchResults() {
        if (!activeSearchResults.length) return;
        const results = getElement('mapSearchResults');
        if (results) results.hidden = false;
        setSearchNavigationVisible(false);
    }

    function selectSearchResult(result, resultIndex = activeSearchResults.indexOf(result)) {
        if (!map || !result) return;
        resetViewTransform();
        selectedSearchResultIndex = Math.max(0, resultIndex);
        const [longitude, latitude] = result.position;
        const position = new window.naver.maps.LatLng(latitude, longitude);
        searchMarker?.setMap?.(null);
        searchMarker = new window.naver.maps.Marker({
            map,
            position,
            title: result.text,
            icon: {
                content: SEARCH_MARKER_ICON,
                anchor: new window.naver.maps.Point(15, 38)
            }
        });
        if (map.getZoom() < 20) map.setZoom(20, true);
        map.panTo(position);
        getElement('mapSearchResults').hidden = true;
        updateSearchNavigation();
    }

    function moveSearchSelection(offset) {
        if (!activeSearchResults.length || selectedSearchResultIndex < 0) return;
        const nextIndex = Math.max(0, Math.min(
            activeSearchResults.length - 1,
            selectedSearchResultIndex + offset
        ));
        if (nextIndex === selectedSearchResultIndex) return;
        selectSearchResult(activeSearchResults[nextIndex], nextIndex);
    }

    function renderSearchResults(resultSet) {
        const resultPanel = getElement('mapSearchResults');
        const resultList = getElement('mapSearchList');
        if (!resultPanel || !resultList) return;

        activeSearchResults = resultSet.results;
        const fragment = document.createDocumentFragment();
        for (let index = 0; index < activeSearchResults.length; index += 1) {
            const result = activeSearchResults[index];
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'map-search-result';

            const title = document.createElement('span');
            title.className = 'map-search-result-title';
            title.textContent = result.text;

            const meta = document.createElement('span');
            meta.className = 'map-search-result-meta';
            meta.textContent = `${result.layerName} · ${index + 1}번째 결과`;

            button.append(title, meta);
            button.addEventListener('click', () => selectSearchResult(result, index));
            fragment.appendChild(button);
        }
        resultList.replaceChildren(fragment);
        resultPanel.hidden = false;

        if (resultSet.total === 0) {
            setSearchStatus('검색 결과가 없습니다.');
        } else if (resultSet.total > resultSet.results.length) {
            setSearchStatus(`전체 ${resultSet.total.toLocaleString()}건 중 ${resultSet.results.length}건을 표시합니다.`);
        } else {
            setSearchStatus(`${resultSet.total.toLocaleString()}건을 찾았습니다.`);
        }
    }

    async function performTextSearch() {
        const input = getElement('mapSearchInput');
        const query = String(input?.value || '').trim();
        const requestSerial = ++searchRequestSerial;
        clearSearchResults();

        if (!query) {
            getElement('mapSearchResults').hidden = false;
            setSearchStatus('검색어를 입력하세요.');
            return;
        }

        getElement('mapSearchResults').hidden = false;
        setSearchStatus('도면 문자를 검색하고 있습니다.');
        try {
            const index = await buildSearchIndex();
            if (requestSerial !== searchRequestSerial) return;
            renderSearchResults(matchingSearchResults(index, query));
        } catch (error) {
            console.error('CAD text search failed:', error);
            if (requestSerial === searchRequestSerial) {
                setSearchStatus('검색 데이터를 불러오지 못했습니다.', true);
            }
        }
    }

    function scheduleTextSearch() {
        activeSearchResults = [];
        selectedSearchResultIndex = -1;
        setSearchNavigationVisible(false);
        window.clearTimeout(searchDebounceTimer);
        searchDebounceTimer = window.setTimeout(performTextSearch, 140);
    }

    function closeTextSearch() {
        window.clearTimeout(searchDebounceTimer);
        searchRequestSerial += 1;
        clearSearchResults();
        const controls = getElement('mapTopControls');
        const primary = getElement('mapPrimaryControls');
        const panel = getElement('mapSearchPanel');
        const results = getElement('mapSearchResults');
        const input = getElement('mapSearchInput');
        controls?.classList.remove('search-active');
        if (primary) primary.hidden = false;
        if (panel) panel.hidden = true;
        if (results) results.hidden = true;
        if (input) input.value = '';
    }

    function openTextSearch() {
        const controls = getElement('mapTopControls');
        const primary = getElement('mapPrimaryControls');
        const panel = getElement('mapSearchPanel');
        const results = getElement('mapSearchResults');
        const input = getElement('mapSearchInput');
        const layerPanel = getElement('cadLayerPanel');
        const settingsButton = getElement('displaySettingsBtn');

        if (layerPanel) layerPanel.hidden = true;
        settingsButton?.classList.remove('active');
        settingsButton?.setAttribute('aria-expanded', 'false');
        controls?.classList.add('search-active');
        if (primary) primary.hidden = true;
        if (panel) panel.hidden = false;
        if (results) results.hidden = false;
        clearSearchResults();
        setSearchStatus('검색 데이터를 준비하고 있습니다.');
        input?.focus();

        buildSearchIndex().then(() => {
            if (!String(input?.value || '').trim()) setSearchStatus('검색어를 입력하세요.');
        }).catch((error) => {
            console.error('CAD search index initialization failed:', error);
            setSearchStatus('검색 데이터를 불러오지 못했습니다.', true);
        });
    }

    function setMapType(type) {
        if (!map) return;

        resetViewTransform();
        currentMapType = type === 'roadmap' ? 'roadmap' : 'skyview';
        map.setMapTypeId(currentMapType === 'skyview'
            ? window.naver.maps.MapTypeId.SATELLITE
            : window.naver.maps.MapTypeId.NORMAL);

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
        const bounds = new window.naver.maps.LatLngBounds(
            new window.naver.maps.LatLng(south, west),
            new window.naver.maps.LatLng(north, east)
        );
        map.fitBounds(bounds, 42);
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

    function updateLocationTrackingButton(active) {
        const button = getElement('currentLocationBtn');
        if (!button) return;
        button.classList.toggle('active', active);
        button.setAttribute('aria-pressed', String(active));
        button.setAttribute('aria-label', active
            ? '현재위치 실시간 추적 중지'
            : '현재위치 실시간 추적 시작');
    }

    function stopLocationTracking(showStatus = true) {
        if (locationWatchId !== null && navigator.geolocation?.clearWatch) {
            navigator.geolocation.clearWatch(locationWatchId);
        }
        locationWatchId = null;
        hasLocationFix = false;
        updateLocationTrackingButton(false);
        if (showStatus) setLocationStatus('현재위치 실시간 추적을 중지했습니다.', 'ready');
    }

    function updateTrackedPosition(position) {
        const latitude = position.coords.latitude;
        const longitude = position.coords.longitude;
        const accuracy = Math.max(1, position.coords.accuracy || 1);
        const latLng = new window.naver.maps.LatLng(latitude, longitude);

        if (!currentPositionAccuracy) {
            currentPositionAccuracy = new window.naver.maps.Circle({
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
            currentPositionAccuracy.setCenter(latLng);
            currentPositionAccuracy.setRadius(accuracy);
        }

        if (!currentPositionMarker) {
            currentPositionMarker = new window.naver.maps.Marker({
                map,
                position: latLng,
                title: '현재위치',
                icon: {
                    content: CURRENT_POSITION_ICON,
                    anchor: new window.naver.maps.Point(12, 12)
                }
            });
        } else {
            currentPositionMarker.setMap(map);
            currentPositionMarker.setPosition(latLng);
        }

        if (!hasLocationFix) {
            if (map.getZoom() < 18) map.setZoom(18, true);
            map.panTo(latLng);
        }
        hasLocationFix = true;
        setLocationStatus(`현재위치 추적 중 · 정확도 약 ${Math.round(accuracy)}m`, 'ready');
    }

    function toggleLocationTracking() {
        const button = getElement('currentLocationBtn');
        if (!map || !navigator.geolocation?.watchPosition || !button) {
            setLocationStatus('이 브라우저에서는 현재 위치 기능을 사용할 수 없습니다.', 'error');
            return;
        }

        if (locationWatchId !== null) {
            stopLocationTracking();
            return;
        }

        resetViewTransform();
        hasLocationFix = false;
        updateLocationTrackingButton(true);
        setLocationStatus('현재위치 권한을 확인하고 실시간 추적을 시작합니다.', 'loading');
        locationWatchId = navigator.geolocation.watchPosition(updateTrackedPosition, (error) => {
            stopLocationTracking(false);
            setLocationStatus(geolocationErrorMessage(error), 'error');
        }, {
            enableHighAccuracy: true,
            timeout: 15000,
            maximumAge: 3000
        });
    }

    function setLabelSize(size) {
        const normalizedSize = Object.hasOwn(LABEL_FONT_SIZES, size) ? size : 'small';
        currentLabelSize = normalizedSize;
        const fontSize = `${LABEL_FONT_SIZES[normalizedSize]}px`;
        if (canvas?.style?.setProperty) {
            canvas.style.setProperty('--cad-label-font-size', fontSize);
        } else if (canvas?.style) {
            canvas.style['--cad-label-font-size'] = fontSize;
        }

        for (const button of document.querySelectorAll('[data-cad-label-size]')) {
            const active = button.dataset.cadLabelSize === normalizedSize;
            button.classList.toggle('active', active);
            button.setAttribute('aria-pressed', String(active));
        }
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
        ].map(([longitude, latitude]) => projection.fromCoordToOffset(
            new window.naver.maps.LatLng(latitude, longitude)
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

    function attachCadOverlayToMap() {
        if (!map || !canvas || cadOverlayView) return;

        class CadSvgOverlay extends window.naver.maps.OverlayView {
            onAdd() {
                this.getPanes().overlayLayer.appendChild(canvas);
            }

            draw() {
                updateOverlayPosition();
            }

            onRemove() {
                canvas.remove();
            }
        }

        cadOverlayView = new CadSvgOverlay();
        cadOverlayView.setMap(map);
    }

    function queuePositionUpdate() {
        if (positionFrame) return;
        positionFrame = window.requestAnimationFrame(() => {
            positionFrame = 0;
            updateOverlayPosition();
        });
    }

    function beginPinch(event) {
        if (event.touches.length > 1) {
            transformedPanGesture = null;
            getElement('mapZoomStage')?.classList.remove('dragging');
            return;
        }
        if (customTransformActive() && event.touches.length === 1) {
            const touch = event.touches[0];
            transformedPanGesture = {
                x: touch.clientX,
                y: touch.clientY
            };
            getElement('mapZoomStage')?.classList.add('dragging');
        }
    }

    function updatePinch(event) {
        if (transformedPanGesture && event.touches.length === 1 && customTransformActive()) {
            const touch = event.touches[0];
            panTransformedMap(
                touch.clientX - transformedPanGesture.x,
                touch.clientY - transformedPanGesture.y
            );
            transformedPanGesture.x = touch.clientX;
            transformedPanGesture.y = touch.clientY;
        }
    }

    function endPinch(event) {
        if (transformedPanGesture && event.touches.length === 0) {
            transformedPanGesture = null;
            getElement('mapZoomStage')?.classList.remove('dragging');
        }
    }

    function beginTransformedPointerPan(event) {
        if (!customTransformActive() || event.pointerType === 'touch' || event.button !== 0) return;
        const stage = getElement('mapZoomStage');
        transformedPanGesture = {
            pointerId: event.pointerId,
            x: event.clientX,
            y: event.clientY
        };
        stage?.setPointerCapture?.(event.pointerId);
        stage?.classList.add('dragging');
        event.preventDefault();
    }

    function updateTransformedPointerPan(event) {
        if (!transformedPanGesture || transformedPanGesture.pointerId !== event.pointerId) return;
        panTransformedMap(
            event.clientX - transformedPanGesture.x,
            event.clientY - transformedPanGesture.y
        );
        transformedPanGesture.x = event.clientX;
        transformedPanGesture.y = event.clientY;
        event.preventDefault();
    }

    function endTransformedPointerPan(event) {
        if (!transformedPanGesture || transformedPanGesture.pointerId !== event.pointerId) return;
        transformedPanGesture = null;
        getElement('mapZoomStage')?.classList.remove('dragging');
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

        const requestSerial = ++renderRequestSerial;
        const revision = renderRevision;
        const zoom = map.getZoom();
        const box = overlayScreenBox();
        const width = Math.max(1, box.width);
        const height = Math.max(1, box.height);

        const projection = map.getProjection();
        const toPixel = ([longitude, latitude]) => {
            const point = projection.fromCoordToOffset(
                new window.naver.maps.LatLng(latitude, longitude)
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
                        'font-family': 'Malgun Gothic, sans-serif',
                        transform: `rotate(${-mapRotationDegrees} ${x} ${y})`
                    });
                    text.textContent = label.text;
                    fragment.appendChild(text);
                }
            }
        }

        if (requestSerial !== renderRequestSerial) return;
        if (revision !== renderRevision || zoom !== map.getZoom()) {
            rasterDirty = true;
            return;
        }

        canvas.style.left = `${box.left}px`;
        canvas.style.top = `${box.top}px`;
        canvas.style.width = `${width}px`;
        canvas.style.height = `${height}px`;
        canvas.setAttribute('viewBox', `0 0 ${width} ${height}`);
        canvas.replaceChildren(fragment);
        rasterDirty = false;
        renderedZoom = zoom;
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
        renderRevision += 1;
        rasterDirty = true;
        queuePositionUpdate();
    }

    function onMapIdle() {
        queuePositionUpdate();
        window.clearTimeout(idleTimer);
        idleTimer = window.setTimeout(async () => {
            if (rasterDirty || renderedZoom !== map.getZoom()) {
                await renderRaster(true);
            } else {
                updateOverlayPosition();
            }
            updateZoomControls();
        }, 330);
    }

    function onMapZoomChanged() {
        renderRevision += 1;
        rasterDirty = true;
        queuePositionUpdate();
        updateZoomControls();
    }

    function relayoutMap() {
        if (!map) return false;
        const view = getElement('mapView');
        const surface = getElement('naverMap');
        const viewWidth = Number(view?.clientWidth || 0);
        const viewHeight = Number(view?.clientHeight || 0);
        if (viewWidth <= 0 || viewHeight <= 0 || !surface) return false;
        const width = customTransformActive() ? viewHeight : viewWidth;
        const height = customTransformActive() ? viewWidth : viewHeight;
        surface.style.width = `${width}px`;
        surface.style.height = `${height}px`;
        map.setSize(new window.naver.maps.Size(width, height));
        return true;
    }

    function scheduleResponsiveRelayout() {
        syncOrientationFromDevice();
        applyMapTransform();
        if (!map) return;
        if (relayoutFrame) window.cancelAnimationFrame?.(relayoutFrame);
        relayoutFrame = window.requestAnimationFrame(() => {
            relayoutFrame = window.requestAnimationFrame(async () => {
                relayoutFrame = 0;
                syncOrientationFromDevice();
                applyMapTransform();
                const center = map.getCenter();
                if (!relayoutMap()) return;
                map.setCenter(center);
                renderRevision += 1;
                rasterDirty = true;
                updateOverlayPosition();
                await renderRaster(true);
            });
        });
    }

    function bindControls() {
        if (controlsBound) return;
        getElement('mapTypeToggleBtn')?.addEventListener('click', toggleMapType);
        getElement('currentLocationBtn')?.addEventListener('click', toggleLocationTracking);
        getElement('zoomInBtn')?.addEventListener('click', zoomIn);
        getElement('zoomOutBtn')?.addEventListener('click', zoomOut);
        getElement('mapSearchBtn')?.addEventListener('click', openTextSearch);
        getElement('mapSearchBackBtn')?.addEventListener('click', closeTextSearch);
        getElement('mapSearchInput')?.addEventListener('input', scheduleTextSearch);
        getElement('mapSearchInput')?.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                closeTextSearch();
            } else if (event.key === 'Enter' && activeSearchResults.length === 1) {
                event.preventDefault();
                selectSearchResult(activeSearchResults[0], 0);
            }
        });
        getElement('mapSearchPrevBtn')?.addEventListener('click', () => moveSearchSelection(-1));
        getElement('mapSearchNextBtn')?.addEventListener('click', () => moveSearchSelection(1));
        getElement('mapSearchNavigationText')?.addEventListener('click', reopenSearchResults);
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
        for (const button of document.querySelectorAll('[data-cad-label-size]')) {
            button.addEventListener('click', () => setLabelSize(button.dataset.cadLabelSize));
        }
        getElement('cadOpacity')?.addEventListener('input', (event) => {
            if (canvas) canvas.style.opacity = String(Number(event.target.value) / 100);
        });
        window.addEventListener('resize', scheduleResponsiveRelayout);
        window.screen?.orientation?.addEventListener?.('change', () => {
            window.setTimeout(scheduleResponsiveRelayout, 0);
        });
        if (window.ResizeObserver) {
            new window.ResizeObserver(scheduleResponsiveRelayout).observe(getElement('mapView'));
        }
        controlsBound = true;
    }

    function bindMapEvents() {
        window.naver.maps.Event.addListener(map, 'zooming', onZoomStart);
        window.naver.maps.Event.addListener(map, 'zoom_changed', onMapZoomChanged);
        window.naver.maps.Event.addListener(map, 'center_changed', queuePositionUpdate);
        window.naver.maps.Event.addListener(map, 'bounds_changed', queuePositionUpdate);
        window.naver.maps.Event.addListener(map, 'dragstart', () => canvas?.classList.remove('zooming'));
        window.naver.maps.Event.addListener(map, 'drag', queuePositionUpdate);
        window.naver.maps.Event.addListener(map, 'idle', onMapIdle);

        const mapSurface = getElement('naverMap');
        const touchOptions = { passive: true, capture: true };
        mapSurface?.addEventListener('touchstart', beginPinch, touchOptions);
        mapSurface?.addEventListener('touchmove', updatePinch, touchOptions);
        mapSurface?.addEventListener('touchend', endPinch, touchOptions);
        mapSurface?.addEventListener('touchcancel', endPinch, touchOptions);

        const stage = getElement('mapZoomStage');
        stage?.addEventListener('pointerdown', beginTransformedPointerPan);
        stage?.addEventListener('pointermove', updateTransformedPointerPan);
        stage?.addEventListener('pointerup', endTransformedPointerPan);
        stage?.addEventListener('pointercancel', endTransformedPointerPan);
    }

    async function initialize() {
        const loading = getElement('mapLoading');

        try {
            const results = await Promise.all([loadNaverMapSdk(), loadManifest()]);
            manifest = results[1];
            canvas = getElement('cadOverlay');
            syncOrientationFromDevice();
            if (!canvas) throw new Error('도면 표시 화면을 준비하지 못했습니다.');
            setLabelSize(currentLabelSize);

            if (!map) {
                const [longitude, latitude] = manifest.center_wgs84;
                map = new window.naver.maps.Map(getElement('naverMap'), {
                    center: new window.naver.maps.LatLng(latitude, longitude),
                    zoom: 17,
                    minZoom: NAVER_MIN_ZOOM,
                    maxZoom: NAVER_MAX_ZOOM,
                    mapTypeId: window.naver.maps.MapTypeId.SATELLITE,
                    mapTypeControl: false,
                    scaleControl: true,
                    zoomControl: false
                });
                bindMapEvents();
            }

            attachCadOverlayToMap();

            buildLayerList();
            bindControls();
            relayoutMap();
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
            console.error('NAVER CAD map initialization failed:', error);
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
            scheduleResponsiveRelayout();
        }
    };
}());
