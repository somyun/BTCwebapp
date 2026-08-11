(() => {
    'use strict';

    const STORAGE_PREFIX = '';
    const BASE = 'https://asia-northeast3-btcwebapp-551bd.cloudfunctions.net';
    const ENDPOINTS = Object.freeze({
        register: `${BASE}/registerNotificationDevice`,
        active: `${BASE}/setNotificationDeviceActive`,
        status: `${BASE}/getNotificationDeviceStatus`,
        acknowledge: `${BASE}/acknowledgeNotification`,
        selfTest: `${BASE}/sendNotificationSelfTest`
    });
    const FIREBASE_CONFIG = Object.freeze({
        apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
        authDomain: 'btcwebapp-551bd.firebaseapp.com',
        projectId: 'btcwebapp-551bd',
        storageBucket: 'btcwebapp-551bd.firebasestorage.app',
        messagingSenderId: '237989935469',
        appId: '1:237989935469:web:07fc002a5c2ab2f5858264'
    });
    const VAPID_KEY = 'BCIeuJhwW92Usr-QS3BFOUWnP2pZ4rqulcmZBlxXdv8Ayms7zllnqLy-jNj9NtmOrkJfE9ywMkkj0IegbKxDDmE';
    const REQUEST_TIMEOUT_MS = 10000;
    const KEYWORD_SAVE_DELAY_MS = 700;

    let identity = null;
    let registration = null;
    let messaging = null;
    let lastMessagingToken = '';
    let messagingTokenPromise = null;
    let iosGuideRequested = false;
    let keywordSaveTimer = null;
    let keywordSyncInFlight = null;
    let keywordSyncQueued = false;
    let lastSyncedKeywords = '';
    let deviceSyncTail = Promise.resolve();
    let toggleOperationId = 0;

    function storageKey(name) {
        return `${STORAGE_PREFIX}${name}`;
    }

    function getFromStorage(name, fallback = null) {
        try {
            const value = localStorage.getItem(storageKey(name));
            return value === null ? fallback : JSON.parse(value);
        } catch (_) {
            return fallback;
        }
    }

    function saveToStorage(name, value) {
        localStorage.setItem(storageKey(name), JSON.stringify(value));
    }

    function removeFromStorage(name) {
        localStorage.removeItem(storageKey(name));
    }

    function delay(milliseconds) {
        return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
    }

    async function fetchWithTimeout(url, options) {
        const controller = new AbortController();
        const timer = window.setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
        try {
            return await fetch(url, { ...options, signal: controller.signal });
        } catch (error) {
            if (error.name === 'AbortError') throw new Error('NETWORK_TIMEOUT');
            throw error;
        } finally {
            window.clearTimeout(timer);
        }
    }

    function isTransientNetworkError(error) {
        return /Failed to fetch|NetworkError|NETWORK_TIMEOUT|HTTP_5\d\d/i.test(error?.message || '');
    }

    async function withNetworkRetry(operation) {
        try {
            return await operation();
        } catch (error) {
            if (!isTransientNetworkError(error)) throw error;
            await delay(350);
            return operation();
        }
    }

    function friendlyError(error) {
        if (isTransientNetworkError(error)) return '네트워크 연결이 불안정합니다.';
        return error?.message || '알 수 없는 오류가 발생했습니다.';
    }

    async function post(url, body) {
        const response = await fetchWithTimeout(url, {
            method: 'POST',
            cache: 'no-store',
            headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok || payload.ok !== true) throw new Error(payload.error || `HTTP_${response.status}`);
        return payload.result;
    }

    function queueDeviceSync(operation) {
        const queued = deviceSyncTail.then(operation, operation);
        deviceSyncTail = queued.catch(() => {});
        return queued;
    }

    function setText(id, value) {
        const element = document.getElementById(id);
        if (element) element.textContent = value ?? '-';
    }

    function setFeedback(id, message, type = '') {
        const element = document.getElementById(id);
        if (!element) return;
        element.textContent = message;
        element.className = `setting-feedback${type ? ` ${type}` : ''}`;
        element.hidden = !message;
    }

    function timestampMillis(value) {
        if (!value) return 0;
        if (typeof value === 'string') return Date.parse(value) || 0;
        if (typeof value === 'number') return value;
        if (typeof value === 'object') {
            const seconds = value._seconds ?? value.seconds;
            if (Number.isFinite(seconds)) return Number(seconds) * 1000;
        }
        return 0;
    }

    function formatTime(value) {
        const milliseconds = timestampMillis(value);
        return milliseconds ? new Date(milliseconds).toLocaleString('ko-KR') : '기록 없음';
    }

    function permissionValue() {
        if (!('Notification' in window)) return '지원하지 않음';
        return Notification.permission === 'granted'
            ? '허용됨'
            : Notification.permission === 'denied'
                ? '차단됨'
                : '아직 선택하지 않음';
    }

    function isIos() {
        return /iPad|iPhone|iPod/.test(navigator.userAgent) ||
            (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    }

    function isAndroid() {
        return /Android/i.test(navigator.userAgent);
    }

    function isSamsungBrowser() {
        return /SamsungBrowser/i.test(navigator.userAgent);
    }

    function isStandalone() {
        return navigator.standalone === true || window.matchMedia('(display-mode: standalone)').matches;
    }

    function updateGuidanceVisibility() {
        const active = getFromStorage('isNotificationActive') === true;
        const permissionGuide = document.getElementById('permissionGuide');
        const permissionGuideText = document.getElementById('permissionGuideText');
        const permissionAction = document.getElementById('permissionActionBtn');
        const unsupported = !('Notification' in window);
        const denied = !unsupported && Notification.permission === 'denied';
        const undecided = !unsupported && Notification.permission === 'default';
        const activeWithoutPermission = active && !unsupported && Notification.permission !== 'granted';
        const iosNeedsInstall = isIos() && !isStandalone();
        permissionGuide.hidden = !((unsupported && !iosNeedsInstall) || denied || activeWithoutPermission);
        permissionAction.hidden = true;
        if (unsupported) {
            permissionGuideText.textContent = '이 브라우저는 웹앱 알림을 지원하지 않습니다.';
        } else if (denied) {
            permissionGuideText.textContent = '알림이 브라우저에서 차단되어 자동 권한창을 다시 열 수 없습니다.';
            permissionAction.textContent = '설정 방법 보기';
            permissionAction.hidden = false;
        } else if (undecided && active) {
            permissionGuideText.textContent = '아래 버튼을 누르면 브라우저의 알림 권한 선택창이 열립니다.';
            permissionAction.textContent = '알림 권한 허용';
            permissionAction.hidden = false;
        } else {
            permissionGuideText.textContent = '기기 설정에서 이 웹앱의 알림을 허용해 주세요.';
        }

        const iosGuide = document.getElementById('iosSettingsGuide');
        iosGuide.hidden = !((iosGuideRequested || active) && iosNeedsInstall);
        if (!iosGuide.hidden) iosGuide.open = true;
    }

    function permissionSettingsInstructions() {
        if (isSamsungBrowser()) {
            return '삼성 인터넷의 주소창 왼쪽 사이트 정보 아이콘을 누른 뒤 권한 → 알림 → 허용으로 변경하세요. 항목이 보이지 않으면 삼성 인터넷 메뉴(≡) → 설정 → 사이트 및 다운로드 → 알림에서 이 사이트를 허용하세요.';
        }
        if (isAndroid()) {
            return '브라우저 주소창 왼쪽의 사이트 정보 아이콘을 누른 뒤 권한 → 알림 → 허용으로 변경하세요. 변경 후 이 페이지로 돌아오면 자동으로 다시 연결합니다.';
        }
        if (isIos()) {
            return 'iPhone의 설정 → 알림에서 이 웹앱을 선택해 알림 허용을 켜세요. Safari 탭으로 사용 중이라면 먼저 공유 → 홈 화면에 추가한 뒤 설치된 웹앱에서 알림을 켜야 합니다.';
        }
        return '주소창 왼쪽의 사이트 정보 아이콘을 누른 뒤 사이트 설정 → 알림 → 허용으로 변경하세요. 변경 후 이 페이지로 돌아오면 자동으로 다시 연결합니다.';
    }

    function showPermissionSettingsInstructions() {
        setText('settingsInfoTitle', '알림 권한 설정 방법');
        setText('settingsInfoText', permissionSettingsInstructions());
        const overlay = document.getElementById('settingsInfoOverlay');
        overlay.hidden = false;
        overlay.querySelector('.info-modal').focus();
    }

    async function registerAfterPermissionGranted() {
        if (getFromStorage('isNotificationActive') !== true || Notification.permission !== 'granted') return;
        const keywords = String(getFromStorage('userKeywords', '') || '').trim();
        markPendingSync(true, keywords);
        await registerActiveDevice(keywords, { requireStillActive: true });
        clearPendingSync(true, keywords);
        updateGuidanceVisibility();
        await refresh();
    }

    async function handlePermissionAction() {
        const button = document.getElementById('permissionActionBtn');
        if (!('Notification' in window)) return;
        if (Notification.permission === 'denied') {
            showPermissionSettingsInstructions();
            return;
        }
        button.disabled = true;
        try {
            const permission = await Notification.requestPermission();
            updateGuidanceVisibility();
            if (permission === 'granted') {
                await registerAfterPermissionGranted();
            } else if (permission === 'denied') {
                showPermissionSettingsInstructions();
            }
        } catch (error) {
            setFeedback('toggleStatus', `알림 권한을 확인하지 못했습니다. ${friendlyError(error)}`, 'error');
        } finally {
            button.disabled = false;
        }
    }

    function handlePermissionReturn() {
        if (document.visibilityState !== 'visible') return;
        updateGuidanceVisibility();
        if ('Notification' in window && Notification.permission === 'granted') {
            void registerAfterPermissionGranted().catch((error) => {
                setFeedback('toggleStatus', `알림 연결을 완료하지 못했습니다. ${friendlyError(error)}`, 'error');
            });
        }
    }

    function setDependentState(active) {
        const fieldset = document.getElementById('notificationDependentSettings');
        fieldset.disabled = !active;
        fieldset.hidden = !active;
        fieldset.setAttribute('aria-disabled', String(!active));
    }

    function renderToggleState(active, message = '', type = '') {
        const toggle = document.getElementById('notificationToggleSettings');
        toggle.checked = active;
        toggle.disabled = false;
        setDependentState(active);
        if (message) {
            setFeedback('toggleStatus', message, type || (active ? 'success' : 'muted'));
        } else {
            setFeedback('toggleStatus', '');
        }
        updateGuidanceVisibility();
    }

    async function acknowledge(event, phase) {
        try {
            if (!identity) return false;
            await post(ENDPOINTS.acknowledge, {
                deviceId: identity.deviceId,
                deviceSecret: identity.deviceSecret,
                eventId: event.eventId,
                type: event.type,
                phase
            });
            return true;
        } catch (_) {
            return false;
        }
    }

    async function processForeground(payload) {
        const event = window.BWANotificationStore.normalizeEvent(payload, { source: 'notification-settings-page' });
        await window.BWANotificationStore.saveEvent(event);
        await acknowledge(event, 'received');
        if (event.type === 'heartbeat') {
            await window.BWANotificationStore.setMeta('lastHeartbeat', {
                eventId: event.eventId,
                receivedAt: event.receivedAt
            });
        } else if (registration && Notification.permission === 'granted') {
            await registration.showNotification(event.title, {
                body: event.body,
                icon: payload?.data?.icon || './icon-192.png',
                tag: event.eventId,
                data: { eventId: event.eventId, type: event.type, url: event.url }
            });
            await window.BWANotificationStore.patchEvent(event.eventId, { shownAt: new Date().toISOString() });
            await acknowledge(event, 'shown');
        }
        await refresh();
    }

    async function initializeMessaging() {
        if (messaging) return messaging;
        if (!('serviceWorker' in navigator) || !window.firebase?.messaging) return null;
        registration = await navigator.serviceWorker.register('./firebase-messaging-sw.js', { scope: './' });
        if (!window.firebase.apps.length) window.firebase.initializeApp(FIREBASE_CONFIG);
        messaging = window.firebase.messaging();
        messaging.onMessage((payload) => void processForeground(payload));
        return messaging;
    }

    async function getMessagingToken({ requestPermission = false } = {}) {
        if (lastMessagingToken) return lastMessagingToken;
        if (messagingTokenPromise) return messagingTokenPromise;
        messagingTokenPromise = (async () => {
            if (!('Notification' in window)) throw new Error('이 브라우저는 알림을 지원하지 않습니다.');
            let permission = Notification.permission;
            if (requestPermission && permission === 'default') permission = await Notification.requestPermission();
            updateGuidanceVisibility();
            if (permission !== 'granted') throw new Error('알림 권한이 허용되지 않았습니다.');
            if (!messaging) await initializeMessaging();
            if (!messaging || !registration) throw new Error('알림 기능을 초기화할 수 없습니다.');
            const token = await messaging.getToken({ vapidKey: VAPID_KEY, serviceWorkerRegistration: registration });
            if (!token) throw new Error('알림 토큰을 가져올 수 없습니다.');
            lastMessagingToken = token;
            return token;
        })().finally(() => {
            messagingTokenPromise = null;
        });
        return messagingTokenPromise;
    }

    function markPendingSync(active, keywords) {
        saveToStorage('notificationSyncPending', { active, keywords, updatedAt: new Date().toISOString() });
    }

    function clearPendingSync(active, keywords) {
        const pending = getFromStorage('notificationSyncPending');
        if (pending?.active === active && pending?.keywords === keywords) removeFromStorage('notificationSyncPending');
    }

    async function syncDeviceState({ active, keywords, token }) {
        identity = identity || (active
            ? await window.BWANotificationStore.getOrCreateIdentity()
            : await window.BWANotificationStore.getIdentity());
        if (!identity) return;

        if (active) {
            await withNetworkRetry(() => post(ENDPOINTS.register, {
                deviceId: identity.deviceId,
                deviceSecret: identity.deviceSecret,
                token,
                userAgent: navigator.userAgent,
                keywords,
                active: true
            }));
        } else {
            await withNetworkRetry(() => post(ENDPOINTS.active, {
                deviceId: identity.deviceId,
                deviceSecret: identity.deviceSecret,
                active: false,
                keywords
            }));
        }
        clearPendingSync(active, keywords);
    }

    async function registerActiveDevice(keywords, {
        requireStillActive = false,
        requestPermission = false
    } = {}) {
        const token = await getMessagingToken({ requestPermission });
        if (requireStillActive && getFromStorage('isNotificationActive') !== true) return;
        identity = identity || await window.BWANotificationStore.getOrCreateIdentity();
        await queueDeviceSync(() => syncDeviceState({ active: true, keywords, token }));
    }

    async function deactivateDevice(keywords) {
        identity = identity || await window.BWANotificationStore.getIdentity();
        if (!identity) return;
        await queueDeviceSync(() => syncDeviceState({ active: false, keywords, token: '' }));
    }

    async function finishDeactivation(keywords, operationId) {
        try {
            await deactivateDevice(keywords);
            clearPendingSync(false, keywords);
            if (operationId === toggleOperationId && getFromStorage('isNotificationActive') !== true) {
                setFeedback('toggleStatus', '');
            }
        } catch (error) {
            if (operationId === toggleOperationId && getFromStorage('isNotificationActive') !== true) {
                markPendingSync(false, keywords);
                setFeedback('toggleStatus', '알림은 꺼졌습니다. 서버 반영은 다음 접속 때 자동으로 다시 시도합니다.', 'error');
            }
            console.warn('Notification deactivation will be retried', friendlyError(error));
        } finally {
            if (operationId === toggleOperationId && getFromStorage('isNotificationActive') !== true) {
                renderToggleState(false, document.getElementById('toggleStatus').textContent,
                    document.getElementById('toggleStatus').classList.contains('error') ? 'error' : 'muted');
            }
            void refresh().catch(() => {});
        }
    }

    async function handleToggleChange(event) {
        const operationId = ++toggleOperationId;
        const requestedActive = event.currentTarget.checked;
        const keywords = String(getFromStorage('userKeywords', '') || '').trim();

        if (!requestedActive) {
            saveToStorage('isNotificationActive', false);
            markPendingSync(false, keywords);
            renderToggleState(false);
            void finishDeactivation(keywords, operationId);
            return;
        }

        iosGuideRequested = isIos() && !isStandalone();
        if (iosGuideRequested) {
            saveToStorage('isNotificationActive', false);
            renderToggleState(false, 'iPhone에서는 홈 화면에 추가한 웹앱에서 알림을 켜 주세요.', 'error');
            return;
        }

        saveToStorage('isNotificationActive', true);
        markPendingSync(true, keywords);
        renderToggleState(true);
        void finishActivation(keywords, operationId);
    }

    async function finishActivation(keywords, operationId) {
        try {
            await registerActiveDevice(keywords, {
                requireStillActive: true,
                requestPermission: true
            });
            if (operationId !== toggleOperationId || getFromStorage('isNotificationActive') !== true) return;
            clearPendingSync(true, keywords);
            renderToggleState(true);
            void refresh().catch(() => {});
        } catch (error) {
            if (operationId !== toggleOperationId || getFromStorage('isNotificationActive') !== true) return;
            saveToStorage('isNotificationActive', false);
            markPendingSync(false, keywords);
            renderToggleState(false, `알림을 켜지 못했습니다. ${friendlyError(error)}`, 'error');
            void deactivateDevice(keywords).then(() => clearPendingSync(false, keywords)).catch(() => {});
        }
    }

    function currentKeywords() {
        return document.getElementById('keywordInput').value.trim();
    }

    async function flushKeywordSave() {
        window.clearTimeout(keywordSaveTimer);
        keywordSaveTimer = null;
        if (getFromStorage('isNotificationActive') !== true) return;
        if (keywordSyncInFlight) {
            keywordSyncQueued = true;
            return keywordSyncInFlight;
        }

        keywordSyncInFlight = (async () => {
            do {
                keywordSyncQueued = false;
                const keywords = currentKeywords();
                saveToStorage('userKeywords', keywords);
                if (keywords === lastSyncedKeywords && !getFromStorage('notificationSyncPending')) continue;
                markPendingSync(true, keywords);
                setFeedback('keywordStatus', '자동 저장 중…');
                try {
                    await registerActiveDevice(keywords, { requireStillActive: true });
                    if (getFromStorage('isNotificationActive') !== true) return;
                    lastSyncedKeywords = keywords;
                    clearPendingSync(true, keywords);
                    setFeedback('keywordStatus', '자동 저장되었습니다.', 'success');
                } catch (error) {
                    setFeedback('keywordStatus', `자동 저장을 완료하지 못했습니다. ${friendlyError(error)}`, 'error');
                    return;
                }
            } while (keywordSyncQueued || currentKeywords() !== lastSyncedKeywords);
        })().finally(() => {
            keywordSyncInFlight = null;
            if (keywordSyncQueued && getFromStorage('isNotificationActive') === true) {
                keywordSyncQueued = false;
                keywordSaveTimer = window.setTimeout(() => void flushKeywordSave(), KEYWORD_SAVE_DELAY_MS);
            }
        });
        return keywordSyncInFlight;
    }

    function handleKeywordInput() {
        const keywords = currentKeywords();
        saveToStorage('userKeywords', keywords);
        markPendingSync(true, keywords);
        setFeedback('keywordStatus', '입력을 마치면 자동 저장됩니다.', 'muted');
        window.clearTimeout(keywordSaveTimer);
        keywordSaveTimer = window.setTimeout(() => void flushKeywordSave(), KEYWORD_SAVE_DELAY_MS);
    }

    function sendKeywordKeepalive() {
        if (getFromStorage('isNotificationActive') !== true || !identity) return;
        const keywords = currentKeywords();
        const token = lastMessagingToken;
        if (!token || keywords === lastSyncedKeywords) return;
        const common = {
            deviceId: identity.deviceId,
            token,
            userAgent: navigator.userAgent,
            keywords
        };
        fetch(ENDPOINTS.register, {
            method: 'POST',
            keepalive: true,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ...common, deviceSecret: identity.deviceSecret, active: true })
        }).catch(() => {});
    }

    function openInfo(button) {
        const overlay = document.getElementById('settingsInfoOverlay');
        setText('settingsInfoTitle', button.dataset.infoTitle);
        setText('settingsInfoText', button.dataset.infoText);
        overlay.hidden = false;
        overlay.querySelector('.info-modal').focus();
    }

    function closeInfo() {
        document.getElementById('settingsInfoOverlay').hidden = true;
    }

    function initializeInfoPopups() {
        document.querySelectorAll('.info-button').forEach((button) => {
            button.addEventListener('click', () => openInfo(button));
        });
        document.getElementById('closeSettingsInfoBtn').addEventListener('click', closeInfo);
        document.getElementById('settingsInfoOverlay').addEventListener('click', (event) => {
            if (event.target === event.currentTarget) closeInfo();
        });
        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') closeInfo();
        });
    }

    function renderHealth(server, heartbeat, events) {
        const permission = permissionValue();
        setText('permissionValue', permission);
        setText('registrationValue', server?.active ? '활성' : server ? '비활성' : '등록 정보 없음');
        setText('tokenValue', server?.tokenFingerprint ? `…${server.tokenFingerprint}` : '기록 없음');
        setText('heartbeatValue', formatTime(heartbeat?.value?.receivedAt || server?.lastHeartbeatReceivedAt));
        setText('receivedValue', formatTime(server?.lastNotificationReceivedAt || events[0]?.receivedAt));
        setText('shownValue', formatTime(server?.lastShownAt));
        setText('clickedValue', formatTime(server?.lastClickedAt));
        setText('failureValue', server?.lastFailureCode || '없음');

        const banner = document.getElementById('overallHealth');
        const heartbeatAt = timestampMillis(heartbeat?.value?.receivedAt || server?.lastHeartbeatReceivedAt);
        const heartbeatIsFresh = heartbeatAt && Date.now() - heartbeatAt < 36 * 60 * 60 * 1000;
        if (permission !== '허용됨') {
            banner.className = 'health-banner error';
            banner.textContent = permission === '지원하지 않음'
                ? '이 브라우저는 알림을 지원하지 않습니다.'
                : '알림 권한이 허용되어 있지 않습니다.';
        } else if (!server?.active) {
            banner.className = 'health-banner error';
            banner.textContent = '이 브라우저가 알림 서버에 활성 기기로 등록되어 있지 않습니다.';
        } else if (server.lastFailureCode) {
            banner.className = 'health-banner error';
            banner.textContent = `최근 서버 발송 오류가 있습니다: ${server.lastFailureCode}`;
        } else if (!heartbeatAt) {
            banner.className = 'health-banner warning';
            banner.textContent = '등록은 정상입니다. 첫 일일 점검 신호 수신을 기다리고 있습니다.';
        } else if (!heartbeatIsFresh) {
            banner.className = 'health-banner warning';
            banner.textContent = '마지막 일일 점검 신호가 36시간 이상 지났습니다.';
        } else {
            banner.className = 'health-banner healthy';
            banner.textContent = '알림 등록과 최근 일일 점검 수신이 정상입니다.';
        }
    }

    async function refresh() {
        identity = await window.BWANotificationStore.getIdentity();
        const [events, heartbeat] = await Promise.all([
            window.BWANotificationStore.listEvents(),
            window.BWANotificationStore.getMeta('lastHeartbeat')
        ]);
        let server = null;
        if (identity) {
            try {
                server = await post(ENDPOINTS.status, {
                    deviceId: identity.deviceId,
                    deviceSecret: identity.deviceSecret
                });
            } catch (error) {
                if (error.message !== 'UNAUTHORIZED_DEVICE') throw error;
            }
        }
        renderHealth(server, heartbeat, events);
        return { server, events };
    }

    async function retryPendingSync() {
        const pending = getFromStorage('notificationSyncPending');
        const active = getFromStorage('isNotificationActive') === true;
        if (!pending || pending.active !== active) return;
        try {
            if (active) {
                await registerActiveDevice(String(pending.keywords || ''), { requireStillActive: true });
            } else {
                await deactivateDevice(String(pending.keywords || ''));
            }
            clearPendingSync(active, String(pending.keywords || ''));
        } catch (_) {
            // The pending state remains in local storage and will be retried next time.
        }
    }

    async function runSelfTest() {
        const button = document.getElementById('selfTestBtn');
        button.disabled = true;
        try {
            if (getFromStorage('isNotificationActive') !== true) throw new Error('먼저 경조사 알림을 켜 주세요.');
            identity = identity || await window.BWANotificationStore.getIdentity();
            if (!identity) throw new Error('먼저 경조사 알림을 켜 주세요.');
            setFeedback('selfTestStatus', '해피휴게더 최신 글을 확인하고 알림을 보내는 중…');
            const result = await post(ENDPOINTS.selfTest, {
                deviceId: identity.deviceId,
                deviceSecret: identity.deviceSecret
            });
            setFeedback('selfTestStatus', `“${result.postTitle}” 발송 완료 · 실제 수신 확인 중…`);
            const deadline = Date.now() + 20000;
            while (Date.now() < deadline) {
                await delay(1000);
                const events = await window.BWANotificationStore.listEvents();
                if (events.some((event) => event.eventId === result.eventId)) {
                    setFeedback('selfTestStatus', `해피휴게더 최신 글 “${result.postTitle}” 알림을 실제로 수신했습니다.`, 'success');
                    await refresh();
                    return;
                }
            }
            setFeedback('selfTestStatus', '해피휴게더 조회와 서버 발송은 성공했지만 20초 안에 브라우저 수신을 확인하지 못했습니다.', 'error');
        } catch (error) {
            setFeedback('selfTestStatus', `알림 테스트 실패: ${friendlyError(error)}`, 'error');
        } finally {
            button.disabled = false;
        }
    }

    window.addEventListener('DOMContentLoaded', async () => {
        const active = getFromStorage('isNotificationActive') === true;
        const savedKeywords = String(getFromStorage('userKeywords', '') || '');
        lastSyncedKeywords = getFromStorage('notificationSyncPending') ? '' : savedKeywords;
        document.getElementById('keywordInput').value = savedKeywords;
        renderToggleState(active);
        initializeInfoPopups();

        document.getElementById('notificationToggleSettings').addEventListener('change', (event) => void handleToggleChange(event));
        document.getElementById('permissionActionBtn').addEventListener('click', () => void handlePermissionAction());
        document.getElementById('keywordInput').addEventListener('input', handleKeywordInput);
        document.getElementById('keywordInput').addEventListener('blur', () => void flushKeywordSave());
        document.getElementById('keywordInput').addEventListener('keydown', (event) => {
            if (event.key === 'Enter') event.currentTarget.blur();
        });
        document.getElementById('refreshBtn').addEventListener('click', () => void refresh());
        document.getElementById('selfTestBtn').addEventListener('click', () => void runSelfTest());
        window.addEventListener('pagehide', sendKeywordKeepalive);
        document.addEventListener('visibilitychange', handlePermissionReturn);

        try {
            await initializeMessaging();
            identity = await window.BWANotificationStore.getIdentity();
            if ('Notification' in window && Notification.permission === 'granted' && !lastMessagingToken) {
                void getMessagingToken().catch(() => {});
            }
            void retryPendingSync();
            await refresh();
        } catch (error) {
            const banner = document.getElementById('overallHealth');
            banner.className = 'health-banner error';
            banner.textContent = `상태 확인 실패: ${friendlyError(error)}`;
        }
    });
})();
