(() => {
    'use strict';

    const ACKNOWLEDGE_URL = 'https://asia-northeast3-btcwebapp-551bd.cloudfunctions.net/acknowledgeNotification';
    const FIREBASE_CONFIG = Object.freeze({
        apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
        authDomain: 'btcwebapp-551bd.firebaseapp.com',
        projectId: 'btcwebapp-551bd',
        storageBucket: 'btcwebapp-551bd.firebasestorage.app',
        messagingSenderId: '237989935469',
        appId: '1:237989935469:web:07fc002a5c2ab2f5858264'
    });
    let identity = null;
    let registration = null;

    async function post(url, body) {
        const response = await fetch(url, {
            method: 'POST',
            cache: 'no-store',
            headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok || payload.ok !== true) throw new Error(payload.error || `HTTP_${response.status}`);
        return payload.result;
    }

    function formatTime(value) {
        const milliseconds = Date.parse(value || '') || 0;
        return milliseconds ? new Date(milliseconds).toLocaleString('ko-KR') : '기록 없음';
    }

    function safeLink(value) {
        try {
            const url = new URL(value || './', window.location.href);
            return ['http:', 'https:'].includes(url.protocol) ? url.href : './';
        } catch (_) {
            return './';
        }
    }

    async function acknowledge(event, phase) {
        try {
            if (!identity) return false;
            await post(ACKNOWLEDGE_URL, {
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

    function renderHistory(events) {
        const list = document.getElementById('notificationHistory');
        const empty = document.getElementById('emptyHistory');
        list.replaceChildren();
        for (const event of events.filter((item) => item.type !== 'heartbeat')) {
            const item = document.createElement('li');
            item.className = 'history-item';
            const link = document.createElement('a');
            link.href = safeLink(event.url);
            link.target = '_blank';
            link.rel = 'noopener noreferrer';
            const title = document.createElement('span');
            title.className = 'history-title';
            title.textContent = event.title || '알림';
            const body = document.createElement('span');
            body.className = 'history-body';
            body.textContent = event.body || '내용 없음';
            const meta = document.createElement('span');
            meta.className = 'history-meta';
            const received = document.createElement('span');
            received.textContent = formatTime(event.receivedAt);
            const type = document.createElement('span');
            type.className = 'history-badge';
            type.textContent = event.type === 'self-test' ? '테스트' : '알림';
            meta.append(received, type);
            if (event.clickedAt) {
                const clicked = document.createElement('span');
                clicked.textContent = '열어봄';
                meta.appendChild(clicked);
            }
            link.append(title, body, meta);
            item.appendChild(link);
            list.appendChild(item);
        }
        empty.hidden = list.childElementCount > 0;
    }

    async function refreshHistory() {
        renderHistory(await window.BWANotificationStore.listEvents());
    }

    async function processForeground(payload) {
        const event = window.BWANotificationStore.normalizeEvent(payload, { source: 'notification-history-page' });
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
        await refreshHistory();
    }

    async function initializeMessaging() {
        if (!('serviceWorker' in navigator) || !window.firebase?.messaging) return;
        registration = await navigator.serviceWorker.register('./firebase-messaging-sw.js', { scope: './' });
        if (!window.firebase.apps.length) window.firebase.initializeApp(FIREBASE_CONFIG);
        window.firebase.messaging().onMessage((payload) => void processForeground(payload));
    }

    window.addEventListener('DOMContentLoaded', async () => {
        identity = await window.BWANotificationStore.getIdentity();
        await refreshHistory();
        try {
            await initializeMessaging();
        } catch (error) {
            console.warn('알림 수신 초기화 실패', error);
        }
    });
})();
