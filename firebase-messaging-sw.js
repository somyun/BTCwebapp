'use strict';

importScripts('./notification-store.js?v=production-notifications-1');
importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/9.22.0/firebase-messaging-compat.js');

const EXPECTED_SCOPE_PATH = self.location.hostname === 'somyun.github.io' ? '/BTCwebapp/' : '/';
const ACK_URL = 'https://asia-northeast3-btcwebapp-551bd.cloudfunctions.net/acknowledgeNotification';
const PRODUCTION_FIREBASE_CONFIG = Object.freeze({
    apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
    authDomain: 'btcwebapp-551bd.firebaseapp.com',
    projectId: 'btcwebapp-551bd',
    storageBucket: 'btcwebapp-551bd.firebasestorage.app',
    messagingSenderId: '237989935469',
    appId: '1:237989935469:web:07fc002a5c2ab2f5858264'
});

firebase.initializeApp(PRODUCTION_FIREBASE_CONFIG);
const messaging = firebase.messaging();

async function acknowledge(event, phase) {
    try {
        const identity = await self.BWANotificationStore.getIdentity();
        if (!identity) return;
        const response = await fetch(ACK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: identity.deviceId,
                deviceSecret: identity.deviceSecret,
                eventId: event.eventId,
                type: event.type,
                phase
            })
        });
        if (!response.ok) throw new Error(`ACK_HTTP_${response.status}`);
    } catch (error) {
        console.warn('Notification receipt could not be sent', error);
    }
}

async function processBackgroundMessage(payload) {
    const event = self.BWANotificationStore.normalizeEvent(payload, { source: 'background' });
    await self.BWANotificationStore.saveEvent(event);
    await acknowledge(event, 'received');

    if (event.type === 'heartbeat') {
        await self.BWANotificationStore.setMeta('lastHeartbeat', {
            eventId: event.eventId,
            receivedAt: event.receivedAt
        });
        return;
    }

    const icon = payload?.data?.icon || payload?.notification?.icon || './icon-192.png';
    await self.registration.showNotification(event.title, {
        body: event.body,
        icon,
        tag: event.eventId,
        data: {
            eventId: event.eventId,
            type: event.type,
            url: event.url || `${self.location.origin}${EXPECTED_SCOPE_PATH}`
        }
    });
    const shownAt = new Date().toISOString();
    await self.BWANotificationStore.patchEvent(event.eventId, { shownAt });
    await acknowledge(event, 'shown');
}

self.addEventListener('install', () => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        const scopePath = new URL(self.registration.scope).pathname;
        if (scopePath !== EXPECTED_SCOPE_PATH) {
            await self.registration.unregister();
            return;
        }
        await self.clients.claim();
    })());
});

messaging.onBackgroundMessage((payload) => processBackgroundMessage(payload));

self.addEventListener('notificationclick', (event) => {
    const details = event.notification.data || {};
    event.notification.close();
    event.waitUntil((async () => {
        const clickedAt = new Date().toISOString();
        if (details.eventId) {
            await self.BWANotificationStore.patchEvent(details.eventId, { clickedAt });
            await acknowledge({ eventId: details.eventId, type: details.type || 'notification' }, 'clicked');
        }
        await self.clients.openWindow(details.url || `${self.location.origin}${EXPECTED_SCOPE_PATH}`);
    })());
});

// This worker is scoped only to the production BTCwebapp path and does not cache app assets.
