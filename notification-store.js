(function attachNotificationStore(globalScope) {
    'use strict';

    const DB_NAME = 'btcwebapp-production-notifications-v1';
    const DB_VERSION = 1;
    const IDENTITY_STORE = 'identity';
    const EVENT_STORE = 'events';
    const META_STORE = 'meta';
    const MAX_EVENTS = 100;

    function randomBase64Url(byteLength) {
        const bytes = new Uint8Array(byteLength);
        globalScope.crypto.getRandomValues(bytes);
        let binary = '';
        bytes.forEach((value) => { binary += String.fromCharCode(value); });
        return globalScope.btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
    }

    function openDatabase() {
        return new Promise((resolve, reject) => {
            const request = globalScope.indexedDB.open(DB_NAME, DB_VERSION);
            request.onerror = () => reject(request.error || new Error('NOTIFICATION_DB_OPEN_FAILED'));
            request.onsuccess = () => resolve(request.result);
            request.onupgradeneeded = () => {
                const database = request.result;
                if (!database.objectStoreNames.contains(IDENTITY_STORE)) {
                    database.createObjectStore(IDENTITY_STORE, { keyPath: 'id' });
                }
                if (!database.objectStoreNames.contains(EVENT_STORE)) {
                    const events = database.createObjectStore(EVENT_STORE, { keyPath: 'eventId' });
                    events.createIndex('receivedAt', 'receivedAt');
                }
                if (!database.objectStoreNames.contains(META_STORE)) {
                    database.createObjectStore(META_STORE, { keyPath: 'key' });
                }
            };
        });
    }

    async function runTransaction(storeNames, mode, operation) {
        const database = await openDatabase();
        try {
            return await new Promise((resolve, reject) => {
                const transaction = database.transaction(storeNames, mode);
                let result;
                transaction.oncomplete = () => resolve(result);
                transaction.onerror = () => reject(transaction.error || new Error('NOTIFICATION_DB_TRANSACTION_FAILED'));
                transaction.onabort = () => reject(transaction.error || new Error('NOTIFICATION_DB_TRANSACTION_ABORTED'));
                result = operation(transaction);
            });
        } finally {
            database.close();
        }
    }

    function requestResult(request) {
        return new Promise((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error || new Error('NOTIFICATION_DB_REQUEST_FAILED'));
        });
    }

    async function getIdentity() {
        const database = await openDatabase();
        try {
            const transaction = database.transaction(IDENTITY_STORE, 'readonly');
            return await requestResult(transaction.objectStore(IDENTITY_STORE).get('current')) || null;
        } finally {
            database.close();
        }
    }

    async function getOrCreateIdentity() {
        const existing = await getIdentity();
        if (existing?.deviceId && existing?.deviceSecret) return existing;
        const identity = {
            id: 'current',
            deviceId: `d_${randomBase64Url(32)}`,
            deviceSecret: randomBase64Url(48),
            createdAt: new Date().toISOString()
        };
        await runTransaction([IDENTITY_STORE], 'readwrite', (transaction) => {
            transaction.objectStore(IDENTITY_STORE).put(identity);
        });
        return identity;
    }

    function safeEventId(payload = {}) {
        const candidate = String(payload.eventId || payload.messageId || '').trim();
        if (/^[A-Za-z0-9._:-]{8,160}$/.test(candidate)) return candidate;
        return `local:${Date.now()}:${randomBase64Url(12)}`;
    }

    function normalizeEvent(payload = {}, overrides = {}) {
        const data = payload.data || payload;
        const notification = payload.notification || {};
        const now = new Date().toISOString();
        return {
            eventId: safeEventId({ eventId: data.eventId, messageId: payload.messageId }),
            type: String(data.type || overrides.type || 'notification').slice(0, 40),
            title: String(data.title || notification.title || overrides.title || '알림').slice(0, 200),
            body: String(data.body || notification.body || overrides.body || '').slice(0, 1000),
            url: String(data.url || overrides.url || './').slice(0, 1000),
            receivedAt: overrides.receivedAt || now,
            shownAt: overrides.shownAt || null,
            clickedAt: overrides.clickedAt || null,
            source: String(overrides.source || 'fcm').slice(0, 40)
        };
    }

    async function saveEvent(event) {
        const normalized = normalizeEvent(event, event);
        const database = await openDatabase();
        try {
            const transaction = database.transaction(EVENT_STORE, 'readwrite');
            const store = transaction.objectStore(EVENT_STORE);
            const previous = await requestResult(store.get(normalized.eventId));
            store.put({ ...previous, ...normalized, receivedAt: previous?.receivedAt || normalized.receivedAt });
            await new Promise((resolve, reject) => {
                transaction.oncomplete = resolve;
                transaction.onerror = () => reject(transaction.error);
                transaction.onabort = () => reject(transaction.error);
            });
        } finally {
            database.close();
        }
        await trimEvents();
        return normalized;
    }

    async function patchEvent(eventId, updates) {
        const database = await openDatabase();
        try {
            const transaction = database.transaction(EVENT_STORE, 'readwrite');
            const store = transaction.objectStore(EVENT_STORE);
            const previous = await requestResult(store.get(eventId));
            if (previous) store.put({ ...previous, ...updates, eventId });
            await new Promise((resolve, reject) => {
                transaction.oncomplete = resolve;
                transaction.onerror = () => reject(transaction.error);
                transaction.onabort = () => reject(transaction.error);
            });
        } finally {
            database.close();
        }
    }

    async function listEvents(limit = MAX_EVENTS) {
        const database = await openDatabase();
        try {
            const transaction = database.transaction(EVENT_STORE, 'readonly');
            const events = await requestResult(transaction.objectStore(EVENT_STORE).getAll());
            return events
                .sort((left, right) => String(right.receivedAt).localeCompare(String(left.receivedAt)))
                .slice(0, Math.max(1, Math.min(MAX_EVENTS, Number(limit) || MAX_EVENTS)));
        } finally {
            database.close();
        }
    }

    async function trimEvents() {
        const database = await openDatabase();
        try {
            const readTransaction = database.transaction(EVENT_STORE, 'readonly');
            const events = await requestResult(readTransaction.objectStore(EVENT_STORE).getAll());
            const stale = events
                .sort((left, right) => String(right.receivedAt).localeCompare(String(left.receivedAt)))
                .slice(MAX_EVENTS);
            if (!stale.length) return;
            const writeTransaction = database.transaction(EVENT_STORE, 'readwrite');
            const store = writeTransaction.objectStore(EVENT_STORE);
            stale.forEach((event) => store.delete(event.eventId));
            await new Promise((resolve, reject) => {
                writeTransaction.oncomplete = resolve;
                writeTransaction.onerror = () => reject(writeTransaction.error);
                writeTransaction.onabort = () => reject(writeTransaction.error);
            });
        } finally {
            database.close();
        }
    }

    async function setMeta(key, value) {
        await runTransaction([META_STORE], 'readwrite', (transaction) => {
            transaction.objectStore(META_STORE).put({ key, value, updatedAt: new Date().toISOString() });
        });
    }

    async function getMeta(key) {
        const database = await openDatabase();
        try {
            const transaction = database.transaction(META_STORE, 'readonly');
            return await requestResult(transaction.objectStore(META_STORE).get(key)) || null;
        } finally {
            database.close();
        }
    }

    globalScope.BWANotificationStore = Object.freeze({
        getIdentity,
        getMeta,
        getOrCreateIdentity,
        listEvents,
        normalizeEvent,
        patchEvent,
        saveEvent,
        setMeta
    });
})(typeof self !== 'undefined' ? self : window);
