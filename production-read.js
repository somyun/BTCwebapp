(function attachProductionRead(root) {
    'use strict';

    const adapter = root.BWAReadAdapter;
    if (!adapter) throw new Error('READ_ADAPTER_UNAVAILABLE');

    const CONFIG = Object.freeze({
        defaultSource: 'firestore',
        projectId: 'btcwebapp-551bd',
        apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
        spreadsheetId: '19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s',
        requestTimeoutMs: 10000,
        gasFallback: false
    });
    const ALLOWED_SOURCES = new Set(['gas', 'shadow', 'firestore']);
    const requestedSource = new URLSearchParams(root.location.search).get('readSource');
    const isLocalPreview = ['127.0.0.1', 'localhost'].includes(root.location.hostname);
    const source = isLocalPreview && ALLOWED_SOURCES.has(requestedSource)
        ? requestedSource
        : CONFIG.defaultSource;
    const firestoreBase = `https://firestore.googleapis.com/v1/projects/${CONFIG.projectId}` +
        '/databases/(default)/documents';
    const state = {
        source,
        fallbackCount: 0,
        requests: [],
        shadow: { formList: { status: 'idle' }, forms: {} }
    };
    let formsBySheetName = new Map();

    function firestoreUrl(pathSegments, parameters = {}) {
        const path = pathSegments.map((segment) => encodeURIComponent(String(segment))).join('/');
        const url = new URL(`${firestoreBase}/${path}`);
        url.searchParams.set('key', CONFIG.apiKey);
        Object.entries(parameters).forEach(([key, value]) => url.searchParams.set(key, String(value)));
        return url;
    }

    async function fetchJson(url) {
        const controller = new AbortController();
        const timeout = root.setTimeout(() => controller.abort(), CONFIG.requestTimeoutMs);
        const startedAt = Date.now();
        try {
            const response = await root.fetch(url, {
                method: 'GET',
                redirect: 'follow',
                cache: 'no-store',
                headers: { Accept: 'application/json' },
                signal: controller.signal
            });
            if (!response.ok) throw new Error(`HTTP_${response.status}`);
            state.requests.push({
                kind: 'firestore',
                path: new URL(url).pathname,
                status: response.status,
                durationMs: Date.now() - startedAt
            });
            return response.json();
        } catch (error) {
            state.requests.push({
                kind: 'firestore',
                path: new URL(url).pathname,
                error: error.name || 'Error',
                durationMs: Date.now() - startedAt
            });
            throw error;
        } finally {
            root.clearTimeout(timeout);
        }
    }

    async function fetchDocument(pathSegments) {
        const payload = await fetchJson(firestoreUrl(pathSegments));
        return adapter.decodeFirestoreFields(payload.fields || {});
    }

    async function fetchFormListDocument() {
        const document = await fetchDocument(['publicCache', 'formList']);
        const normalized = adapter.normalizeFirestoreFormList(document);
        const items = normalized.map((item) => ({
            ...item,
            spreadsheetId: CONFIG.spreadsheetId
        }));
        formsBySheetName = new Map(items.map((item) => [item.sheetName, item]));
        return { document, items };
    }

    async function fetchFormDocument(formListItem) {
        if (!formListItem?.formKey) throw new Error('FIRESTORE_FORM_KEY_MISSING');
        const document = await fetchDocument(['publicForms', formListItem.formKey]);
        if (document.schemaVersion !== adapter.SCHEMA_VERSION) throw new Error('UNSUPPORTED_FORM_SCHEMA');
        if (document.formKey !== formListItem.formKey || document.sheetName !== formListItem.sheetName) {
            throw new Error('FIRESTORE_FORM_IDENTITY_MISMATCH');
        }
        if (new Date(document.sourceRevision).toISOString() !==
            new Date(formListItem.lastModifiedDate).toISOString()) {
            throw new Error('STALE_FIRESTORE_FORM');
        }
        let rows;
        if (document.storageMode === 'inline') {
            rows = document.rows;
        } else if (document.storageMode === 'chunked') {
            const payload = await fetchJson(firestoreUrl(
                ['publicForms', formListItem.formKey, 'chunks'],
                { pageSize: 1000 }
            ));
            const chunks = (payload.documents || [])
                .map((entry) => adapter.decodeFirestoreFields(entry.fields || {}))
                .sort((left, right) => left.index - right.index);
            if (chunks.length !== document.chunkCount) throw new Error('FIRESTORE_CHUNK_COUNT_MISMATCH');
            rows = chunks.flatMap((chunk) => chunk.rows);
        } else {
            throw new Error('INVALID_FIRESTORE_STORAGE_MODE');
        }
        const normalizedRows = adapter.normalizeRows(rows);
        if (normalizedRows.length !== document.rowCount) throw new Error('FIRESTORE_ROW_COUNT_MISMATCH');
        return { document, rows: normalizedRows };
    }

    async function compareFormListInBackground(gasPayload, firestorePromise) {
        state.shadow.formList = { status: 'pending' };
        try {
            const { document } = await firestorePromise;
            const comparison = adapter.compareFormLists(
                gasPayload,
                document,
                CONFIG.spreadsheetId
            );
            state.shadow.formList = {
                status: comparison.matched ? 'matched' : 'mismatch',
                gasCount: comparison.gasCount,
                firestoreCount: comparison.firestoreCount
            };
        } catch (error) {
            state.shadow.formList = { status: 'error', error: error.message };
        }
    }

    async function compareFormInBackground(sheetName, gasRows, firestorePromise) {
        state.shadow.forms[sheetName] = { status: 'pending' };
        try {
            const firestoreForm = await firestorePromise;
            const comparison = adapter.compareFormRows(gasRows, firestoreForm.rows);
            state.shadow.forms[sheetName] = {
                status: comparison.matched ? 'matched' : 'mismatch',
                gasCount: comparison.gasCount,
                firestoreCount: comparison.firestoreCount,
                firstMismatchIndex: comparison.firstMismatchIndex
            };
        } catch (error) {
            state.shadow.forms[sheetName] = { status: 'error', error: error.message };
        }
    }

    async function loadFormList(gasFetch) {
        if (source === 'gas') return { items: await gasFetch(), servedBy: 'gas' };
        if (source === 'shadow') {
            const firestorePromise = fetchFormListDocument();
            const gasPayload = await gasFetch();
            void compareFormListInBackground(gasPayload, firestorePromise);
            return { items: gasPayload, servedBy: 'gas' };
        }
        try {
            const { items } = await fetchFormListDocument();
            return { items, servedBy: 'firestore' };
        } catch (error) {
            if (!CONFIG.gasFallback) throw error;
            state.fallbackCount += 1;
            state.lastFallback = { operation: 'formList', error: error.message };
            return { items: await gasFetch(), servedBy: 'gas-fallback' };
        }
    }

    async function loadForm(sheetName, selectedForm, gasFetch) {
        if (source === 'gas') return { rows: await gasFetch(), servedBy: 'gas' };
        const formListItem = formsBySheetName.get(sheetName) || selectedForm;
        if (source === 'shadow') {
            const firestorePromise = fetchFormDocument(formListItem);
            const gasRows = await gasFetch();
            void compareFormInBackground(sheetName, gasRows, firestorePromise);
            return { rows: gasRows, servedBy: 'gas' };
        }
        try {
            const result = await fetchFormDocument(formListItem);
            return { rows: result.rows, servedBy: 'firestore', document: result.document };
        } catch (error) {
            if (!CONFIG.gasFallback) throw error;
            state.fallbackCount += 1;
            state.lastFallback = { operation: 'form', sheetName, error: error.message };
            return { rows: await gasFetch(), servedBy: 'gas-fallback' };
        }
    }

    root.BWA_PRODUCTION_READ_STATE = state;
    root.BWAProductionRead = Object.freeze({
        CONFIG,
        loadForm,
        loadFormList
    });
})(window);
