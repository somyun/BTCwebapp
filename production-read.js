(function attachProductionRead(root) {
    'use strict';

    const adapter = root.BWAReadAdapter;
    if (!adapter) throw new Error('READ_ADAPTER_UNAVAILABLE');

    const CONFIG = Object.freeze({
        defaultSource: 'firestore',
        projectId: 'btcwebapp-551bd',
        apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
        spreadsheetId: '19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s',
        requestTimeoutMs: 10000
    });
    const firestoreBase = `https://firestore.googleapis.com/v1/projects/${CONFIG.projectId}` +
        '/databases/(default)/documents';
    const state = {
        source: CONFIG.defaultSource,
        requests: []
    };
    let formsBySheetName = new Map();

    function seoulDateKey(value = new Date()) {
        const parts = new Intl.DateTimeFormat('en-US', {
            timeZone: 'Asia/Seoul',
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        }).formatToParts(value);
        const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
        return `${values.year}-${values.month}-${values.day}`;
    }

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

    async function fetchOptionalDocument(pathSegments) {
        try {
            return await fetchDocument(pathSegments);
        } catch (error) {
            if (error.message === 'HTTP_404') return null;
            throw error;
        }
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
            const chunkPath = document.activeRevisionId
                ? ['publicForms', formListItem.formKey, 'revisions', document.activeRevisionId, 'chunks']
                : ['publicForms', formListItem.formKey, 'chunks'];
            const payload = await fetchJson(firestoreUrl(
                chunkPath,
                { pageSize: 1000 }
            ));
            const chunks = (payload.documents || [])
                .map((entry) => adapter.decodeFirestoreFields(entry.fields || {}))
                .sort((left, right) => left.index - right.index);
            if (chunks.length !== document.chunkCount) throw new Error('FIRESTORE_CHUNK_COUNT_MISMATCH');
            if (chunks.some((chunk) => chunk.formKey !== document.formKey ||
                chunk.contentHash !== document.contentHash)) {
                throw new Error('FIRESTORE_CHUNK_REVISION_MISMATCH');
            }
            rows = chunks.flatMap((chunk) => chunk.rows);
        } else {
            throw new Error('INVALID_FIRESTORE_STORAGE_MODE');
        }
        const normalizedRows = adapter.normalizeRows(rows);
        if (normalizedRows.length !== document.rowCount) throw new Error('FIRESTORE_ROW_COUNT_MISMATCH');
        return { document, rows: normalizedRows };
    }

    function rowIdentity(row) {
        return row.uniqueId
            ? `id:${row.uniqueId}`
            : `pair:${row.location.toLocaleLowerCase('ko-KR')}|${row.item.toLocaleLowerCase('ko-KR')}`;
    }

    async function fetchTodayMeasurements(formListItem, rows) {
        const cacheDate = seoulDateKey();
        const cacheId = `${formListItem.formKey}_${cacheDate}`;
        const document = await fetchOptionalDocument(['dailyMeasurementCaches', cacheId]);
        if (!document) return { dailyCache: null, rows };
        if (document.schemaVersion !== adapter.SCHEMA_VERSION ||
            document.formKey !== formListItem.formKey ||
            document.sheetName !== formListItem.sheetName ||
            document.cacheDate !== cacheDate ||
            !Array.isArray(document.measurements) ||
            document.measurementCount !== document.measurements.length) {
            throw new Error('INVALID_DAILY_MEASUREMENT_CACHE');
        }
        const measurements = adapter.normalizeRows(document.measurements);
        if (measurements.length !== rows.length) throw new Error('DAILY_CACHE_ROW_COUNT_MISMATCH');
        const values = new Map(measurements.map((measurement) => [rowIdentity(measurement), measurement]));
        const mergedRows = rows.map((row) => {
            const measurement = values.get(rowIdentity(row));
            if (!measurement || measurement.uniqueId !== row.uniqueId ||
                measurement.location !== row.location || measurement.item !== row.item ||
                measurement.unit !== row.unit) {
                throw new Error('DAILY_CACHE_ROW_IDENTITY_MISMATCH');
            }
            return { ...row, value: measurement.value };
        });
        return { dailyCache: document, rows: mergedRows };
    }

    async function loadFormList() {
        const { items } = await fetchFormListDocument();
        return { items, servedBy: 'firestore' };
    }

    async function loadForm(sheetName, selectedForm) {
        const formListItem = formsBySheetName.get(sheetName) || selectedForm;
        const result = await fetchFormDocument(formListItem);
        const today = await fetchTodayMeasurements(formListItem, result.rows);
        return {
            rows: today.rows,
            servedBy: 'firestore',
            document: result.document,
            dailyCache: today.dailyCache
        };
    }

    root.BWA_PRODUCTION_READ_STATE = state;
    root.BWAProductionRead = Object.freeze({
        CONFIG,
        loadForm,
        loadFormList
    });
})(window);
