(function attachReadAdapter(root, factory) {
    const adapter = factory();
    if (typeof module === 'object' && module.exports) module.exports = adapter;
    if (root) root.BWAReadAdapter = adapter;
})(typeof window !== 'undefined' ? window : null, function createReadAdapter() {
    'use strict';

    const SCHEMA_VERSION = 1;

    function stableStringify(value) {
        if (Array.isArray(value)) {
            return `[${value.map((entry) => stableStringify(entry)).join(',')}]`;
        }
        if (value && typeof value === 'object') {
            return `{${Object.keys(value)
                .sort()
                .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
                .join(',')}}`;
        }
        return JSON.stringify(value);
    }

    function requiredString(value, fieldName) {
        const normalized = String(value ?? '').trim().normalize('NFC');
        if (!normalized) throw new Error(`INVALID_${fieldName.toUpperCase()}`);
        return normalized;
    }

    function optionalString(value) {
        if (value === null || value === undefined) return null;
        return String(value).normalize('NFC');
    }

    function normalizeRevision(value) {
        const text = requiredString(value, 'lastModifiedDate');
        const timestamp = Date.parse(text);
        if (!Number.isFinite(timestamp)) throw new Error('INVALID_LASTMODIFIEDDATE');
        return new Date(timestamp).toISOString();
    }

    function decodeFirestoreValue(value) {
        if (!value || typeof value !== 'object') throw new Error('INVALID_FIRESTORE_VALUE');
        if ('nullValue' in value) return null;
        if ('stringValue' in value) return value.stringValue;
        if ('integerValue' in value) return Number(value.integerValue);
        if ('doubleValue' in value) return Number(value.doubleValue);
        if ('booleanValue' in value) return Boolean(value.booleanValue);
        if ('timestampValue' in value) return value.timestampValue;
        if ('arrayValue' in value) {
            return (value.arrayValue.values || []).map((entry) => decodeFirestoreValue(entry));
        }
        if ('mapValue' in value) return decodeFirestoreFields(value.mapValue.fields || {});
        throw new Error(`UNSUPPORTED_FIRESTORE_VALUE:${Object.keys(value).join(',')}`);
    }

    function decodeFirestoreFields(fields) {
        if (!fields || typeof fields !== 'object') throw new Error('INVALID_FIRESTORE_FIELDS');
        return Object.fromEntries(
            Object.entries(fields).map(([key, value]) => [key, decodeFirestoreValue(value)])
        );
    }

    function normalizeGasFormList(payload, expectedSpreadsheetId) {
        if (!Array.isArray(payload) || !payload.length) throw new Error('INVALID_GAS_FORM_LIST');
        const seen = new Set();
        return payload.map((entry) => {
            if (!entry || typeof entry !== 'object') throw new Error('INVALID_GAS_FORM_ENTRY');
            if (entry.spreadsheetId !== expectedSpreadsheetId) {
                throw new Error('SOURCE_SPREADSHEET_ID_MISMATCH');
            }
            const sheetName = requiredString(entry.sheetName, 'sheetName');
            if (seen.has(sheetName)) throw new Error('DUPLICATE_GAS_FORM');
            seen.add(sheetName);
            return {
                sheetName,
                displayName: sheetName,
                lastModifiedDate: normalizeRevision(entry.lastModifiedDate)
            };
        });
    }

    function normalizeFirestoreFormList(document) {
        if (!document || typeof document !== 'object') throw new Error('INVALID_FIRESTORE_FORM_LIST');
        if (document.schemaVersion !== SCHEMA_VERSION) throw new Error('UNSUPPORTED_FORM_LIST_SCHEMA');
        if (!Array.isArray(document.items) || !document.items.length) throw new Error('EMPTY_FIRESTORE_FORM_LIST');
        if (document.itemCount !== document.items.length) throw new Error('FIRESTORE_FORM_LIST_COUNT_MISMATCH');
        const seenNames = new Set();
        const seenKeys = new Set();
        const items = document.items.map((entry) => {
            const sheetName = requiredString(entry.sheetName, 'sheetName');
            const formKey = requiredString(entry.formKey, 'formKey');
            if (!/^f_[a-f0-9]{32}$/.test(formKey)) throw new Error('INVALID_FORM_KEY');
            if (seenNames.has(sheetName) || seenKeys.has(formKey)) throw new Error('DUPLICATE_FIRESTORE_FORM');
            seenNames.add(sheetName);
            seenKeys.add(formKey);
            return {
                formKey,
                sheetName,
                displayName: requiredString(entry.displayName, 'displayName'),
                lastModifiedDate: normalizeRevision(entry.lastModifiedDate)
            };
        });
        const sourceRevision = normalizeRevision(document.sourceRevision);
        const latestItemRevision = items.map((item) => item.lastModifiedDate).sort().at(-1);
        if (sourceRevision !== latestItemRevision) throw new Error('STALE_FIRESTORE_FORM_LIST');
        return items;
    }

    function normalizeValidation(validation) {
        if (validation === null || validation === undefined) return null;
        if (typeof validation !== 'object') throw new Error('INVALID_VALIDATION');
        return {
            minValue: optionalString(validation.minValue),
            maxValue: optionalString(validation.maxValue)
        };
    }

    function normalizeRecentInfo(recentInfo) {
        if (recentInfo === null || recentInfo === undefined) return null;
        if (typeof recentInfo !== 'object') throw new Error('INVALID_RECENT_INFO');
        return {
            value: optionalString(recentInfo.value),
            date: optionalString(recentInfo.date)
        };
    }

    function normalizeRows(payload) {
        if (!Array.isArray(payload) || !payload.length) throw new Error('INVALID_FORM_ROWS');
        const seenIds = new Set();
        return payload.map((entry) => {
            if (!entry || typeof entry !== 'object') throw new Error('INVALID_FORM_ROW');
            const uniqueId = optionalString(entry.uniqueId) ?? '';
            if (uniqueId && seenIds.has(uniqueId)) throw new Error('DUPLICATE_UNIQUE_ID');
            if (uniqueId) seenIds.add(uniqueId);
            return {
                uniqueId,
                location: optionalString(entry.location) ?? '',
                item: optionalString(entry.item) ?? '',
                value: optionalString(entry.value) ?? '',
                unit: optionalString(entry.unit) ?? '',
                validation: normalizeValidation(entry.validation),
                recentInfo: normalizeRecentInfo(entry.recentInfo)
            };
        });
    }

    function compareFormLists(gasPayload, firestoreDocument, expectedSpreadsheetId) {
        const gasItems = normalizeGasFormList(gasPayload, expectedSpreadsheetId);
        const firestoreItems = normalizeFirestoreFormList(firestoreDocument);
        const selectComparable = ({ sheetName, displayName, lastModifiedDate }) => ({
            sheetName,
            displayName,
            lastModifiedDate
        });
        return {
            matched: stableStringify(gasItems.map(selectComparable)) ===
                stableStringify(firestoreItems.map(selectComparable)),
            gasCount: gasItems.length,
            firestoreCount: firestoreItems.length,
            gasItems,
            firestoreItems
        };
    }

    function compareFormRows(gasPayload, firestorePayload) {
        const gasRows = normalizeRows(gasPayload);
        const firestoreRows = normalizeRows(firestorePayload);
        const matched = stableStringify(gasRows) === stableStringify(firestoreRows);
        let firstMismatchIndex = -1;
        if (!matched) {
            const length = Math.max(gasRows.length, firestoreRows.length);
            for (let index = 0; index < length; index += 1) {
                if (stableStringify(gasRows[index]) !== stableStringify(firestoreRows[index])) {
                    firstMismatchIndex = index;
                    break;
                }
            }
        }
        return {
            matched,
            gasCount: gasRows.length,
            firestoreCount: firestoreRows.length,
            firstMismatchIndex,
            gasRows,
            firestoreRows
        };
    }

    return Object.freeze({
        SCHEMA_VERSION,
        compareFormLists,
        compareFormRows,
        decodeFirestoreFields,
        normalizeFirestoreFormList,
        normalizeGasFormList,
        normalizeRows,
        stableStringify
    });
});
