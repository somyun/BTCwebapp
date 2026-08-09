import {
    getBytes,
    getStorage,
    ref
} from 'https://www.gstatic.com/firebasejs/9.22.0/firebase-storage.js';

const DEFAULT_FIREBASE = Object.freeze({
    projectId: 'btcwebapp-551bd',
    storageBucket: 'btcwebapp-551bd.firebasestorage.app'
});
const CAD_ROOT = 'cad/hopo';
const DEFAULT_MAX_BYTES = 4 * 1024 * 1024;
const HUMETRO_EMAIL = /@humetro\.busan\.kr$/i;

function runtimeFirebaseConfig() {
    return window.BWA_AUTH_CONFIG?.firebase ||
        window.BWA_TEST_CONFIG?.firebase ||
        DEFAULT_FIREBASE;
}

function firebaseAppForProject(projectId) {
    if (!window.firebase?.apps?.length) throw new Error('FIREBASE_APP_UNAVAILABLE');
    const app = window.firebase.apps.find((candidate) => candidate.options?.projectId === projectId);
    if (!app) throw new Error('FIREBASE_PROJECT_APP_UNAVAILABLE');
    return app;
}

function safeRelativePath(value) {
    const path = String(value || '').replaceAll('\\', '/').replace(/^\/+/, '');
    if (!path || path.split('/').some((segment) => !segment || segment === '.' || segment === '..') ||
        path.includes('://')) {
        throw new Error('INVALID_CAD_STORAGE_PATH');
    }
    return path;
}

async function authorizedStorage() {
    const config = runtimeFirebaseConfig();
    const app = firebaseAppForProject(config.projectId);
    const user = app.auth().currentUser;
    if (!user) throw new Error('CAD_AUTH_REQUIRED');
    const token = await user.getIdTokenResult();
    const email = String(token.claims.humetroEmail || user.email || '').toLowerCase();
    if (token.claims.humetro !== true || !HUMETRO_EMAIL.test(email)) {
        throw new Error('CAD_AUTH_FORBIDDEN');
    }
    return getStorage(app._delegate, `gs://${config.storageBucket}`);
}

async function readJson(relativePath, { maxBytes = DEFAULT_MAX_BYTES } = {}) {
    const path = safeRelativePath(relativePath);
    const storage = await authorizedStorage();
    const bytes = await getBytes(ref(storage, `${CAD_ROOT}/${path}`), maxBytes);
    try {
        return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
    } catch (error) {
        throw new Error(`INVALID_CAD_JSON:${path}`, { cause: error });
    }
}

window.BWACadStorage = Object.freeze({
    readJson,
    rootPath: CAD_ROOT
});
window.dispatchEvent(new CustomEvent('bwa-cad-storage-ready'));
