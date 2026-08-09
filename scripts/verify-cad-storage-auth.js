const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {
    cert,
    deleteApp,
    initializeApp
} = require('../firebase-production/functions/node_modules/firebase-admin/lib/app');
const { getAuth } = require('../firebase-production/functions/node_modules/firebase-admin/lib/auth');

const ROOT = path.resolve(__dirname, '..');
const PROJECT_ID = 'btcwebapp-551bd';
const BUCKET = `${PROJECT_ID}.firebasestorage.app`;
const MANIFEST_OBJECT = 'cad/hopo/manifest.json';
const VERIFY_EMAIL = 'storage-verifier@humetro.busan.kr';

function loadServiceAccount() {
    const source = fs.readFileSync(path.join(ROOT, 'apps-script', 'firebase_key.js'), 'utf8');
    const context = Object.create(null);
    vm.createContext(context);
    vm.runInContext(`${source}\nthis.__serviceAccount = SERVICE_ACCOUNT_KEY;`, context, {
        filename: 'apps-script/firebase_key.js',
        timeout: 1000
    });
    const account = context.__serviceAccount;
    if (!account?.project_id || !account?.client_email || !account?.private_key) {
        throw new Error('Production Firebase service account is incomplete.');
    }
    if (account.project_id !== PROJECT_ID) {
        throw new Error(`Unexpected service account project: ${account.project_id}`);
    }
    return account;
}

function loadWebApiKey() {
    const source = fs.readFileSync(path.join(ROOT, 'auth.js'), 'utf8');
    const match = source.match(/apiKey:\s*'([^']+)'/u);
    if (!match) throw new Error('Firebase Web API key was not found in auth.js.');
    return match[1];
}

function mediaUrl() {
    const objectName = encodeURIComponent(MANIFEST_OBJECT);
    return `https://firebasestorage.googleapis.com/v0/b/${BUCKET}/o/${objectName}?alt=media`;
}

async function assertUnauthenticatedReadIsDenied() {
    const response = await fetch(mediaUrl());
    if (response.status !== 401 && response.status !== 403) {
        throw new Error(`Unauthenticated read unexpectedly returned HTTP ${response.status}.`);
    }
    return response.status;
}

async function exchangeCustomToken(customToken, apiKey) {
    const response = await fetch(
        `https://identitytoolkit.googleapis.com/v1/accounts:signInWithCustomToken?key=${encodeURIComponent(apiKey)}`,
        {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ token: customToken, returnSecureToken: true })
        }
    );
    const data = await response.json();
    if (!response.ok || !data.idToken) {
        throw new Error(`Custom-token exchange failed with HTTP ${response.status}.`);
    }
    return data.idToken;
}

async function assertAuthenticatedReadSucceeds(idToken) {
    const response = await fetch(mediaUrl(), {
        headers: { authorization: `Bearer ${idToken}` }
    });
    if (!response.ok) {
        throw new Error(`Authenticated read failed with HTTP ${response.status}.`);
    }
    const manifest = await response.json();
    if (!Array.isArray(manifest.layers) || manifest.layers.length === 0) {
        throw new Error('CAD manifest does not contain layers.');
    }
    return manifest.layers.length;
}

async function main() {
    const unauthenticatedStatus = await assertUnauthenticatedReadIsDenied();
    const app = initializeApp({
        credential: cert(loadServiceAccount()),
        projectId: PROJECT_ID
    }, 'cad-storage-verifier');
    try {
        const customToken = await getAuth(app).createCustomToken(
            `humetro:${VERIFY_EMAIL}`,
            { humetro: true, humetroEmail: VERIFY_EMAIL }
        );
        const idToken = await exchangeCustomToken(customToken, loadWebApiKey());
        const layerCount = await assertAuthenticatedReadSucceeds(idToken);
        console.log(JSON.stringify({
            success: true,
            unauthenticatedStatus,
            authenticatedStatus: 200,
            layerCount
        }, null, 2));
    } finally {
        await deleteApp(app);
    }
}

main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
});
