'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.join(__dirname, '..');
const read = (...segments) => fs.readFileSync(path.join(root, ...segments), 'utf8');

test('notification pages and worker use only production Firebase endpoints and scope', () => {
    const sources = [
        read('notification-settings.js'),
        read('notifications.js'),
        read('firebase-messaging-sw.js'),
        read('notification-settings.html'),
        read('notifications.html')
    ].join('\n');
    assert.match(sources, /btcwebapp-551bd/);
    assert.match(sources, /\/BTCwebapp\//);
    assert.doesNotMatch(sources, /btcwebapp-test|bwa[_-]test|BWA_TEST|\/bwa_test\//);
});

test('production keeps notification settings unprefixed and history in a separate database', () => {
    assert.match(read('notification-settings.js'), /const STORAGE_PREFIX = '';/);
    assert.match(read('notification-store.js'), /btcwebapp-production-notifications-v1/);
});

test('production notification dispatch is fail-closed behind an admin gate', () => {
    const source = read('firebase-production', 'functions', 'index.js');
    assert.match(source, /doc\("notificationDispatch"\)/);
    assert.match(source, /gate\.get\("enabled"\) !== true/);
    assert.match(source, /exports\.setNotificationDispatchGate = manualAdmin/);
});

test('test-only environment values do not leak into adopted production code', () => {
    const adoptedFiles = [
        'firebase-messaging-sw.js',
        'notification-settings.js',
        'notifications.js',
        path.join('firebase-production', 'functions', 'index.js'),
        path.join('firebase-production', 'functions', 'lib', 'publisher.js'),
        path.join('firebase-production', 'functions', 'lib', 'humetro-client.js'),
        path.join('firebase-production', 'functions', 'lib', 'notification-service.js'),
        path.join('firebase-production', 'functions', 'lib', 'sheets-sync.js')
    ];
    const sources = adoptedFiles.map((file) => read(file)).join('\n');
    assert.doesNotMatch(sources, /btcwebapp-test|BWA_TEST_PUBLISHER_TOKEN|bwa[_-]test|\/bwa_test\//);
    assert.match(sources, /19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s/);
    assert.match(sources, /BWA_PUBLISHER_TOKEN/);
});

test('the existing public CAD loader remains in place without private Storage adoption', () => {
    const index = read('index.html');
    assert.match(index, /<script src="map\.js\?v=hopo-cad-18"><\/script>/);
    assert.doesNotMatch(index, /cad-storage\.js|firebasestorage\.googleapis\.com|storage\.googleapis\.com/);
});

test('CSP-safe controls use script listeners instead of inline handlers', () => {
    const index = read('index.html');
    const script = read('script.js');
    assert.doesNotMatch(index, /onchange=|onclick=/);
    assert.match(script, /getElementById\('uploadBtn'\)\?\.addEventListener\('click', handleUpload\)/);
    assert.match(script, /getElementById\('formSelect'\)\?\.addEventListener\('change', loadSelectedForm\)/);
});
