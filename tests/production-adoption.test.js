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

test('production collects Happy Hugether posts directly without the GAS bridge', () => {
    const index = read('firebase-production', 'functions', 'index.js');
    const client = read('firebase-production', 'functions', 'lib', 'humetro-client.js');
    const combined = `${index}\n${client}`;
    assert.match(combined, /HUMETRO_ID/);
    assert.match(combined, /HUMETRO_PW/);
    assert.match(client, /https:\/\/www\.humetro\.busan\.kr/);
    assert.doesNotMatch(combined, /HUMETRO_BRIDGE_TOKEN|bridgeTokenProvider|script\.google\.com/);
});

test('legacy notification devices can be imported and silently claimed by the same browser token', () => {
    const migration = read('firebase-production', 'functions', 'lib', 'legacy-notification-migration.js');
    const service = read('firebase-production', 'functions', 'lib', 'notification-service.js');
    const settings = read('notification-settings.js');
    const app = read('script.js');
    assert.match(migration, /FCM_Tokens/);
    assert.match(migration, /legacy:\s*true/);
    assert.match(service, /migratedLegacy/);
    assert.match(service, /notificationTokenOwners/);
    assert.match(settings, /migratedLegacy/);
    assert.match(app, /bootstrapFirebaseNotificationMigration/);
});

test('production cache versions force the notification migration code and worker to refresh', () => {
    assert.match(read('index.html'), /production-read\.js\?v=daily-measurement-cache-1/);
    assert.match(read('index.html'), /script\.js\?v=daily-measurement-cache-1/);
    assert.match(read('notification-settings.html'), /notification-settings\.js\?v=production-notifications-2/);
    assert.match(read('script.js'), /firebase-messaging-sw\.js\?v=production-notifications-2/);
});

test('legacy notification UI is absent and removed again after browser history restoration', () => {
    const index = read('index.html');
    const script = read('script.js');
    const styles = read('style.css');
    assert.doesNotMatch(index, /id="(?:notificationToggle|notificationToggleMain|mainToggleContainer|keywordModalOverlay|keywordModal)"/);
    assert.match(script, /const LEGACY_NOTIFICATION_SELECTORS = \[/);
    assert.match(script, /window\.addEventListener\('pageshow',/);
    assert.match(script, /const APP_DOCUMENT_VERSION = 'production-adoption-4';/);
    assert.match(script, /window\.history\.replaceState\(window\.history\.state, '', url\)/);
    assert.match(read('notifications.html'), /href="\.\/\?app=production-adoption-4"/);
    assert.match(read('notification-settings.html'), /href="\.\/\?app=production-adoption-4"/);
    assert.doesNotMatch(styles, /\.keyword-modal-(?:overlay|content)|\.switch-container/);
});

test('notification documents and assets are prefetched for faster navigation', () => {
    const index = read('index.html');
    assert.match(index, /rel="prefetch" href="\.\/notifications\.html\?v=production-notifications-3" as="document"/);
    assert.match(index, /rel="prefetch" href="\.\/notification-settings\.html\?v=production-notifications-3" as="document"/);
    assert.match(index, /href="\.\/notifications\.html\?v=production-notifications-3"/);
    assert.match(index, /href="\.\/notification-settings\.html\?v=production-notifications-3"/);
});

test('dynamic form header controls use CSP-safe event listeners', () => {
    const script = read('script.js');
    assert.doesNotMatch(script, /on(?:click|change)=/);
    assert.match(script, /getElementById\('toggleSortBtn'\)\?\.addEventListener\('click', toggleSortMode\)/);
    assert.match(script, /getElementById\('spacingSelect'\)\?\.addEventListener\('change'/);
    assert.match(script, /getElementById\('resetOrderBtn'\)\?\.addEventListener\('click', handleResetOrder\)/);
});

test('the header hamburger overrides global full-width button styles', () => {
    const index = read('index.html');
    const style = read('style.css');
    assert.match(index, /style\.css\?v=production-adoption-2/);
    assert.match(style, /\.hamburger-menu\s*\{[\s\S]*?width:\s*38px;[\s\S]*?height:\s*38px;[\s\S]*?margin:\s*0;[\s\S]*?background:\s*transparent;/);
    assert.match(style, /\.hamburger-menu:hover\s*\{[\s\S]*?background:\s*#f0f0f0;/);
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

test('today measurement caches are publicly readable but client writes remain blocked', () => {
    const rules = read('firebase-production', 'firestore.rules');
    assert.match(rules, /match \/dailyMeasurementCaches\/\{cacheId\}/);
    assert.match(rules, /match \/dailyMeasurementCaches[\s\S]*allow read: if true;[\s\S]*allow write: if false;/);
});
