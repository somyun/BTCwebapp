'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const authSource = fs.readFileSync(path.join(root, 'auth.js'), 'utf8');
const styles = fs.readFileSync(path.join(root, 'style.css'), 'utf8');

test('map authentication asks only for the Humetro email id', () => {
    assert.match(html, /id="mapAuthEmail"[^>]*autocomplete="username"/);
    assert.match(html, /class="map-auth-email-domain"[^>]*>@humetro\.busan\.kr</);
    assert.doesNotMatch(html, /placeholder="name@humetro\.busan\.kr"/);
    assert.match(styles, /\.map-auth-email-entry\s*{[\s\S]*?display: flex/);
});

test('the fixed Humetro domain is appended before auth API calls', () => {
    assert.match(authSource, /function addressFromEmailId\(value\)/);
    assert.match(authSource, /return `\$\{emailId\}\$\{ALLOWED_DOMAIN\}`/);
    assert.match(authSource, /callAuthApi\('requestMapAuthCode', \{ email: value \}\)/);
    assert.match(authSource, /email: pendingEmail/);
});
