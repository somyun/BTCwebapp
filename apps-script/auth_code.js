var MAP_AUTH_DOMAIN = '@humetro.busan.kr';
var MAP_AUTH_CODE_TTL_SECONDS = 600;
var MAP_AUTH_RESEND_SECONDS = 60;
var MAP_AUTH_MAX_ATTEMPTS = 5;
var MAP_AUTH_GLOBAL_REQUESTS_PER_MINUTE = 20;
var MAP_AUTH_UID_PREFIX = 'humetro:';

// Apps Script 편집기에서 최초 1회 실행하여 메일 발송 권한을 승인합니다.
function authorizeMapAuthMail() {
    return { remainingDailyQuota: MailApp.getRemainingDailyQuota() };
}

function normalizeMapAuthEmail_(value) {
    return String(value || '').trim().toLowerCase();
}

function isAllowedMapAuthEmail_(email) {
    return /^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@humetro\.busan\.kr$/i.test(email);
}

function mapAuthDigest_(value) {
    var bytes = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        String(value),
        Utilities.Charset.UTF_8
    );
    return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function mapAuthCacheKey_(prefix, email) {
    return 'map_auth_' + prefix + '_' + mapAuthDigest_(email).substring(0, 32);
}

function mapAuthPepper_() {
    var properties = PropertiesService.getScriptProperties();
    var pepper = properties.getProperty('MAP_AUTH_CODE_PEPPER');
    if (!pepper) {
        pepper = Utilities.getUuid() + Utilities.getUuid();
        properties.setProperty('MAP_AUTH_CODE_PEPPER', pepper);
    }
    return pepper;
}

function mapAuthCodeHash_(email, code, nonce) {
    return mapAuthDigest_([email, code, nonce, mapAuthPepper_()].join('|'));
}

function generateMapAuthCode_() {
    var seed = Utilities.getUuid() + '|' + new Date().getTime() + '|' + Math.random();
    var bytes = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        seed,
        Utilities.Charset.UTF_8
    );
    var number = ((bytes[0] & 255) * 65536) + ((bytes[1] & 255) * 256) + (bytes[2] & 255);
    return String(number % 1000000).padStart(6, '0');
}

function allowMapAuthRequest_(email) {
    var cache = CacheService.getScriptCache();
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
        var emailKey = mapAuthCacheKey_('send', email);
        if (cache.get(emailKey)) return false;

        var minute = Math.floor(new Date().getTime() / 60000);
        var globalKey = 'map_auth_global_' + minute;
        var current = Number(cache.get(globalKey) || 0);
        if (current >= MAP_AUTH_GLOBAL_REQUESTS_PER_MINUTE) return false;

        cache.put(globalKey, String(current + 1), 120);
        cache.put(emailKey, '1', MAP_AUTH_RESEND_SECONDS);
        return true;
    } finally {
        lock.releaseLock();
    }
}

function requestMapAuthCode(emailValue) {
    var email = normalizeMapAuthEmail_(emailValue);
    if (!isAllowedMapAuthEmail_(email)) {
        return { success: false, code: 'INVALID_DOMAIN' };
    }
    if (!allowMapAuthRequest_(email)) {
        return { success: false, code: 'RATE_LIMITED' };
    }
    if (MailApp.getRemainingDailyQuota() < 1) {
        return { success: false, code: 'MAIL_QUOTA_EXCEEDED' };
    }

    var code = generateMapAuthCode_();
    var nonce = Utilities.getUuid();
    var now = new Date().getTime();
    var record = {
        hash: mapAuthCodeHash_(email, code, nonce),
        nonce: nonce,
        expiresAt: now + (MAP_AUTH_CODE_TTL_SECONDS * 1000),
        attempts: 0
    };
    CacheService.getScriptCache().put(
        mapAuthCacheKey_('code', email),
        JSON.stringify(record),
        MAP_AUTH_CODE_TTL_SECONDS
    );

    try {
        MailApp.sendEmail({
            to: email,
            subject: '[BTCwebapp] 도면 지도 인증코드',
            name: 'BTCwebapp 도면 지도',
            body: '도면 지도 인증코드는 ' + code + ' 입니다.\n\n이 코드는 10분 동안 유효합니다. 본인이 요청하지 않았다면 이 메일을 무시하세요.',
            htmlBody: '<div style="font-family:Arial,sans-serif;line-height:1.6">' +
                '<h2 style="margin:0 0 12px">도면 지도 인증코드</h2>' +
                '<p>아래 6자리 숫자를 BTCwebapp에 입력하세요.</p>' +
                '<p style="font-size:30px;font-weight:700;letter-spacing:8px;margin:20px 0">' + code + '</p>' +
                '<p>이 코드는 <strong>10분</strong> 동안 유효합니다.</p>' +
                '<p style="color:#64748b;font-size:13px">본인이 요청하지 않았다면 이 메일을 무시하세요.</p>' +
                '</div>'
        });
        return { success: true, expiresIn: MAP_AUTH_CODE_TTL_SECONDS };
    } catch (error) {
        CacheService.getScriptCache().remove(mapAuthCacheKey_('code', email));
        console.error('Map auth email failed:', error);
        return { success: false, code: 'MAIL_SEND_FAILED' };
    }
}

function mapAuthSafeEqual_(left, right) {
    left = String(left || '');
    right = String(right || '');
    if (left.length !== right.length) return false;
    var difference = 0;
    for (var i = 0; i < left.length; i++) {
        difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
    }
    return difference === 0;
}

function verifyMapAuthCode(emailValue, codeValue) {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
        return verifyMapAuthCodeLocked_(emailValue, codeValue);
    } finally {
        lock.releaseLock();
    }
}

function verifyMapAuthCodeLocked_(emailValue, codeValue) {
    var email = normalizeMapAuthEmail_(emailValue);
    var code = String(codeValue || '').replace(/\D/g, '').substring(0, 6);
    if (!isAllowedMapAuthEmail_(email) || !/^\d{6}$/.test(code)) {
        return { success: false, code: 'INVALID_CODE' };
    }

    var cache = CacheService.getScriptCache();
    var cacheKey = mapAuthCacheKey_('code', email);
    var serialized = cache.get(cacheKey);
    if (!serialized) return { success: false, code: 'CODE_EXPIRED' };

    var record = JSON.parse(serialized);
    var remainingSeconds = Math.floor((record.expiresAt - new Date().getTime()) / 1000);
    if (remainingSeconds < 1) {
        cache.remove(cacheKey);
        return { success: false, code: 'CODE_EXPIRED' };
    }
    if (record.attempts >= MAP_AUTH_MAX_ATTEMPTS) {
        cache.remove(cacheKey);
        return { success: false, code: 'TOO_MANY_ATTEMPTS' };
    }

    var submittedHash = mapAuthCodeHash_(email, code, record.nonce);
    if (!mapAuthSafeEqual_(record.hash, submittedHash)) {
        record.attempts += 1;
        if (record.attempts >= MAP_AUTH_MAX_ATTEMPTS) {
            cache.remove(cacheKey);
            return { success: false, code: 'TOO_MANY_ATTEMPTS' };
        }
        cache.put(cacheKey, JSON.stringify(record), remainingSeconds);
        return { success: false, code: 'INVALID_CODE' };
    }

    cache.remove(cacheKey);
    try {
        return { success: true, token: createMapAuthCustomToken_(email) };
    } catch (error) {
        console.error('Map auth custom token failed:', error);
        return { success: false, code: 'TOKEN_CREATE_FAILED' };
    }
}

function mapAuthBase64Url_(value) {
    var encoded = typeof value === 'string'
        ? Utilities.base64EncodeWebSafe(value, Utilities.Charset.UTF_8)
        : Utilities.base64EncodeWebSafe(value);
    return encoded.replace(/=+$/g, '');
}

function createMapAuthCustomToken_(email) {
    if (!SERVICE_ACCOUNT_KEY || !SERVICE_ACCOUNT_KEY.private_key || !SERVICE_ACCOUNT_KEY.client_email) {
        throw new Error('SERVICE_ACCOUNT_KEY_MISSING');
    }
    var now = Math.floor(new Date().getTime() / 1000);
    var header = { alg: 'RS256', typ: 'JWT' };
    var payload = {
        iss: SERVICE_ACCOUNT_KEY.client_email,
        sub: SERVICE_ACCOUNT_KEY.client_email,
        aud: 'https://identitytoolkit.googleapis.com/google.identity.identitytoolkit.v1.IdentityToolkit',
        iat: now,
        exp: now + 3600,
        uid: MAP_AUTH_UID_PREFIX + email,
        claims: {
            humetro: true,
            humetroEmail: email
        }
    };
    var unsignedToken = mapAuthBase64Url_(JSON.stringify(header)) + '.' + mapAuthBase64Url_(JSON.stringify(payload));
    var signature = Utilities.computeRsaSha256Signature(unsignedToken, SERVICE_ACCOUNT_KEY.private_key);
    return unsignedToken + '.' + mapAuthBase64Url_(signature);
}
