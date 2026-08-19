// fcm_service.js

/**
 * FCM HTTP v1 API를 통해 푸시 알림을 전송합니다.
 * @param {string} token - 수신자 기기 토큰
 * @param {string} title - 알림 제목
 * @param {string} body - 알림 내용
 * @param {string} icon - (선택) 아이콘 URL
 * @param {string} link - (선택) 클릭 시 이동할 링크
 */
function sendFCMNotification(token, title, body, icon, link) {
    console.log(`[FCM] 발송 시작: ${title} -> ${token.substring(0, 15)}...`);

    const accessToken = getAccessToken();
    if (!accessToken) {
        console.error('[FCM] 액세스 토큰 발급 실패');
        return { success: false, message: 'Access Token Error' };
    }

    // 프로젝트 ID는 키 파일에서 가져옵니다.
    const projectId = SERVICE_ACCOUNT_KEY.project_id;
    const url = `https://fcm.googleapis.com/v1/projects/${projectId}/messages:send`;

    const payload = {
        "message": {
            "token": token,
            "data": {
                "title": title,
                "body": body,
                "url": link || "https://somyun.github.io/btc_webapp/",
                "icon": icon || '/icon.png'
            }
        }
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'Authorization': `Bearer ${accessToken}`
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode === 200) {
            console.log('[FCM] 발송 성공:', responseBody);
            return { success: true, message: 'Sent' };
        } else {
            console.error(`[FCM] 발송 실패 (${responseCode}):`, responseBody);
            return { success: false, message: responseBody };
        }
    } catch (e) {
        console.error('[FCM] 발송 중 예외 발생:', e);
        return { success: false, message: e.message };
    }
}

/**
 * 서비스 계정 키를 사용하여 OAuth 2.0 액세스 토큰을 발급받습니다.
 */
function getAccessToken() {
    if (!SERVICE_ACCOUNT_KEY || !SERVICE_ACCOUNT_KEY.private_key) {
        console.error('[FCM] SERVICE_ACCOUNT_KEY가 설정되지 않았습니다. firebase_key.js를 확인하세요.');
        return null;
    }

    // JWT 생성
    const privateKey = SERVICE_ACCOUNT_KEY.private_key;
    const clientEmail = SERVICE_ACCOUNT_KEY.client_email;
    const now = Math.floor(new Date().getTime() / 1000);
    const oneHour = 3600;

    const header = {
        alg: "RS256",
        typ: "JWT"
    };

    const claimSet = {
        iss: clientEmail,
        scope: "https://www.googleapis.com/auth/firebase.messaging",
        aud: "https://oauth2.googleapis.com/token",
        exp: now + oneHour,
        iat: now
    };

    const toSign = Utilities.base64EncodeWebSafe(JSON.stringify(header)) + "." + Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));

    // 서명 생성
    const signatureBytes = Utilities.computeRsaSha256Signature(toSign, privateKey);
    const signature = Utilities.base64EncodeWebSafe(signatureBytes);

    const jwt = toSign + "." + signature;

    // 액세스 토큰 요청
    const options = {
        method: 'post',
        payload: {
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion: jwt
        },
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", options);
        const json = JSON.parse(response.getContentText());

        if (json.access_token) {
            return json.access_token;
        } else {
            console.error('[FCM] 토큰 응답 오류:', json);
            return null;
        }
    } catch (e) {
        console.error('[FCM] 토큰 요청 실패:', e);
        return null;
    }
}
