// kakao.js
// -----------------------------------------------------------------------
// 카카오톡 알림 API 모듈
// -----------------------------------------------------------------------

/**
 * [설정] 카카오 REST API 키를 가져옵니다.
 * GAS 프로젝트 설정 > 스크립트 속성에서 'KAKAO_API_KEY'를 설정해야 합니다.
 */
function getKakaoApiKey() {
    return PropertiesService.getScriptProperties().getProperty('KAKAO_API_KEY');
}

/**
 * [최초 설정용] 인증 코드를 토큰으로 교환합니다.
 * 이 함수는 최초 1회만 수동으로 실행하면 됩니다.
 * @param {string} authCode - 카카오 로그인 후 받은 Authorization Code
 */
function setupKakaoToken(authCode) {
    const apiKey = getKakaoApiKey();
    if (!apiKey) {
        throw new Error("스크립트 속성에 'KAKAO_API_KEY'가 설정되지 않았습니다.");
    }

    // 인증 코드가 있는지 확인
    if (!authCode) {
        throw new Error("인증 코드(authCode)가 필요합니다.");
    }

    const payload = {
        grant_type: "authorization_code",
        client_id: apiKey,
        redirect_uri: "https://global.sites.kakao.com/web/login/oauth",
        code: authCode
    };

    const options = {
        method: "post",
        payload: payload,
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", options);
    const result = JSON.parse(response.getContentText());

    if (result.access_token) {
        // 토큰 저장
        const props = PropertiesService.getScriptProperties();
        props.setProperty('KAKAO_ACCESS_TOKEN', result.access_token);
        if (result.refresh_token) {
            props.setProperty('KAKAO_REFRESH_TOKEN', result.refresh_token);
        }
        console.log("카카오 토큰 설정 완료! 이제 자동 갱신됩니다.");
        console.log("Access Token:", result.access_token.substring(0, 10) + "...");
    } else {
        console.error("토큰 발급 실패:", result);
        throw new Error("토큰 발급 실패: " + JSON.stringify(result));
    }
}

/**
 * [내부용] 유효한 Access Token을 가져옵니다 (만료 시 자동 갱신).
 */
function getValidKakaoToken() {
    const props = PropertiesService.getScriptProperties();
    let accessToken = props.getProperty('KAKAO_ACCESS_TOKEN');
    const refreshToken = props.getProperty('KAKAO_REFRESH_TOKEN');
    const apiKey = getKakaoApiKey();

    if (!refreshToken) {
        throw new Error("저장된 Refresh Token이 없습니다. setupKakaoToken을 먼저 실행하세요.");
    }

    // 토큰 유효성 검사 (간단히 사용자 정보 요청으로 테스트)
    // 매번 호출하면 느리므로, 실제 API 호출 시 401 에러가 나면 갱신하는 방식이 효율적이나,
    // 여기서는 안정성을 위해 갱신 로직을 별도로 분리합니다.
    // 우선 현재 토큰을 그냥 반환하고, API 호출부에서 401 처리를 하는 것이 정석입니다.
    // 하지만 GAS 구조상 호출부마다 try-catch하기 번거로우므로,
    // 여기서는 '갱신이 필요한 경우'를 판단해서 갱신합니다.

    // (단순화를 위해) 만료 시간을 저장하지 않았다면, 일단 기존 토큰을 반환하고
    // 실제 API 호출 함수에서 401 발생 시 이 함수를 '강제 갱신 모드'로 부르는게 좋지만,
    // 여기서는 refreshAccessToken 함수를 따로 만들어서 필요할 때 부르겠습니다.

    return accessToken;
}

/**
 * Access Token을 Refresh Token으로 갱신합니다.
 */
function refreshKakaoToken() {
    const props = PropertiesService.getScriptProperties();
    const refreshToken = props.getProperty('KAKAO_REFRESH_TOKEN');
    const apiKey = getKakaoApiKey();

    if (!refreshToken || !apiKey) {
        throw new Error("토큰 갱신 실패: Refresh Token 또는 API Key가 없습니다.");
    }

    const payload = {
        grant_type: "refresh_token",
        client_id: apiKey,
        refresh_token: refreshToken
    };

    const options = {
        method: "post",
        payload: payload,
        muteHttpExceptions: true
    };

    console.log("카카오 토큰 갱신 시도...");
    const response = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", options);
    const result = JSON.parse(response.getContentText());

    if (result.access_token) {
        props.setProperty('KAKAO_ACCESS_TOKEN', result.access_token);
        // Refresh Token도 갱신될 수 있음 (유효기간 연장 등)
        if (result.refresh_token) {
            props.setProperty('KAKAO_REFRESH_TOKEN', result.refresh_token);
        }
        console.log("카카오 토큰 갱신 성공!");
        return result.access_token;
    } else {
        console.error("토큰 갱신 실패:", result);
        throw new Error("토큰 갱신 실패: " + JSON.stringify(result));
    }
}

/**
 * 카카오톡 '나에게 보내기' 메시지를 전송합니다.
 * @param {string} text - 보낼 메시지 내용
 */
function sendKakaoMemo(text) {
    try {
        _sendKakaoMemoInternal(text);
    } catch (e) {
        // 401 Unauthorized 에러인 경우 토큰 갱신 후 재시도
        if (e.message.includes("401") || e.message.includes("unauthorized")) {
            console.warn("토큰 만료 감지됨. 갱신 후 재시도합니다.");
            refreshKakaoToken();
            _sendKakaoMemoInternal(text); // 재시도에서 실패하면 진짜 에러
        } else {
            throw e;
        }
    }
}

// 내부 실제 전송 로직
function _sendKakaoMemoInternal(text) {
    const accessToken = getValidKakaoToken();
    if (!accessToken) throw new Error("유효한 Access Token이 없습니다.");

    const url = "https://kapi.kakao.com/v2/api/talk/memo/default/send";

    // 템플릿 객체 (Text 타입)
    const templateObject = {
        object_type: "text",
        text: text,
        link: {
            web_url: "https://www.humetro.busan.kr",
            mobile_web_url: "https://www.humetro.busan.kr"
        },
        button_title: "바로가기"
    };

    const options = {
        method: "post",
        headers: {
            "Authorization": "Bearer " + accessToken
        },
        payload: {
            template_object: JSON.stringify(templateObject)
        },
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
        console.log("카카오톡 전송 성공 (나에게 보내기)");
    } else {
        // 401이면 상위 catch 블록에서 잡아서 갱신 로직 수행하도록 에러 throw
        if (responseCode === 401) {
            throw new Error("Kakao API 401 Unauthorized");
        }
        console.error(`카카오톡 전송 실패: ${responseCode} - ${responseBody}`);
        throw new Error(`카카오톡 전송 실패: ${responseBody}`);
    }
}

/**
 * [확장용] 카카오톡 '친구에게 보내기' 메시지를 전송합니다.
 * (친구 목록 조회 권한 및 친구의 동의 필요)
 * @param {string} text - 보낼 메시지 내용
 * @param {Array<string>} receiverUuids - 받을 친구들의 UUID 배열 (최대 5명)
 */
function sendKakaoFriendsMessage(text, receiverUuids) {
    if (!receiverUuids || receiverUuids.length === 0) {
        console.warn("수신자 UUID가 없어 친구에게 보내기를 건너뜁니다.");
        return;
    }

    try {
        _sendKakaoFriendsInternal(text, receiverUuids);
    } catch (e) {
        if (e.message.includes("401")) {
            console.warn("토큰 만료 감지됨 (친구 전송). 갱신 후 재시도합니다.");
            refreshKakaoToken();
            _sendKakaoFriendsInternal(text, receiverUuids);
        } else {
            throw e;
        }
    }
}

// 내부 친구 전송 로직
function _sendKakaoFriendsInternal(text, receiverUuids) {
    const accessToken = getValidKakaoToken();
    const url = "https://kapi.kakao.com/v1/api/talk/friends/message/default/send";

    const templateObject = {
        object_type: "text",
        text: text,
        link: {
            web_url: "https://www.humetro.busan.kr",
            mobile_web_url: "https://www.humetro.busan.kr"
        },
        button_title: "확인"
    };

    const options = {
        method: "post",
        headers: {
            "Authorization": "Bearer " + accessToken
        },
        payload: {
            receiver_uuids: JSON.stringify(receiverUuids),
            template_object: JSON.stringify(templateObject)
        },
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
        const result = JSON.parse(response.getContentText());
        console.log(`친구에게 전송 완료. 성공: ${result.successful_receiver_uuids?.length}, 실패: ${result.failure_info?.length}`);
    } else {
        if (responseCode === 401) throw new Error("Kakao API 401 Unauthorized");
        console.error(`친구에게 전송 실패: ${response.getContentText()}`);
        throw new Error(`친구에게 전송 실패: ${response.getContentText()}`);
    }
}

/**
 * [관리자용] 친구 목록을 조회하여 로그에 출력합니다. (UUID 확인용)
 * 이 함수를 실행하고 로그에서 친구의 UUID를 찾아보세요.
 */
function printFriendList() {
    try {
        const accessToken = getValidKakaoToken();
        const url = "https://kapi.kakao.com/v1/api/talk/friends";

        const options = {
            method: "get",
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();

        if (responseCode === 200) {
            const result = JSON.parse(response.getContentText());
            console.log(`[친구 목록 조회 결과] 총 ${result.total_count}명`);

            if (result.elements && result.elements.length > 0) {
                result.elements.forEach((friend, index) => {
                    console.log(`${index + 1}. 이름: ${friend.profile_nickname}, UUID: ${friend.uuid}`);
                });
                console.log("--------------------------------------------------");
                console.log("위 UUID를 복사해서 sendKakaoFriendsMessage 함수에 사용하세요.");
            } else {
                console.log("친구 목록이 비어있습니다. (친구가 이 앱에 동의하지 않았거나 권한이 없습니다.)");
            }
        } else {
            console.error(`친구 목록 조회 실패: ${response.getContentText()}`);
        }
    } catch (e) {
        console.error("오류 발생:", e);
    }
}


/**
 * [사용자 실행용] 최초 설정을 위해 이 함수를 사용하세요.
 * 1. 아래 변수에 방금 복사한 '인증 코드'를 넣으세요.
 * 2. 상단 실행 함수 목록에서 `_runManualSetup`을 선택하고 '실행'을 누르세요.
 */
function _runManualSetup() {
    const myAuthCode = "61PQFUSgxq0pU_1UkoFjLg1JT_WCVwliWpD8pgJ7iixqDmQbAKIs1QAAAAQKDQxeAAABnBhxhlV-jFVpBnvzXw";

    setupKakaoToken(myAuthCode);
}

