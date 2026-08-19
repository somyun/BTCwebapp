//hugether.js
// -----------------------------------------------------------------------
//부산교통공사 경조사 게시판 크롤러 로직
// -----------------------------------------------------------------------

/**
 * 게시판 데이터를 가져오는 메인 핸들러
 */
function handleGetBoardData() {
    try {
        // [디버깅] 현재 GAS 실행 IP 확인
        try {
            const ipResponse = UrlFetchApp.fetch("http://checkip.dyndns.org/");
            const ip = ipResponse.getContentText().match(/Current IP Address: ([0-9\.]+)/)[1];
            console.log(`현재 GAS 실행 IP: ${ip}`);
        } catch (e) {
            console.warn("IP 확인 실패");
        }

        // 1. 로그인 (쿠키 획득)
        const loginCookie = loginToHumetro();
        if (!loginCookie) {
            throw new Error("로그인 실패: 세션 쿠키를 획득하지 못했습니다. 아이디/비번을 확인하거나 로그를 확인하세요.");
        }

        // 2. 게시판 리스트 조회
        const boardData = fetchBoardList(loginCookie);

        // 3. [알림] 카카오톡 전송 (첫 번째 게시글)
        if (boardData && boardData.length > 0) {
            try {
                const firstTitle = boardData[0].title;
                console.log(`알림 보낼 제목: ${firstTitle}`);
                // kakao.js의 함수 호출
                sendKakaoMemo(`[경조사 알림] ${firstTitle}`);
            } catch (e) {
                console.warn(`카카오톡 전송 중 오류 (크롤링은 성공): ${e.message}`);
            }
        }

        return {
            success: true,
            data: boardData,
            message: `성공적으로 ${boardData.length}개의 게시글을 가져왔습니다.`
        };

    } catch (error) {
        console.error("게시판 크롤링 중 오류:", error);
        return {
            success: false,
            message: error.toString()
        };
    }
}

/**
 * 부산교통공사 로그인 수행 및 세션 쿠키 반환
 */
function loginToHumetro() {
    const loginUrl = "https://www.humetro.busan.kr/homepage/default/member/page/loginProcEvent.do";

    // [보안] 스크립트 속성에서 아이디/비번 가져오기
    const scriptProperties = PropertiesService.getScriptProperties();
    const userId = scriptProperties.getProperty('HUMETRO_ID');
    const userPw = scriptProperties.getProperty('HUMETRO_PW');

    if (!userId || !userPw) {
        throw new Error("아이디/비밀번호가 설정되지 않았습니다. GAS 프로젝트 설정 > 스크립트 속성에서 'HUMETRO_ID'와 'HUMETRO_PW'를 추가해주세요.");
    }

    // x-www-form-urlencoded 형식의 문자열 body 생성
    const rawPayload = `RETURNURL=&userID=${encodeURIComponent(userId)}&password=${encodeURIComponent(userPw)}`;


    // 사용자가 제공한 fetch 헤더를 최대한 유사하게 모방
    const options = {
        "method": "post",
        "payload": rawPayload,
        "headers": {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "ko,en;q=0.9,en-US;q=0.8",
            "Cache-Control": "max-age=0",
            "Connection": "keep-alive",
            "DNT": "1",
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer": "https://www.humetro.busan.kr/event.do",
            "Origin": "https://www.humetro.busan.kr",
            "Sec-Ch-Ua": "\"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"144\", \"Microsoft Edge\";v=\"144\"",
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": "\"Windows\"",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36 Edg/144.0.0.0"
        },
        "followRedirects": false,
        "muteHttpExceptions": true
    };

    console.log("로그인 시도 중...");
    const response = UrlFetchApp.fetch(loginUrl, options);
    const responseCode = response.getResponseCode();
    const headers = response.getAllHeaders();

    //[수정] EUC-KR 디코딩 시도 (한글 깨짐 해결)
    let content = "";
    try {
        const blob = response.getBlob();
        // 부산교통공사 등 공공기관은 EUC-KR을 자주 사용함
        content = blob.getDataAsString("EUC-KR");
    } catch (e) {
        content = response.getContentText(); // 실패 시 기본 (UTF-8)
    }

    console.log(`로그인 응답 코드: ${responseCode}`);
    console.log(`응답 헤더: ${JSON.stringify(headers)}`);
    // HTML 태그 제거하고 텍스트만 보려 노력
    const cleanContent = content.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "").replace(/<[^>]+>/g, "\n").replace(/\s+/g, " ").trim();
    console.log(`응답 본문(텍스트 변환): ${cleanContent.substring(0, 500)}`);

    // Set-Cookie 확인 logic 강화
    let cookies = [];
    if (headers['Set-Cookie']) {
        const setCookie = headers['Set-Cookie'];
        cookies = Array.isArray(setCookie) ? setCookie : [setCookie];
    } else if (headers['set-cookie']) { // 소문자 대응
        const setCookie = headers['set-cookie'];
        cookies = Array.isArray(setCookie) ? setCookie : [setCookie];
    }

    if (cookies.length > 0) {
        // JSESSIONID만 추출하거나 전체를 합침
        const cookieHeader = cookies.map(c => c.split(';')[0]).join('; ');
        console.log("획득한 쿠키:", cookieHeader);
        return cookieHeader;
    }

    console.warn("경고: 응답에 Set-Cookie 헤더가 없습니다.");
    return null;
}

/**
 * 게시판 리스트 HTML을 가져와서 파싱
 * - 공지사항(번호 없음) 제외
 * - 게시글 번호, 제목, 링크 추출
 */
function fetchBoardList(cookieHeader) {
    const targetUrl = "https://www.humetro.busan.kr/homepage/default/board/listEvent.do?conf_no=151&menu_no=1001060402";
    const payload = { "RETURNURL": "null" };

    // 헤더를 수정하지 말것. 헤더 수정 시 게시글이 제대로 가져오지 않음
    const options = {
        "method": "post",
        "payload": payload,
        "headers": {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "ko,en;q=0.9,en-US;q=0.8",
            "Cache-Control": "max-age=0",
            "Content-Type": "application/x-www-form-urlencoded",
            "Cookie": cookieHeader,
            "Origin": "https://www.humetro.busan.kr",
            "Referer": "https://www.humetro.busan.kr/homepage/default/member/page/loginProcEvent.do",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Upgrade-Insecure-Requests": "1"
        },
        "followRedirects": true,
        "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(targetUrl, options);
    let content = response.getContentText();

    //tbody 추출
    const tbodyMatch = content.match(/<tbody>([\s\S]*?)<\/tbody>/i);
    if (!tbodyMatch) {
        console.error("게시판 Parsing Error: tbody not found");
        return [];
    }

    const rows = tbodyMatch[1].split('</tr>');
    const posts = [];

    // 각 행(tr) 파싱
    for (const row of rows) {
        if (!row.trim()) continue;

        // 1. 번호 추출 (th 태그 안)
        // 정상 글: <th> ... 6009 ... </th>
        // 공지 글: <th> ... img ... </th>
        const thMatch = row.match(/<th[^>]*>([\s\S]*?)<\/th>/i);
        if (!thMatch) continue;

        let idStr = thMatch[1].replace(/<[^>]+>/g, "").trim(); // 태그 제거 후 텍스트만
        // 숫자가 아니면(공지 등) 스킵
        if (!/^\d+$/.test(idStr)) continue;

        const postId = parseInt(idStr, 10);

        // 2. 제목 및 링크 추출 (td.subject 안의 a 태그)
        const subjectMatch = row.match(/<td[^>]*class=["']?subject["']?[^>]*>([\s\S]*?)<\/td>/i);
        if (!subjectMatch) continue;

        const linkMatch = subjectMatch[1].match(/<a[^>]*href=["']?([^"'>]+)["']?[^>]*>([\s\S]*?)<\/a>/i);
        if (!linkMatch) continue;

        let link = linkMatch[1];
        let titleRaw = linkMatch[2];

        // 제목 정제 (태그 제거, 공백 정리)
        let title = titleRaw.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/\s+/g, " ").trim();

        if (link.startsWith("/")) {
            link = "https://www.humetro.busan.kr" + link;
        }

        posts.push({
            id: postId,
            title: title,
            link: link
        });
    }

    console.log(`파싱 결과: 총 ${posts.length}개의 유효 게시글 발견`);
    return posts;
}
