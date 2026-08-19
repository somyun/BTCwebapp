# T2 테스트 환경 리소스 기록

- 기록일: 2026-07-29
- 상태: T2 테스트 GAS 공개, 프런트엔드 병합·배포 및 공개 읽기 검증 완료
- 운영 리소스 변경: 없음

## Firebase

- 프로젝트 ID: `btcwebapp-test`
- 웹앱 닉네임: `bwa_test-web`
- 웹앱 ID: `1:119870762952:web:9418a56f9dd72963bf5afd`
- Firestore: `(default)`, Standard, `asia-northeast3`
- Firestore 초기 규칙: `allow read, write: if false;`
- FCM Sender ID: `119870762952`
- Web Push 공개 키: `BNPH82AhKzYvUyBlY-DodY1mfWR1Dw_Lbe5rG0mmvE8o2asNscCvdzW2KCLflVsLuKKW0SJpr6CslPFpxVkMiB4`

```js
const firebaseConfig = {
  apiKey: "AIzaSyCLoslBUFGIqCXhSq1gm_jbqEXEMPx-6ow",
  authDomain: "btcwebapp-test.firebaseapp.com",
  projectId: "btcwebapp-test",
  storageBucket: "btcwebapp-test.firebasestorage.app",
  messagingSenderId: "119870762952",
  appId: "1:119870762952:web:9418a56f9dd72963bf5afd"
};
```

Firebase 웹 설정과 Web Push 키는 브라우저 클라이언트에 포함되는 공개 식별자다. 서비스 계정 개인키는 이 문서와 테스트 소스에 저장하지 않는다.

## Google Drive와 Sheets

- 테스트 폴더: `BTCwebapp_TEST_bwa_test`
- 테스트 폴더 ID: `17drZL2aYVyfrxBUG9Ne0b_frovZhgvEC`
- 테스트 스프레드시트: `ERP점검웹앱_TEST_bwa_test`
- 테스트 스프레드시트 ID: `1vdd9Z78My2f8TCHoDqgF5TZb4TxNsEJgqEwRVNtdwi8`
- 공유 범위: 소유자 전용

## Apps Script

- 테스트 프로젝트 원격 이름: `ERP웹앱_TEST_bwa_test`
- 테스트 Script ID: `1ynXimyiusVX7LqUqYFMW4cApQ5A0gtkOTfCx1dgkVQY-qQ5n1EoN76zv`
- 로컬 소스: `apps-script-test/`
- `clasp push --force`: 완료
- 운영 spreadsheet/folder/Script ID 포함 여부: 0건
- `firebase_key.gs`: 원격 파일 목록에서 제거 확인
- 테스트 환경 가드: Script ID, spreadsheet ID, Drive folder ID, Firebase project ID 고정
- 공개 테스트 배포 ID: `AKfycbxSPNbqB8xKK0eMskBZupnSHS4RbiKX3CaWKmchB3-v7W1AQSQs7kssAraUepHlas1T`
- 공개 테스트 URL: `https://script.google.com/macros/s/AKfycbxSPNbqB8xKK0eMskBZupnSHS4RbiKX3CaWKmchB3-v7W1AQSQs7kssAraUepHlas1T/exec`
- 공개 테스트 배포 버전: `@4` (`ANYONE_ANONYMOUS`)

### 1회 초기화 및 격리 확인

2026-07-29 최초 OAuth 승인 후 `initializeTestEnvironment` 실행을 완료했다. 다음 항목은 테스트 프로젝트 안에서만 정리됐다.

1. `FCM_Tokens` 내용을 삭제하고 헤더만 재생성
2. `FCM_Logs` 내용을 삭제하고 헤더만 재생성
3. 복사된 Script Properties 전체 삭제
4. 복사된 설치형 트리거 전체 삭제
5. `FormList`의 스프레드시트 ID를 테스트 스프레드시트 ID로 제한

추가 검증 결과:

- 전체 테스트 스프레드시트에서 운영 스프레드시트 ID 검색 결과: 0건
- 테스트 GAS 소스의 운영 Script/spreadsheet/folder ID 참조: 0건
- 서비스 계정 개인키 파일: 없음
- JavaScript 문법 검사: 통과
- 임시 초기화 HTTP 경로와 실행 API 설정: 제거 완료
- `getFormList()` 응답의 `spreadsheetId`: 테스트 ID로 강제
- 공개 기본 응답·환경 상태·양식 목록·선택 양식 요청: 성공
- 공개 환경 상태: Script/Spreadsheet ID 일치, FCM 데이터·Script Properties·트리거·비테스트 FormList ID 모두 0
- GitHub Pages Origin CORS: HTTP 200, `Access-Control-Allow-Origin: *`

### T2 프런트엔드 검증

`testbed/bwa_test_publish/`에 테스트 GAS 읽기 전용 연결을 적용했다.

- 네트워크 허용: 동일 출처 정적 GET/HEAD, 승인된 테스트 GAS GET
- 네트워크 차단: 운영 GAS/Firebase, POST, XHR, WebSocket, EventSource
- 테스트 Firebase config와 VAPID 공개 키: 테스트 값으로 기록
- 양식 목록: 6개 표시
- 선택 양식: 33개 항목 표시
- 업로드·알림·즐겨찾기 초기화: 비활성
- 브라우저 오류·경고: 0건
- 로컬 실측 예시: 양식 목록 약 1.9초, 선택 양식 약 4.5초(각 요청의 네트워크·GAS 처리 포함)

### GitHub 원격 게시 및 공개 URL 검증

- 저장소: `somyun/bwa_test`
- 작업 브랜치: `codex/t2-test-gas-read` (원격 보존)
- 작업 커밋: `9e75480`
- PR: `#1` (`T2: connect test frontend to isolated GAS read API`)
- 병합 방식: squash merge
- `main` 병합 커밋: `1f555cdae6368d542b1ae6ce133f757bb522c1fa`
- 병합 시각: 2026-07-29 11:17:45 KST
- GitHub Pages 워크플로 실행: `30416447099`, 성공
- 공개 URL: `https://somyun.github.io/bwa_test/`
- 공개 페이지 양식 목록: 6개, 약 3.2초
- 공개 페이지 `율리24`: 33개 항목, 약 4.0초
- 공개 페이지 업로드·알림·즐겨찾기 초기화: 비활성
- 공개 페이지 브라우저 오류·경고: 0건
- 운영 저장소 `somyun/BTCwebapp` 및 운영 서비스 변경: 없음

## 미완료

- 테스트 FCM 기기 등록과 송수신 검증
