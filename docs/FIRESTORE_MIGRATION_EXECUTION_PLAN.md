# Firestore 게시 캐시 및 비동기 저장 이전 실행계획

- 개정일: 2026-07-29
- 대상 운영 사이트: `https://somyun.github.io/BTCwebapp/`
- 테스트 사이트 예정 주소: `https://somyun.github.io/bwa_test/`
- 문서 상태: 테스트베드 우선·운영 무중단 실행계획
- 최우선 조건: 운영 GitHub Pages, GAS, 시트 및 FCM에 영향 금지

## 1. 현재 구조 재확인 결과

### 1.1 저장소에서 확인된 사실

현재 저장소에는 다음 Firebase 인프라 파일이 없다.

- `firebase.json`
- `.firebaserc`
- `firestore.rules`
- `firestore.indexes.json`
- `functions/`
- Functions용 `package.json`

현재 프런트엔드의 Firebase 사용은 FCM에 한정된다.

- [`index.html`](../index.html#L146)은 Firebase Messaging compat SDK만 로드한다.
- [`script.js`](../script.js#L10)은 Firebase 앱과 Messaging만 초기화한다.
- [`firebase-messaging-sw.js`](../firebase-messaging-sw.js#L1)은 백그라운드 FCM 수신만 처리한다.
- Firestore SDK 초기화, Firestore 읽기/쓰기 및 Callable Functions 호출 코드는 없다.

현재 서버 기능은 별도 Apps Script 프로젝트가 담당한다.

- 양식 목록 및 선택 양식 조회
- 측정값 저장
- 엑셀 업로드 및 XLSX 다운로드
- FCM 토큰·로그 시트 관리
- FCM HTTP v1 발송
- 게시판 조회와 정기 트리거

### 1.2 Firebase 콘솔 상태에 대한 판단

2026-07-29 Firebase 콘솔에서 운영 프로젝트 `btcwebapp-551bd`를 직접 확인했다.

- 요금제: Spark
- Firestore: `데이터베이스 만들기` 상태로 DB 없음
- Functions: 배포 목록 없음, 요금제 업그레이드 안내 상태

따라서 운영 Firebase 프로젝트는 현재 FCM 용도로만 사용하며, 운영 Firestore와 Functions는 아직 사용하지 않는 것으로 확정한다. 운영 단계 진입 전에는 변경 여부만 다시 확인한다.

## 2. 확정된 전략

### 2.1 승인된 방향

1. 양식 목록과 선택 양식을 Firestore 게시용 캐시로 구성한다.
2. GitHub Pages가 Firestore 게시 캐시를 직접 조회한다.
3. 테스트 환경의 Firestore 조회 실패는 오류로 종료하며 레거시 GAS로 자동 폴백하지 않는다.
4. 읽기 전환 후에는 Firestore 우선 저장과 비동기 시트 동기화를 검증한다.
5. XLSX 다운로드는 시트 동기화 완료 후 활성화되어도 된다.
6. 모든 개발과 위험 검증은 `bwa_test` 테스트베드에서 먼저 수행한다.

### 2.2 현재 보류 범위

1. 브라우저 영구 캐시(IndexedDB persistence)
2. 운영 Apps Script의 기존 측정값 저장 로직 최적화
3. 운영 Firebase 프로젝트에 Firestore/Functions 생성
4. 운영 사이트의 Firestore 읽기 전환
5. 운영 사이트의 Firestore 우선 저장 전환

3~5번은 테스트베드 검증 결과를 보고 별도로 승인한 뒤 수행한다.

## 3. 환경 완전 분리 원칙

GitHub 저장소만 분리해서는 안전하지 않다. 다음 항목을 모두 분리해야 한다.

| 구분 | 운영 환경 | 테스트 환경 |
|---|---|---|
| GitHub 저장소 | `somyun/BTCwebapp` | `somyun/bwa_test` |
| Pages 경로 | `/BTCwebapp/` | `/bwa_test/` |
| Firebase 프로젝트 | `btcwebapp-551bd` | `btcwebapp-test` |
| Firestore | 2026-07-29 콘솔 확인: DB 없음 | `(default)`, Standard, `asia-northeast3` |
| Functions | 2026-07-29 콘솔 확인: 배포 없음 | 코드 구현 후 테스트 프로젝트에 최초 배포 |
| FCM | 운영 사용자 토큰 | 테스트 기기 전용 |
| VAPID 키 | 운영 키 | 테스트 키 |
| Apps Script | 현재 운영 Script ID | 신규 테스트 Script ID |
| GAS 웹앱 URL | 현재 운영 URL | 신규 테스트 URL |
| 스프레드시트 | 운영 원본 | 운영 시트 복사본 |
| Drive 폴더 | 운영 업로드 폴더 | 테스트 전용 폴더 |
| 정기 트리거 | 운영 트리거 | 기본 비활성, 필요 시 테스트 전용 |
| 외부 API | 운영 Kakao/Humetro | 기본 비활성 또는 테스트 자격 증명 |

### 절대 금지 사항

- `bwa_test`에서 운영 GAS URL 호출
- 테스트 Function에서 운영 스프레드시트 접근
- 테스트 FCM에서 운영 프로젝트 VAPID 키 사용
- 테스트 시트에 운영 `FCM_Tokens`와 `FCM_Logs` 복사
- 테스트 Apps Script에 운영 서비스 계정 개인키 복사
- 테스트용 Firestore Rules 또는 Functions를 운영 Firebase 프로젝트에 배포
- 테스트 결과 승인 전에 운영 저장소 코드 변경 또는 Pages 재배포

## 4. 테스트베드 목표 구조

```text
https://somyun.github.io/bwa_test/
  │
  ├─ 읽기: 테스트 Firestore
  │    ├─ publicCache/formList
  │    └─ publicForms/{formKey}
  │
  ├─ 게시 원본·검증: 테스트 GAS URL
  │    └─ 테스트 스프레드시트 복사본
  │
  ├─ 저장 접수: 테스트 Firebase Function
  │    └─ measurementSubmissions/{submissionId}
  │         └─ 테스트 시트 동기화 Worker
  │
  └─ 알림: 테스트 FCM
       └─ 등록된 테스트 기기만 대상
```

운영 구조는 테스트 기간 동안 그대로 유지한다.

```text
https://somyun.github.io/BTCwebapp/
  └─ 현재 GAS URL
       └─ 운영 스프레드시트
```

## 5. 테스트베드 보호 장치

### 5.1 환경 식별 배너

테스트 사이트 모든 화면 상단에 다음 배너를 고정 표시한다.

```text
TEST ENVIRONMENT · 운영 데이터가 아닙니다
```

운영 사이트와 다른 색상을 사용해 잘못된 사이트에서 입력하는 실수를 방지한다.

### 5.2 시작 시 환경 가드

테스트 사이트는 시작할 때 설정을 검증하고 하나라도 운영값이면 API 호출을 차단한다.

검증 항목:

- URL 경로가 `/bwa_test/`인지 확인
- Firebase project ID가 테스트 프로젝트인지 확인
- GAS URL이 운영 URL과 다른지 확인
- spreadsheet ID가 테스트 복사본인지 확인
- VAPID 키가 테스트 키인지 확인
- 환경명이 `test`인지 확인

개념 예시:

```javascript
assertTestEnvironment({
  pathname: location.pathname,
  environment: APP_CONFIG.environment,
  firebaseProjectId: APP_CONFIG.firebase.projectId,
  gasApiUrl: APP_CONFIG.gasApiUrl
});
```

가드 실패 시 빨간 오류 화면만 표시하고 GAS, Firestore, FCM 요청을 보내지 않는다.

### 5.3 서버 측 이중 가드

클라이언트 가드는 우회될 수 있으므로 Functions와 GAS도 대상 리소스를 검증한다.

- 테스트 Function 환경변수에 허용된 테스트 spreadsheet ID를 고정한다.
- 요청에서 spreadsheet ID를 직접 받지 않는다.
- 테스트 GAS는 Script Properties의 테스트 ID만 사용한다.
- 서비스 계정에는 테스트 시트만 공유한다.
- Function 시작 시 Firebase project ID가 테스트 프로젝트인지 확인한다.

### 5.4 PWA 및 서비스 워커 격리

현재 manifest와 서비스 워커 등록은 상대 경로를 사용한다.

- `manifest.json`: `start_url`, `scope`가 `./`
- `script.js`: `./firebase-messaging-sw.js` 등록

따라서 테스트 Pages에서는 `/bwa_test/` scope로 제한되어야 한다. 배포 후 브라우저 개발자 도구에서 실제 scope가 다음인지 검증한다.

```text
https://somyun.github.io/bwa_test/
```

운영 scope `/BTCwebapp/`를 침범하면 배포를 중지한다.

## 6. 환경별 설정 관리

운영값과 테스트값을 소스 곳곳에 직접 작성하지 않는다. 환경 설정을 한 곳으로 모은다.

```text
config/
  app-config.js
```

테스트 설정 예시:

```javascript
window.APP_CONFIG = {
  environment: 'test',
  expectedPathPrefix: '/bwa_test/',
  gasApiUrl: 'TEST_GAS_URL',
  firebase: {
    projectId: 'TEST_FIREBASE_PROJECT_ID'
  },
  readSource: 'gas',
  enableFcm: false
};
```

장기적으로 GitHub Actions가 저장소별 Variables에서 설정 파일을 생성하도록 한다. 설정 파일이 다르다는 이유로 검증된 기능 코드가 운영·테스트 간에 갈라지지 않게 한다.

운영 반영 PR에서는 다음 값이 변경 대상에 포함되지 않았는지 자동 검사한다.

- 운영 GAS URL
- 운영 Firebase config
- 운영 VAPID 키
- 운영 spreadsheet ID
- Apps Script `.clasp.json`

## 7. 테스트 데이터 준비

### 7.1 스프레드시트 복사본

1. 운영 시트를 복사해 테스트 시트를 만든다.
2. 테스트 시트 이름에 `[TEST]` 접두사를 붙인다.
3. `FCM_Tokens`의 운영 토큰을 모두 제거한다.
4. `FCM_Logs`의 운영 로그를 제거한다.
5. 필요하다면 개인정보·민감정보를 마스킹한다.
6. 양식, FormList, DB 구조는 성능과 변환 검증을 위해 유지한다.
7. 테스트 spreadsheet ID를 별도로 기록한다.

### 7.2 테스트 Drive 폴더

- 테스트 업로드 파일 전용 폴더를 만든다.
- 테스트 GAS와 Function만 접근하도록 공유한다.
- 운영 폴더 ID는 테스트 설정에 넣지 않는다.

### 7.3 테스트 Apps Script

1. 현재 Apps Script 소스를 신규 프로젝트로 복제한다.
2. 신규 Script ID로 별도 `.clasp.json`을 구성한다.
3. `TARGET_SPREADSHEET_ID`, `GLOBAL_TARGET_SPREADSHEET_ID`, `FOLDER_ID`를 테스트 Script Properties로 분리한다.
4. 별도 웹앱으로 배포해 테스트 GAS URL을 만든다.
5. 익명 접근 범위는 기존 기능 재현에 필요한 최소 수준으로 설정한다.
6. 시간 기반 트리거는 기본 비활성화한다.
7. Kakao/Humetro 호출은 테스트 범위가 아니면 비활성화한다.

복제 시 운영 `firebase_key.js`를 테스트 프로젝트로 복사하지 않는다. 테스트 FCM은 테스트 프로젝트의 별도 자격 증명 또는 권장되는 서버 기본 자격 증명을 사용한다.

### 7.4 테스트 FCM

1. 테스트 Firebase 프로젝트에 별도 웹앱을 등록한다.
2. 별도 VAPID 키를 발급한다.
3. `firebase-messaging-sw.js`와 프런트 config를 테스트 프로젝트로 변경한다.
4. 개발자 테스트 기기만 토큰을 등록한다.
5. 초기 Firestore 읽기 검증 동안에는 FCM을 비활성화해도 된다.
6. FCM 검증 단계에서만 명시적으로 활성화한다.

## 8. Firestore 게시 데이터 모델

### 8.1 양식 목록

경로: `publicCache/formList`

```json
{
  "schemaVersion": 1,
  "sourceRevision": "ISO-8601",
  "publishedAt": "server timestamp",
  "contentHash": "sha256",
  "itemCount": 10,
  "items": [
    {
      "formKey": "stable-url-safe-key",
      "sheetName": "sheet name",
      "displayName": "display name",
      "lastModifiedDate": "ISO-8601"
    }
  ]
}
```

### 8.2 선택 양식

경로: `publicForms/{formKey}`

```json
{
  "schemaVersion": 1,
  "formKey": "stable-url-safe-key",
  "sheetName": "sheet name",
  "sourceRevision": "ISO-8601",
  "publishedAt": "server timestamp",
  "contentHash": "sha256",
  "rowCount": 120,
  "rows": [
    {
      "uniqueId": "id",
      "location": "location",
      "item": "item",
      "value": "value",
      "unit": "unit",
      "validation": {
        "minValue": "min",
        "maxValue": "max"
      },
      "recentInfo": {
        "value": "latest value",
        "date": "latest date"
      }
    }
  ]
}
```

- `items`와 `rows`는 쿼리에 사용하지 않으므로 인덱스 제외를 적용한다.
- 문서 직렬화 크기가 900KiB를 넘으면 manifest와 `chunks` 하위 컬렉션으로 분할한다.
- 기존 GAS 응답 형태와 동일하게 만들어 UI 변경 범위를 줄인다.

### 8.3 측정값 제출

경로: `measurementSubmissions/{idempotencyKey}`

```json
{
  "schemaVersion": 1,
  "idempotencyKey": "client-generated-uuid",
  "formKey": "form key",
  "formRevision": "revision seen by client",
  "measurements": [],
  "status": "pending",
  "attemptCount": 0,
  "createdAt": "server timestamp",
  "processingStartedAt": null,
  "syncedAt": null,
  "failedAt": null,
  "errorCode": null
}
```

상태 전이:

```text
pending → processing → synced
                    └→ failed
```

## 9. 단계별 실행계획

### Phase T0. 운영 기준선 고정

운영 시스템에 쓰기 변경 없이 수행한다.

1. 현재 운영 Git 커밋 SHA를 기록한다.
2. `clasp pull`로 운영 GAS 최신 소스를 확인한다.
3. 운영 GAS Script ID, 배포 ID, URL을 기록한다.
4. 운영 사이트의 목록·양식·저장 p50/p95 기준선을 수집한다.
5. 운영 사이트 주요 기능의 수동 회귀 체크리스트를 만든다.

완료 조건:

- 운영 복구 지점과 성능 기준선이 문서화되어 있다.

### Phase T1. `bwa_test` GitHub Pages 생성

1. `somyun/bwa_test` 저장소를 만든다.
2. 운영 저장소의 고정된 커밋을 복제한다.
3. GitHub Pages를 `/bwa_test/`에 배포한다.
4. TEST 배너를 추가한다.
5. 환경 가드를 추가한다.
6. 아직 운영 GAS를 호출하지 않는다.
7. 테스트 GAS가 준비되기 전에는 조회·저장 버튼을 비활성화한다.
8. manifest, 아이콘, 상대경로, 서비스 워커 scope를 검증한다.

완료 조건:

- 테스트 사이트가 독립 URL에서 열리고 운영 백엔드 요청이 한 건도 발생하지 않는다.

실행 결과(2026-07-29):

- 공개 저장소 `somyun/bwa_test` 생성
- 기준 운영 커밋 `2e2dc00138c4585178494556aaeea65497f20262`의 화면 자산을 T1 정적 셸로 분리
- 테스트 커밋 `c3b135d`를 `main`에 게시
- GitHub Pages `https://somyun.github.io/bwa_test/` 활성화
- TEST 배너 표시, 조회·저장·업로드·알림 제어 14개 전체 비활성화 확인
- CSP `connect-src 'none'` 및 JavaScript 전송 가드 적용
- 로드 자산에서 운영 GAS·Firebase·FCM 주소 0건 확인
- 서비스 워커 스크립트와 manifest scope를 `/bwa_test/`로 고정
- 운영 복구 식별자가 포함된 T0 기준선은 공개 저장소에서 제외하고 로컬 문서로만 보관

### Phase T2. 테스트 백엔드 완전 분리

1. 테스트 Firebase 프로젝트를 만든다.
2. Blaze 요금제를 연결하고 테스트 Firestore를 활성화한다.
3. Functions 코드를 구현한 뒤 테스트 프로젝트에 최초 배포한다.
4. 테스트 웹앱과 FCM을 등록한다.
5. 운영 시트 복사본과 테스트 Drive 폴더를 만든다.
6. 테스트 Apps Script 프로젝트를 만들고 테스트 시트에 연결한다.
7. 테스트 GAS URL을 발급한다.
8. 테스트 사이트의 `GAS_API_URL`, Firebase config, VAPID 키를 테스트값으로 교체한다.
9. 클라이언트와 서버 환경 가드를 통과하는지 확인한다.
10. 네트워크 로그에서 운영 GAS, 운영 Firebase, 운영 시트 접근이 없는지 확인한다.

완료 조건:

- 양식 조회, 저장, 업로드, 다운로드, FCM 테스트가 모두 테스트 리소스 안에서만 이루어진다.

부분 실행 결과(2026-07-29):

- Firebase 프로젝트 표시 이름과 프로젝트 ID를 모두 `btcwebapp-test`로 생성
- Google Analytics와 Gemini는 사용하지 않도록 생성
- 승인된 결제 계정에 Blaze 요금제를 연결하고 월 10,000원 예산 알림 설정
- Cloud Firestore `(default)` 데이터베이스를 Standard 버전, `asia-northeast3`(서울)에 생성
- 초기 보안 규칙이 `allow read, write: if false;`인 프로덕션 모드임을 콘솔에서 재확인
- 운영 Firebase 프로젝트 `btcwebapp-551bd`에는 변경하지 않음
- Functions는 별도 생성 버튼이 있는 리소스가 아니므로, T3 코드 구현 후 `btcwebapp-test`에 최초 배포할 때 활성화
- Firebase 웹앱 `bwa_test-web`과 전용 Web Push 공개 키 생성
- 전용 Drive 폴더 `BTCwebapp_TEST_bwa_test`와 소유자 전용 테스트 시트 사본 생성
- 시트 사본에 바인딩된 테스트 GAS Script ID를 발급하고 `apps-script-test/` 소스를 `clasp`로 분리
- 테스트 GAS에서 운영 리소스 ID와 `firebase_key.gs`를 제거하고 테스트 환경 가드 적용
- 최초 GAS OAuth 승인과 `initializeTestEnvironment` 실행 완료: FCM 토큰·로그, Script Properties, 트리거 정리
- 테스트 전체 시트에서 운영 스프레드시트 ID 검색 0건 확인, `getFormList()` 응답은 테스트 ID로 강제
- 테스트 GAS 프로젝트를 `ERP웹앱_TEST_bwa_test`로 변경하고 소유자 전용 안전 기준 배포 `@3` 생성
- 임시 초기화 HTTP 경로와 실행 API 설정은 검증 후 제거
- 승인된 범위로 테스트 GAS 익명 공개 배포 `@4` 생성, 비로그인 조회와 GitHub Pages Origin CORS 검증 통과
- `bwa_test` 프런트엔드에 테스트 GAS 읽기 전용 연결을 적용하고 양식 목록 6개·선택 양식 33개 항목 표시 검증
- PR `#1`을 squash merge하고 `main` 병합 커밋 `1f555cdae6368d542b1ae6ce133f757bb522c1fa`를 GitHub Pages에 배포
- 공개 URL에서 양식 목록 6개·`율리24` 33개 항목, 변경 기능 비활성, 브라우저 오류·경고 0건을 재검증
- 아직 미수행: Functions 구현 및 배포

### Phase T3. 테스트 게시 캐시 생성기

테스트 Functions에 다음을 구현한다.

1. `publishFormList`
   - 테스트 FormList를 읽어 `publicCache/formList`에 게시한다.
2. `publishForm`
   - 테스트 양식과 DB를 결합해 `publicForms/{formKey}`에 게시한다.
3. `publishAllChangedForms`
   - revision 또는 hash가 변경된 양식만 갱신한다.
4. 수동 재게시 명령
   - 특정 양식과 전체 양식을 재게시할 수 있게 한다.
5. 테스트 Scheduler
   - 초기에는 2~5분 간격으로 실행한다.

이 단계에서 `bwa_test` 화면은 계속 테스트 GAS 결과만 사용한다.

완료 조건:

- 테스트 GAS 응답과 Firestore 게시 문서가 자동 비교 기준을 통과한다.

2026-07-29 완료 결과:

- Node.js 22 2세대 Functions 4개를 `btcwebapp-test`, `asia-northeast3`에 배포
- 5분 Scheduler가 최초 6개 양식을 게시하고 다음 실행에서 변경 0개를 확인
- `publicCache/formList`와 `publicForms/{formKey}` 6개 양식 게시 완료
- 테스트 GAS와 목록 hash, 양식 hash, 전체 행 데이터 자동 비교 통과
- 공개 읽기 성공과 클라이언트 직접 쓰기 HTTP 403 확인
- 화면 데이터 소스는 테스트 GAS를 계속 유지

### Phase T4. GAS → Shadow → Firestore 읽기 전환

테스트 사이트에 데이터 소스 어댑터를 추가한다.

```text
gas        테스트 GAS만 화면에 사용
shadow     테스트 GAS를 화면에 사용하고 Firestore는 비교만 수행
firestore  테스트 Firestore만 사용하고 실패 시 오류 표시
```

진행 순서:

1. `gas` 모드에서 기존 기능 회귀 확인
2. `shadow` 모드에서 목록과 양식 데이터 비교
3. 불일치 0건 또는 승인된 차이만 남을 때까지 수정
4. `firestore` 모드로 전환
5. 네트워크, 권한, 문서 누락, stale, schema 오류별 실패 화면과 재시도 동작 확인
6. 실패율과 성능 측정

주의:

- Shadow 모드에서도 운영 GAS는 절대 호출하지 않는다.
- 브라우저 영구 캐시는 아직 활성화하지 않는다.

완료 조건:

- Firestore 조회 성능 목표 달성
- Firestore 장애 시 레거시 요청 없이 안전한 오류 표시
- 운영 리소스 요청 0건

2026-07-29 완료 기록:

- `bwa_test` PR #3에서 GAS/Shadow/Firestore 어댑터와 무폴백 오류·재시도를 배포
- 공개 Shadow에서 목록과 6개 양식 전체 행 불일치 0건 확인
- 공개 선택 양식 p50 GAS 약 2,511ms, Firestore 약 40ms로 약 98% 단축
- PR #4에서 테스트 사이트 기본 읽기를 Firestore로 전환
- PR #5에서 정적 자산 버전을 부여해 이전 스크립트 HTTP 캐시 재사용 방지
- 최종 Pages 배포 커밋 `3ebc6194709b9d6c0e6b349b0b16d0dd6757dc19`

### Phase T5. Firestore 우선 저장과 비동기 테스트 시트 동기화

1. `submitMeasurements` 테스트 Function을 구현한다.
2. 멱등성 키로 `measurementSubmissions` 문서를 생성한다.
3. Function은 시트 반영을 기다리지 않고 접수 상태를 반환한다.
4. Firestore Trigger Worker가 테스트 시트에 `values.batchUpdate`를 수행한다.
5. 동일 이벤트 재실행에도 중복 부작용이 없도록 멱등 처리한다.
6. 성공하면 `synced`, 실패하면 `failed`로 변경한다.
7. 양식 게시 캐시와 목록 revision을 다시 게시한다.
8. 클라이언트는 `synced` 후 XLSX 버튼을 활성화한다.
9. 네트워크 타임아웃 시 동일 멱등성 키의 상태를 먼저 확인한다.
10. 접수 후 운영 GAS로 자동 재전송하지 않는다.

완료 조건:

- 저장 접수부터 테스트 시트 반영까지 상태 추적 가능
- 중복 반영 0건
- 실패 재처리 검증 완료
- XLSX 비상 기능 정상

2026-07-29 완료 기록:

- `submitMeasurements`, `getMeasurementSubmission`, `syncMeasurementSubmission`, `retryMeasurementSubmission`, `setSubmissionGate`를 `btcwebapp-test`의 `asia-northeast3`에 배포
- 브라우저 Firestore 직접 쓰기는 계속 거부하고, 허용 Origin의 공개 Function만 제출·상태 확인 가능하도록 구성
- `measurementSubmissions/{idempotencyKey}` 접수, 분당 12건 제한, 128KiB·250행 상한, 양식 리비전·전체 행 정체성 검증 적용
- 테스트 프로젝트에서 Sheets API를 활성화하고 기본 2세대 Functions 서비스 계정에 테스트 시트만 편집 권한 부여
- Firestore 생성 이벤트 워커가 `율리24` 33개 측정값과 `FormList` 리비전 1개, 총 34개 셀을 자동 동기화한 것을 확인
- 새 자동 제출 `c93db1f4-61b5-47bd-a4d6-bc553222e75c`이 1회 시도로 `synced` 완료되고 약 6초 안에 브라우저 완료 상태가 표시됨
- 동일 완료 제출 재처리는 458ms에 기존 상태만 반환해 추가 Sheets 부작용이 없음을 확인
- 공개 상태 응답에 `measurements`와 `requestHash`가 없음을 확인
- 최초 워커 초기화 오류를 수동 재처리해 복구 경로를 실검증하고, Firebase Admin 기본 앱 초기화를 보강
- 저장 직후 GAS 캐시의 이전 리비전으로 목록·양식이 어긋나는 문제를 발견하여 제출 데이터 기반 동시 리비전 게시와 `stale_skipped` 역행 방지를 적용
- 동기화 후 페이지를 새로 열어 Firestore 양식 33개가 stale 오류 없이 조회되고, 비상용 `율리24_260729.xlsx` 버튼이 `synced` 뒤 활성화됨을 확인
- Functions 자동 테스트 29개, 프런트 스크립트 문법 검사 통과
- 제출 게이트 기본 차단과 관리자 Secret 기반 활성화·수동 재처리 절차 검증, Draft PR 병합 전 `enabled: false`로 재차단
- 운영 GitHub Pages, 운영 GAS, 운영 시트, 운영 Firebase·FCM 변경 없음

2026-07-29 공개 테스트베드 전환 기록:

- `bwa_test` PR #6을 squash merge하고 main 커밋 `63c429dd150b69ddc5eb5ad0248a84a531a4042e`로 반영
- GitHub Pages 실행 `30424323695`의 build·deploy 성공 확인
- 공개 기본 Firestore 모드에서 목록 6개 약 37ms, `율리24` 33개 항목 약 34ms 조회 확인
- 공개 제출 `7d417206-4a6b-4524-848d-6d771ce40720`이 1회 시도로 34개 셀을 갱신하고 서버 기준 약 4.1초에 `synced` 완료
- 공개 페이지에서 동기화 완료 뒤 `율리24_260729.xlsx` 비상 다운로드 활성화 확인
- 저장 직후 새로고침에서도 목록·양식 리비전 불일치 없이 Firestore 양식 재조회 성공
- 공개 GAS 비교 모드는 목록 약 5.3초, 선택 양식 약 3.0초였으며 저장 버튼이 비활성화된 읽기 전용 상태 확인
- 배포된 `index.html`·`testbed.js`에 테스트 프로젝트와 테스트 GAS만 포함되고 운영 Script ID 및 `/BTCwebapp` 경로가 없음을 확인
- 공개 상태 응답에 측정값·요청 해시가 없고 `attemptCount: 1`, `updatedCellCount: 34`임을 확인
- 장기 테스트를 위해 테스트 제출 게이트를 `enabled: true`로 유지; 운영 리소스에는 영향 없음

### Phase T6. 테스트베드 장기 검증

최소 검증 항목:

- 여러 모바일·데스크톱 브라우저
- 여러 양식 크기
- 동시 저장
- Firestore 일시 장애
- Functions 재시도
- Sheets API 일시 장애
- stale 게시 캐시
- FCM 테스트 기기 알림
- 서비스 워커 scope
- Firestore 비용과 호출량

테스트 결과 보고서에 다음을 포함한다.

- 목록/양식/저장 p50 및 p95
- Firestore 조회 실패율과 사유
- 캐시 불일치 건수
- 시트 동기화 평균·최대 지연
- 실패 및 수동 재처리 건수
- Firestore 읽기·쓰기 사용량

### Phase P0. 운영 도입 승인 게이트

테스트베드가 완료되어도 운영에 자동 반영하지 않는다. 다음을 사용자에게 보고하고 별도 승인을 받는다.

1. 테스트 결과와 성능 개선 폭
2. 발견된 오류와 잔여 위험
3. 운영 Firebase에 새로 생성할 리소스
4. 예상 비용
5. 운영 변경 파일 목록
6. 운영 배포·롤백 절차

승인 전에는 `BTCwebapp`, 운영 GAS, 운영 Firebase, 운영 시트를 변경하지 않는다.

### Phase P1. 운영 Firestore/Functions 준비

별도 승인 후 수행한다.

1. `btcwebapp-551bd`가 FCM 외 Firestore/Functions를 사용하지 않는지 콘솔에서 재확인한다.
2. 운영 Firestore 위치를 확정한다. 생성 후 변경이 어려우므로 별도 승인한다.
3. 운영 Functions와 서비스 계정을 구성한다.
4. 운영 시트는 읽기 publisher 권한부터 최소 범위로 공유한다.
5. Firestore Rules는 기본 거부로 배포한다.
6. App Check는 모니터링 모드로 시작하고 즉시 enforcement하지 않는다.
7. 운영 게시 캐시를 백필하되 사이트는 계속 GAS만 사용한다.

### Phase P2. 운영 GAS → Shadow → Firestore 읽기 전환

1. 운영 사이트를 `gas` 모드 코드로 먼저 배포한다.
2. 회귀가 없으면 `shadow` 모드로 변경한다.
3. 운영 GAS와 운영 Firestore 결과를 비교한다.
4. 검증 후 `firestore` 모드로 전환한다.
5. 장애 시 운영 GAS로 자동 폴백한다.
6. 문제가 생기면 `READ_SOURCE=gas`로 즉시 되돌린다.

운영 Firestore 우선 저장은 읽기 안정화 후 다시 승인받아 별도 단계로 진행한다.

## 10. Firestore Rules와 권한

테스트와 운영 모두 동일한 정책을 사용하되 프로젝트는 분리한다.

```text
기본: 모든 read/write 거부
publicCache: 클라이언트 read만 허용
publicForms: 클라이언트 read만 허용
measurementSubmissions: 클라이언트 직접 write 금지
Functions: Admin SDK와 IAM으로 접근
```

현재 GAS 웹앱이 익명 접근이므로 읽기 게시 데이터의 공개 범위를 배포 전에 확인해야 한다. 외부 공개가 불가능한 데이터라면 Firebase Authentication을 먼저 적용한다.

App Check 적용 순서:

1. 테스트 프로젝트 등록
2. 테스트 메트릭 확인
3. 테스트 enforcement
4. 운영 도입 시 모니터링 모드
5. 정상 요청 검증 후 운영 enforcement

## 11. 테스트 및 승인 기준

### 데이터 일치

- FormList 항목 수와 식별자 일치
- 선택 양식 행 수와 `uniqueId` 집합 일치
- 값, 단위, 최소·최대, 최근값 일치
- 날짜·숫자 정규화 차이는 문서화
- 문서/청크 hash 일치

### 기능 회귀

- 양식 목록
- 선택 양식
- 즐겨찾기
- 유효성 경고
- 측정값 저장
- 업로드
- XLSX 생성과 다운로드
- 테스트 FCM
- 모바일 PWA

### 격리 검증

- 테스트 사이트에서 운영 GAS URL 요청 0건
- 테스트 서비스 계정의 운영 시트 권한 없음
- 테스트 FCM에서 운영 토큰 0건
- 테스트 Functions가 운영 Firebase 프로젝트에 없음
- 테스트 GAS가 운영 spreadsheet ID를 포함하지 않음
- 서비스 워커 scope가 `/bwa_test/`로 제한됨

### 성능 목표

- 양식 목록 p50/p95 현행 대비 50% 이상 단축 목표
- 선택 양식 p50/p95 현행 대비 50% 이상 단축 목표
- 정상 상태의 Firestore 조회 실패율 1% 미만 목표
- 테스트 저장 접수 응답과 시트 동기화 시간을 분리 측정

## 12. 롤백 전략

### 테스트 환경

테스트베드는 운영과 완전히 분리되므로 장애 시 `bwa_test` Pages 또는 테스트 Functions만 중지한다. 운영 서비스는 조치하지 않는다.

### 운영 읽기 전환 이후

| 장애 | 즉시 조치 | 데이터 조치 |
|---|---|---|
| Firestore 읽기 장애 | `READ_SOURCE=gas` | 없음 |
| 게시 데이터 불일치 | 해당 양식 GAS 폴백 | 운영 시트에서 재게시 |
| App Check 오탐 | Firestore enforcement 해제 | 없음 |
| Functions publisher 장애 | Scheduler/Function 중지 | GAS 읽기 유지 |
| 저장 Worker 장애 | 신규 접수 플래그 닫기 | 멱등성 키 기준 재처리 |

운영 장애 시 테스트 사이트로 사용자를 우회시키지 않는다. 테스트 사이트는 테스트 데이터만 사용하므로 운영 대체 서비스가 아니다.

## 13. 구현 체크리스트

### 테스트베드 생성 전

- [x] 운영 Git SHA 기록
- [x] 운영 GAS 최신 소스 확인
- [x] 운영 URL·Script ID·배포 ID 기록
- [x] 운영 Firebase가 FCM 전용이라는 전제 확인

### 테스트베드 격리

- [x] `bwa_test` 저장소 생성
- [x] 테스트 Pages 활성화
- [x] TEST 배너 적용
- [x] 테스트 Firebase 프로젝트 `btcwebapp-test` 생성
- [x] Blaze 연결 및 10,000원 예산 알림 설정
- [x] 테스트 Firestore `(default)` 생성 (`asia-northeast3`, 기본 거부 규칙)
- [x] 테스트 Functions 구현 및 최초 배포
- [x] Firebase 테스트 웹앱 `bwa_test-web` 등록
- [x] 테스트 FCM Web Push 공개 키 생성
- [x] 테스트 시트 복사
- [x] 테스트 시트의 FCM 토큰·로그 제거
- [x] 테스트 Drive 폴더 생성
- [x] 테스트 Apps Script 생성 및 `clasp` 분리
- [x] 테스트 GAS 최초 권한 승인 및 환경 초기화
- [x] 테스트 GAS 소유자 전용 안전 기준 배포 생성
- [x] 테스트 GAS 익명 공개 URL 생성 및 CORS 검증
- [x] `bwa_test` 테스트 GAS 읽기 전용 변경사항 원격 게시
- [x] PR `#1` squash merge 및 GitHub Pages 공개 URL 검증
- [x] T1 환경 가드 통과
- [x] T1 운영 리소스 네트워크 요청 0건 확인

### 읽기 전환

- [x] publisher 백필
- [x] GAS/Firestore 자동 비교
- [x] `gas` 모드 회귀 테스트
- [x] `shadow` 모드 검증
- [x] `firestore` 모드 검증
- [x] Firestore 오류 표시·재시도 검증
- [x] 성능 지표 기록

### 비동기 저장

- [x] 제출 스키마 검증
- [x] 멱등성 키 검증
- [x] Sheets batchUpdate 검증
- [x] 중복 이벤트 재실행 검증
- [x] 실패 재처리 검증
- [x] `synced` 후 XLSX 활성화

### 운영 승인 전

- [ ] 테스트 결과 보고서 작성
- [ ] 비용 추정
- [ ] 운영 변경 파일 목록 확정
- [ ] 운영 롤백 리허설
- [ ] 사용자 별도 승인

## 14. 후속 검토 항목

테스트베드 검증 이후 별도로 다룬다.

1. 브라우저 IndexedDB 영구 캐시
2. 기존 GAS 저장 로직의 Map 기반 최적화
3. FCM 토큰과 로그의 Firestore 이전
4. 운영 `firebase_key.js` 제거와 키 폐기
5. Firebase compat SDK에서 modular SDK로의 업그레이드
6. 운영 GAS 폴백 종료 조건

SDK 업그레이드는 Firestore 기능 이전과 동시에 하지 않는다. 기능 변경과 라이브러리 변경을 분리해 회귀 위험을 낮춘다.

## 15. 공식 참고자료

- [Cloud Firestore 오프라인 및 로컬 캐시](https://firebase.google.com/docs/firestore/manage-data/enable-offline)
- [Cloud Firestore 사용량 및 문서 크기 제한](https://firebase.google.com/docs/firestore/quotas)
- [Cloud Firestore Security Rules](https://firebase.google.com/docs/firestore/security/overview)
- [Firebase App Check 웹 설정](https://firebase.google.com/docs/app-check/web/recaptcha-provider)
- [Cloud Functions 재시도 및 멱등성](https://firebase.google.com/docs/functions/retries)
- [Google Sheets API 배치 요청](https://developers.google.com/workspace/sheets/api/guides/batch)
