# 테스트 코드 운영 적용 계획

작성 기준일: 2026-08-11 (Asia/Seoul)

## 1. 문서 목적

안정화된 `bwa_test`의 기능을 운영 `BTCwebapp`에 적용하기 위한 작업 범위, 환경별 치환값, 단계별 배포 순서, 검증 기준과 복구 방법을 정리한다.

이 문서는 후속 작업의 인수인계 자료다. 문서 작성 시점에는 운영 적용 코드를 수정하거나 배포하지 않았다.

## 2. 기준 상태

### 운영 저장소

- 경로: `C:\Users\bysub\Documents\btcwebapp`
- 저장소: `somyun/BTCwebapp`
- 현재 커밋: `076857c` (`8d8c880` 전환 커밋을 되돌린 복구 커밋)
- 현재 파일 트리: 복구 전 운영 버전 `35a13d1`과 동일
- 배포 URL: `https://somyun.github.io/BTCwebapp/`
- 현재 배포된 Firebase Function:
  - `publishAllChangedFormsScheduled`
- 현재 저장 방식: 브라우저에서 운영 GAS의 `saveMeasurementsToSheet` 직접 호출
- 현재 알림 방식: 브라우저가 FCM 토큰을 GAS에 등록하고 GAS가 알림 발송

### 테스트 저장소

- 경로: `C:\Users\bysub\Documents\btcwebapp\testbed\bwa_test_publish`
- 저장소: `somyun/bwa_test`
- 현재 커밋: `26e56c1`
- 배포 URL: `https://somyun.github.io/bwa_test/`
- 실제 배포된 Firebase Functions: 16개
  - 양식 게시
  - 측정값 비동기 접수·조회·동기화·재처리·게이트
  - 알림 기기 등록·상태·활성화·수신 확인·자체 테스트
  - 알림 heartbeat 및 해피휴게더 예약 발송

## 3. 핵심 적용 원칙

테스트 저장소 전체를 운영에 덮어쓰지 않는다. 테스트에서 검증된 기능을 운영 구조에 맞춰 다음 단위로 분리 이식한다.

1. 화면 상태 통합(기존 뒤로가기 동작은 유지·회귀 검증)
2. 양식 조회 실패 처리
3. Functions 비동기 저장
4. 알림 설정·내역 UI
5. Firebase 알림 백엔드
6. GAS 비공개 게시글 브리지

각 기능은 독립 커밋과 독립 활성화 게이트를 사용한다. 신규 백엔드를 먼저 비활성 상태로 배포하고 검증한 뒤 운영 웹의 호출 경로를 전환한다.

## 4. 현재 운영과 테스트의 차이

| 구분 | 현재 운영앱 | 안정화된 테스트앱 |
|---|---|---|
| 양식 조회 | Firestore 우선, 실패 시 GAS 자동 대체 | Firestore 실패 시 오류·재시도 표시 |
| 측정값 저장 | 브라우저가 GAS에 직접 저장 | Functions 접수 → Firestore → Sheets API 저장 |
| XLSX | GAS 저장 직후 준비 | Sheets 동기화 완료 후 준비 |
| 화면 전환 | 요소별 표시·숨김 로직이 분산됨 | 홈 전용 요소 목록을 한곳에서 관리 |
| 안드로이드 뒤로가기 | 이미 폼에서 홈으로 복귀 | 동일 동작(신규 변경 없음) |
| 알림 설정 | 메인·사이드 토글, GAS에 FCM 토큰 저장 | 별도 설정 페이지와 Firebase Functions 사용 |
| 알림 내역 | 별도 내역 없음 | IndexedDB 기반 최근 알림 내역 |
| 알림 발송 | GAS 시간 트리거 | Firebase 예약 Function |
| 배포된 Functions | 예약 양식 게시 1개 | 조회·저장·알림 등 16개 |
| 로컬 설정 | 접두사 없는 운영 키 | `bwa_test:` 접두사로 운영과 분리 |

## 5. 운영 웹 변경 범위

### `index.html`

- 홈 전용 요소에 공통 표시 상태 적용
- 헤더에 알림 내역 버튼 추가
- 사이드 메뉴의 기존 알림 토글을 알림 설정 링크로 교체
- 알림 설정·내역 스크립트 연결
- 운영 Functions와 Firebase Messaging에 필요한 CSP 주소 추가
- 정적 파일 캐시 버전 갱신

### `script.js`

다음 홈 전용 요소를 하나의 목록으로 관리한다.

- `favoritesSection`
- `formMessage`
- `mainToggleContainer`
- `openMapBtn`

적용할 동작:

- 양식 진입 시 홈 전용 요소 일괄 숨김
- 홈 복귀 시 일괄 복원
- 진행 중 양식 요청의 응답 무효화
- 기존 안드로이드 뒤로가기(폼 → 홈) 동작 유지 및 회귀 테스트
- 목록·양식 조회 오류를 오류·재시도 화면으로 표시
- 저장 버튼 중복 클릭 방지
- Functions 제출 생명주기 상태는 내부 상태로 관리
  - 화면에 별도의 상시 상태 행을 추가하지 않음
  - 사용자가 보는 진행 정보는 저장 버튼 문구와 기존 상단 상태 메시지로 표시
  - 내부 상태값은 중복 저장 방지, 재조회, XLSX 활성화 판단에만 사용
- 네트워크 응답이 불확실하면 같은 멱등성 키의 상태를 먼저 조회
- Sheets 동기화가 끝난 후에만 XLSX 준비

기존 운영의 즐겨찾기·간격·정렬 저장 키는 변경하지 않는다.

### `production-read.js`

- `gasFallback: true` 제거
- Firestore 실패 시 GAS 자동 대체 금지
- 운영과 로컬 모두 Firestore 단일 조회 경로 사용
- `gas`, `shadow` 모드와 관련 비교·상태·분기 코드 삭제
- `readSource` URL 파라미터 해석 코드 전체 삭제
- 빈 양식, 잘못된 `formKey`, 오래된 revision을 오류로 처리

### `style.css`

- `.home-only-hidden`
  - 양식 화면에서만 홈 전용 요소를 `display: none !important`로 레이아웃에서도 제거
  - 즐겨찾기, 홈 안내문, 기존 알림 영역, 도면 버튼에 동일한 규칙 적용
- 조회 오류 화면
  - 오류 문구의 여백·색상·줄바꿈을 통일
  - 목록 오류는 좁은 영역, 폼 오류는 전체 폼 영역에 맞게 배치
- `.read-retry-button`
  - 폼 오류용 기본 크기와 목록 오류용 `compact` 크기를 구분
  - 터치 가능한 여백과 기존 버튼 디자인 유지
- 제출 상태
  - 테스트 코드의 `.submission-status`는 현재 `display: none`이므로 운영에도 별도 상태 행을 노출하지 않음
  - 저장 버튼의 `저장 중`, `저장 완료`, `다시 저장` 문구와 기존 상단 상태 영역만 사용
- `.notification-history-btn`
  - 헤더 알림 내역 버튼의 크기, 정렬, 기본 색상 지정
- `.notification-off` 및 `.notification-off-mark`
  - 알림 OFF일 때 아이콘을 회색으로 바꾸고 빨간 사선을 표시
- `.notification-settings-link`
  - 사이드 메뉴의 알림 설정 링크를 버튼 형태로 표시하고 hover·키보드 focus 상태 제공
- 알림 설정·내역 전용 레이아웃은 `style.css`가 아니라 새 `notifications.css`에서 관리

## 6. 측정값 비동기 저장 백엔드

### 변경 파일

- `firebase-production/functions/index.js`
- `firebase-production/functions/package.json`
- `firebase-production/functions/package-lock.json`
- `firebase-production/functions/lib/publisher.js`
- `firebase-production/functions/test/helpers/fake-firestore.js`
- `firebase-production/firestore.rules`
- `firebase-production/firestore.indexes.json`
- `firebase-production/firebase.json`

### 추가 모듈

- `functions/lib/submission.js`
  - 요청 크기·형식 검증
  - 멱등성 키 검증
  - 중복·충돌 방지
  - 제출 게이트 및 요청 빈도 제한
- `functions/lib/sheets-sync.js`
  - Sheets API `values.batchUpdate`
  - 운영 시트 F열과 FormList revision 갱신
- `functions/lib/synchronizer.js`
  - Firestore 생성 이벤트 처리
  - 중복 이벤트 방지
  - 재시도 가능 오류 분류

### 추가 Functions

- `submitMeasurements`
- `getMeasurementSubmission`
- `syncMeasurementSubmission`
- `retryMeasurementSubmission`
- `setSubmissionGate`
- 관리자용 양식 게시 endpoint

### 운영 환경 치환값

- Firebase 프로젝트: `btcwebapp-551bd`
- 스프레드시트: `19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s`
- 운영 GAS URL 사용
- 관리자 Secret: `BWA_PUBLISHER_TOKEN`
- 허용 Origin: `https://somyun.github.io`
- Functions 리전: `asia-northeast3`

### Firestore 서버 전용 컬렉션

- `measurementSubmissions`
- `submissionRateLimits`
- `systemConfig`

브라우저 직접 읽기·쓰기는 모두 거부하고 Functions의 Admin SDK만 접근한다.

### 필요한 Firebase·Google Cloud 설정

- Google Sheets API 활성화
- 운영 Functions 서비스 계정에 운영 시트 편집 권한 부여
- `BWA_PUBLISHER_TOKEN` Secret 생성
- Functions·Firestore Rules·Indexes 배포
- 최초 배포 시 제출 게이트는 반드시 OFF

## 7. 알림 웹 변경 범위

### 추가 파일

- `notification-settings.html`
- `notification-settings.js`
- `notifications.html`
- `notifications.js`
- `notifications.css`
- `notification-store.js`

### 제공 기능

- 알림 전체 켜기·끄기
- 키워드 자동 저장
- 브라우저 권한 상태와 복구 안내
- iPhone/PWA 설치 안내
- 최신 해피휴게더 글을 이용한 실제 알림 테스트
- 마지막 수신·표시·클릭 상태
- 최근 알림 내역 최대 100건 로컬 보관
- 알림 OFF 시 UI를 먼저 전환하고 서버 동기화는 백그라운드 처리

### `firebase-messaging-sw.js`

파일명과 운영 scope는 유지하고 내부 구현만 교체한다.

- 알림을 IndexedDB에 기록
- 수신·표시·클릭 확인을 Functions에 전송
- heartbeat는 사용자에게 표시하지 않고 상태만 기록
- 알림 클릭 시 정확한 URL 열기
- 잘못된 scope에서 실행되면 등록 해제
- 현재 잘못된 기본 경로 `btc_webapp`을 `BTCwebapp`으로 수정

운영 IndexedDB 이름은 테스트의 `bwa-test-notifications-v1`을 복사하지 않고 별도 운영 이름을 사용한다.

## 8. 알림 Functions 변경 범위

### 추가 모듈

- `functions/lib/notification-service.js`
  - 기기 Secret 해시 저장
  - FCM 토큰 교체
  - 키워드 매칭
  - 결정적 이벤트 ID를 이용한 중복 방지
  - 폐기된 FCM 토큰 자동 비활성화
- `functions/lib/humetro-client.js`
  - GAS 비공개 게시글 조회
  - 허용된 해피휴게더 호스트 검증
  - 게시글 정렬·중복 제거

### 추가 Functions

- `registerNotificationDevice`
- `setNotificationDeviceActive`
- `getNotificationDeviceStatus`
- `acknowledgeNotification`
- `sendNotificationSelfTest`
- `sendNotificationHeartbeatScheduled`
- `sendHappyHugetherNotificationsScheduled`

### Firestore 서버 전용 데이터

- `notificationDevices/{deviceId}`
- `notificationDevices/{deviceId}/receipts/{eventId}`
- `systemConfig/happyHugetherNotifications`

### 운영 전용 보완

테스트 구현에는 예약 알림 전체를 운영자가 끄는 명시적 게이트가 부족하다. 운영 이식 시 다음을 추가한다.

- `notificationDispatch.enabled` 기본값 `false`
- 예약 발송 Function은 게이트가 `true`일 때만 실제 발송
- self-test와 heartbeat는 예약 게시글 발송 게이트와 분리

## 9. GAS 변경 계획과 승인 경계

전체 알림 구조를 테스트와 동일하게 이전하려면 운영 GAS에 최소 변경이 필요하다.

현재 운영 GAS에는 테스트 Functions가 사용하는 비공개 게시글 브리지 action이 없다.

### 예상 변경

- `apps-script/Code.js`
  - Functions 전용 최신 게시글 조회 POST action
  - Functions 전용 최근 게시글 목록 조회 POST action
- `apps-script/hugether.js`
  - 최신 게시글 정규화
  - 누락 없는 최근 게시글 목록 조회
  - `HUMETRO_BRIDGE_TOKEN` 검증
- Apps Script Script Property
  - `HUMETRO_BRIDGE_TOKEN`

### 주의사항

`hugether.js`를 포함한 운영 GAS 전체 소스 일부는 현재 Git에서 추적되지 않는다. GAS 변경 전 다음 절차가 필수다.

1. 현재 배포된 운영 GAS를 별도 작업 폴더로 가져온다.
2. 로컬 미추적 파일과 배포본을 비교한다.
3. 변경할 두 action과 토큰 검증 부분만 별도 diff로 만든다.
4. 사용자가 GAS 변경을 직접 검토하고 승인한다.
5. 승인 후에만 GAS 배포와 Script Property 설정을 수행한다.

최종 알림 전환 시 기존 GAS 알림 시간 트리거를 중지해야 중복 알림을 방지할 수 있다. GAS 소스 변경과 트리거 중지는 웹·Firebase 변경과 별도 승인으로 처리한다.

## 10. 운영에 적용하지 않을 테스트 전용 코드

- `bwa_test:` 로컬 저장소 접두사
- `btcwebapp-test` 프로젝트 ID·API 키·Functions URL
- 테스트 GAS URL과 테스트 스프레드시트 ID
- `/bwa_test/` Service Worker scope
- `TEST` 배지
- 테스트 네트워크 차단기
- 로컬 오류 시뮬레이션 옵션
- `BWA_TEST_PUBLISHER_TOKEN`
- 테스트앱에서 운영 정적 파일을 원격 참조하는 구조

## 11. 도면 기능 제외

초기 운영 적용에서 도면 Storage 전환은 제외한다.

현재 테스트 HTML은 운영의 `cad-storage.js`를 참조하지만 운영 복구로 이 파일은 저장소에서 제거됐다. 테스트의 알림 안정성과 도면 Storage 경로의 안정성은 별개다.

다음은 그대로 유지한다.

- 기존 `map.js`
- 기존 `cad-data/hopo`
- 기존 이메일 인증
- 현재 공개 도면 로딩 방식

도면 Storage 재전환은 별도 작업으로 계획·검증한다.

## 12. 단계별 실행 계획

### 1단계 — 기준선과 기능 플래그

- 운영 `main`에서 별도 작업 브랜치 생성
- 현재 운영 웹·GAS·Firebase 설정과 Functions 목록 저장
- `submission.enabled = false`
- `notificationDispatch.enabled = false`
- 기존 GAS 저장·알림 경로 유지

### 2단계 — 화면 전환 로직

- 홈 전용 요소 통합
- 도면 버튼 잔존 문제 수정
- 조회 오류·재시도 화면
- 기존 안드로이드 뒤로가기 동작 회귀 검증(로직 변경 없음)
- 로컬 테스트와 `bwa_test` 회귀 테스트만 수행

### 3단계 — Functions 비활성 배포

- 비동기 저장 Functions
- 알림 기기 관리 Functions
- Firestore Rules·Indexes
- Secret과 IAM
- Sheets API 권한

운영 웹은 아직 기존 GAS 경로를 사용한다.

### 4단계 — 제한된 측정값 저장 시험

- 제출 게이트를 특정 양식 allowlist 방식으로 확장
- 검토용 운영 양식 한 개만 허용
- 실제 시트 셀·revision 확인
- 같은 요청 재전송 시 한 번만 저장되는지 확인
- Sheets 반영 후 XLSX 준비 확인
- 실패 시 게이트 OFF, 기존 GAS 저장 유지

### 5단계 — 알림 기기 이관

- 운영 알림 설정 페이지 배포
- 페이지 방문 시 기존 FCM 토큰을 새 Functions에 재등록
- 기존 GAS의 `FCM_Tokens` 시트는 삭제하지 않음
- self-test와 heartbeat만 사용
- 예약 게시글 알림은 계속 OFF

### 6단계 — GAS 브리지 별도 승인

- 실제 배포 GAS와 로컬 파일 비교
- 변경 diff 제출
- 승인 후 브리지와 Script Property 배포
- Firebase Secret과 GAS Script Property에 동일한 임의 토큰 설정
- 누락·오인증·허용되지 않은 링크 차단 검증

### 7단계 — 알림 발송 전환

- 현재 최신 게시글을 기준값으로만 기록
- 새 게시글 한 건으로 키워드·대상·중복 방지 확인
- 기존 GAS 알림 시간 트리거 중지
- Firebase 예약 알림 활성화
- 기존 GAS 코드와 토큰 시트는 복구용으로 일정 기간 보존

### 8단계 — 운영 웹 최종 전환

- Firestore 실패 시 오류·재시도
- Functions 비동기 저장
- 새 알림 설정·내역
- Service Worker 갱신
- 정적 파일 캐시 버전 갱신

화면·저장·알림을 각각 독립 커밋으로 배포한다.

## 13. 필수 검증 기준

### 화면·조회

- 양식 선택 후 즐겨찾기·안내·알림·도면 버튼이 모두 사라짐
- 홈 복귀 시 모두 복원
- 뒤로가기 시 앱 종료가 아니라 홈 복귀
- Firestore 장애 시 GAS 대체 없이 오류·재시도 표시
- 오래된 요청의 응답이 새 화면을 덮어쓰지 않음

### 저장

- 저장 버튼 중복 실행 방지
- 같은 멱등성 키는 Sheets에 한 번만 기록
- 잘못된 revision·행 identity 거부
- Sheets 반영 전 XLSX 버튼 비활성화
- Sheets 반영 후 XLSX 준비
- 실패 시 기존 운영 GAS로 자동 재전송하지 않음

### 알림

- 알림 토글 UI 즉시 반응
- 기존 사용자의 FCM 토큰 재등록
- 백그라운드 수신·표시·클릭 확인
- 알림 내역 저장
- 키워드별 발송
- 게시글 중복 발송 방지
- 폐기된 FCM 토큰 자동 비활성화
- 기존 GAS와 Firebase에서 중복 알림이 발생하지 않음

### 환경 격리

- 테스트 프로젝트 ID·시트 ID·Secret이 운영 코드에 남지 않음
- 운영 즐겨찾기·정렬·간격·알림 설정 유지
- `bwa_test:` 키가 운영에서 사용되지 않음
- 도면 파일과 로더는 변경되지 않음

## 14. 복구 전략

- 웹 전환 전까지 기존 GAS 저장·알림 경로 유지
- 제출 게이트 OFF로 신규 저장 즉시 차단 가능
- 알림 발송 게이트 OFF로 예약 발송 즉시 차단 가능
- 기존 GAS 알림 트리거를 일정 기간 삭제하지 않고 비활성 상태로 보존
- 기존 `FCM_Tokens` 시트를 삭제하지 않음
- 화면·저장·알림을 독립 커밋으로 만들어 기능별 revert 가능
- Firestore의 신규 서버 전용 컬렉션은 복구 시 삭제하지 않고 접근만 차단
- 도면 기능은 변경 대상에서 제외하므로 별도 복구 불필요

## 15. 별도 승인이 필요한 지점

다음 두 지점에서는 작업을 멈추고 사용자 승인을 받아야 한다.

1. 운영 제출 게이트를 열어 실제 운영 Google Sheet에 Functions 저장을 허용할 때
2. 운영 GAS 브리지 추가 및 기존 GAS 알림 시간 트리거 중지를 수행할 때

이 두 승인 전에는 운영 데이터 저장 방식과 운영 알림 발송 주체를 변경하지 않는다.
