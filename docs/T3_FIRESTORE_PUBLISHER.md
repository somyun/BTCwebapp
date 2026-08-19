# T3 Firestore 게시 캐시 생성기 기록

- 기록일: 2026-07-29
- 상태: 구현·배포·초기 게시·자동 비교 완료
- 대상 프로젝트: `btcwebapp-test` (`119870762952`)
- 리전: `asia-northeast3`
- 운영 리소스 변경: 없음
- `bwa_test` 화면 읽기 소스: 테스트 GAS 유지

## 로컬 도구

- Node.js: `22.23.1`
- npm: `10.9.8`
- Firebase CLI: `15.24.0`
- Functions SDK: `firebase-functions@7.3.2`
- Admin SDK: `firebase-admin@14.2.0`
- 런타임 의존성 감사: 알려진 취약점 0건

## 배포 리소스

다음 Node.js 22 2세대 Functions를 `btcwebapp-test`에만 배포했다.

- `publishFormList`: 수동 목록 게시
- `publishForm`: 수동 단일 양식 게시
- `publishAllChangedForms`: 수동 전체 변경 게시
- `publishAllChangedFormsScheduled`: 5분 간격 전체 변경 확인

수동 Functions는 `BWA_TEST_PUBLISHER_TOKEN` Secret의 `X-BWA-Publisher-Token` 헤더를 요구한다. Secret 버전 1과 2는 폐기했으며 버전 3만 활성화했다. 토큰 값은 소스·문서·로그에 기록하지 않았다.

Artifact Registry에는 7일이 지난 Functions 빌드 이미지를 자동 삭제하는 정책을 설정했다.

소스 게시 기록:

- 저장소: `somyun/bwa_test`
- 브랜치: `codex/t3-firestore-publisher`
- 커밋: `500385f` (`Add T3 Firestore publisher`)
- Draft PR: `#2` (`T3: add isolated Firestore publisher`)

## Firestore 권한

- `publicCache/{documentId}`: 익명 읽기 허용, 직접 쓰기 거부
- `publicForms/{formKey}`: 익명 읽기 허용, 직접 쓰기 거부
- `publicForms/{formKey}/chunks/{chunkId}`: 익명 읽기 허용, 직접 쓰기 거부
- 그 밖의 모든 경로: 읽기·쓰기 거부
- 익명 직접 쓰기 검증: HTTP `403`

`items`와 `rows` 필드는 인덱스에서 제외했다. 작은 양식은 단일 문서에 저장하고, 문서 크기가 기준을 넘으면 `chunks` 하위 컬렉션으로 분할하도록 구현했다.

## 초기 게시 및 데이터 일치

- `publicCache/formList`: 6개
- 목록 content hash: `0ca51cafc948c837a5959de2dd7090b11bc1566c5d5761de4c2720204f6cc373`

| 양식 | 행 수 | 결과 |
|---|---:|---|
| 율리24 | 33 | 일치 |
| 호포24 | 37 | 일치 |
| 호포154 | 21 | 일치 |
| 덕포 변전소 | 32 | 일치 |
| 덕천변전소 | 36 | 일치 |
| 구명154 | 36 | 일치 |

`scripts/compare-published.js`가 테스트 GAS와 공개 Firestore 문서를 목록 hash, 양식 hash, 전체 행 데이터 기준으로 비교해 모두 통과했다.

스케줄러 로그:

- 최초 실행: 6개 양식 게시, 약 11.0초
- 다음 실행: 변경 양식 0개, 약 9.6초
- orphan 삭제: 0개

## 격리 결과

- Functions 메타데이터의 프로젝트: `btcwebapp-test`
- 운영 GAS Script ID 참조: 0건
- 운영 GAS 배포 ID 참조: 0건
- 운영 Firebase 프로젝트 ID 참조: 0건
- 운영 GitHub Pages와 운영 GAS·시트·Firebase 변경: 없음
