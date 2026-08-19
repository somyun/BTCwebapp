# T0 운영 기준선

- 기록일: 2026-07-29
- 운영 사이트: `https://somyun.github.io/BTCwebapp/`
- 운영 Git 커밋: `2e2dc00138c4585178494556aaeea65497f20262`
- 운영 커밋 시각: 2026-07-13 14:59:56 +09:00
- 운영 GAS Script ID: `12kZx3kFBuJKIq_vjKt6EY27p0ZtprWrza4HG8ykFjrQlWVJlzyGK6_DP`
- 운영 프런트가 호출하는 GAS 배포 ID: `AKfycbzuWS4Q5kTzDRH4IBpeXBa69KngElRdArtTCzTV0NDQsB3y4oABBIzrTLuPOZH5KOPP`
- 운영 GAS 소스: 2026-07-29 `clasp pull`로 받아둔 `apps-script/` 사본을 기준으로 기록
- 운영 Firebase 프로젝트: `btcwebapp-551bd`
- Firebase 콘솔 확인 결과: Spark 요금제, Firestore DB 없음, 배포된 Functions 없음

## 읽기 성능 기준선

운영 데이터를 변경하지 않는 GET 요청만 측정했다. 아래 수치는 네트워크와 GAS 콜드 스타트 영향을 포함한 5회 표본이다.

| 작업 | 표본 | p50 | p95 | 관측값(ms) |
|---|---:|---:|---:|---|
| 양식 목록 `getFormList` | 5 | 1,725ms | 2,041ms | 1,679 / 2,041 / 1,725 / 1,725 / 2,038 |
| 선택 양식 `getFormDataForWeb` | 5 | 2,767ms | 3,166ms | 2,985 / 2,745 / 2,357 / 2,767 / 3,166 |

측정값 저장은 운영 시트 쓰기가 발생하므로 T0에서 자동 측정하지 않았다. T2 이후 테스트 시트에서 p50/p95를 측정한다.

## 운영 수동 회귀 체크리스트

- 운영 URL이 정상적으로 열린다.
- 양식 목록이 기존과 동일하게 표시된다.
- 기존 양식을 선택하면 측정 항목이 표시된다.
- 즐겨찾기와 홈 이동이 기존처럼 작동한다.
- FCM 토글과 알림 흐름이 기존처럼 유지된다.
- 운영 저장과 비상용 XLSX 준비 흐름은 T1 작업에서 변경되지 않는다.

## 복구 기준

T1은 별도 저장소와 URL만 사용한다. 문제가 발생하면 `bwa_test` Pages만 비활성화하며 운영 저장소, 운영 GAS, 운영 시트 및 운영 Firebase에는 복구 작업을 수행하지 않는다.
