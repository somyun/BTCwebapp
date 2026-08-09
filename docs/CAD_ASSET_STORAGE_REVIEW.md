# 도면 자산 비공개 저장소 검토

## 결론

도면 manifest와 레이어 JSON의 비공개 저장소로 Firestore보다 **Cloud Storage for Firebase**가 적합하다.
Firestore는 도면 버전·게시 상태·객체 경로 같은 작은 메타데이터에만 선택적으로 사용한다.

## 판단 근거

- Firestore 문서는 최대 1 MiB이며 도면 레이어가 커지면 분할 문서와 조립 로직이 필요하다.
  문서 읽기 단위의 비용 모델도 정적 파일 전달보다 불리하다.
  [Firestore quotas and limits](https://firebase.google.com/docs/firestore/quotas)
- Cloud Storage는 파일·객체 다운로드를 위한 서비스이며 Firebase Authentication을 조건으로
  Security Rules를 적용할 수 있다.
  [Download files with Cloud Storage on Web](https://firebase.google.com/docs/storage/web/download-files)
- Storage Rules에서 `request.auth`와 Custom Claims를 검사할 수 있으므로 현재 이메일 인증이
  발급하는 `humetro: true` claim으로 도면 읽기를 제한할 수 있다.
  [Storage Security Rules conditions](https://firebase.google.com/docs/storage/security/rules-conditions)
- Cloud Storage for Firebase 사용에는 Blaze 요금제가 필요하다.
  [Cloud Storage plan requirements](https://firebase.google.com/docs/storage/faqs-storage-changes-announced-sept-2024)

## 권장 구조

```text
Cloud Storage (비공개)
└─ cad/hopo/{releaseId}/manifest.json
   └─ layers/*.json

Firestore (선택 사항)
└─ cadReleases/hopo
   ├─ activeReleaseId
   ├─ manifestObjectPath
   └─ publishedAt
```

브라우저는 Firebase Auth 로그인 후 Storage SDK의 `getBytes`/`getBlob` 방식으로 객체를 읽는다.
엄격한 비공개 접근이 필요하면 장기 다운로드 토큰이 포함되는 공개형 URL을 앱 코드에 저장하지 않는다.

## 이전 전 승인·준비 항목

1. Blaze 요금제와 Storage 버킷 생성 승인
2. 운영·테스트 버킷 및 Firebase 프로젝트 완전 분리
3. `humetro == true` 읽기 조건과 관리자 전용 쓰기 규칙 검토
4. 현재 공개 자산을 버전 경로로 업로드하고 해시·레이어 수 비교
5. 인증된 Storage 로더와 실패·재시도 화면 구현
6. 배포 검증 후 공개 저장소 자산 제거

2026-08-10 일괄 전환 승인을 받아 운영·테스트 Cloud Storage 버킷, 인증 규칙,
Storage SDK 로더와 자산 업로드를 적용한다. 공개 저장소의 기존 도면 JSON은 배포 검증 뒤 제거한다.
