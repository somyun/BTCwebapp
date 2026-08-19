# 테스트 서비스 외부 리소스 삭제 안내

이 안내는 운영 앱(`BTCwebapp`, Firebase 프로젝트 `btcwebapp-551bd`)이 아닌 테스트 전용 리소스만 대상으로 합니다. 삭제 전에 아래 식별자가 맞는지 반드시 다시 확인하세요.

| 대상 | 삭제할 식별자 | 바로가기 |
| --- | --- | --- |
| GitHub 저장소 | `somyun/bwa_test` | [저장소 설정](https://github.com/somyun/bwa_test/settings) |
| Google Apps Script | 프로젝트 ID `1ynXimyiusVX7LqUqYFMW4cApQ5A0gtkOTfCx1dgkVQY-qQ5n1EoN76zv` | [프로젝트 열기](https://script.google.com/home/projects/1ynXimyiusVX7LqUqYFMW4cApQ5A0gtkOTfCx1dgkVQY-qQ5n1EoN76zv/edit) |
| Firebase | 프로젝트 ID `btcwebapp-test` | [Firebase 콘솔](https://console.firebase.google.com/project/btcwebapp-test/overview) |

## 권장 순서

1. **GitHub Pages 먼저 중지합니다.** [Pages 설정](https://github.com/somyun/bwa_test/settings/pages)에서 Build and deployment의 Source를 **None**으로 바꾸고 저장합니다. 테스트 URL이 더 이상 열리지 않는지 확인합니다.
2. **GAS 배포를 중지하고 프로젝트를 삭제합니다.** 위 Apps Script 링크에서 왼쪽 **Deployments**를 열어 각 웹 앱 배포를 보관(Archive)합니다. 이어서 **Project Settings** 맨 아래에서 **Move project to trash**를 선택하고 확인합니다. 트리거(시계 아이콘)가 남아 있지 않은지도 확인합니다.
3. **Firebase 프로젝트를 삭제합니다.** 위 Firebase 콘솔에서 톱니바퀴 **Project settings** → **General** → 맨 아래 **Delete project**를 선택합니다. 확인 입력란에는 정확히 `btcwebapp-test`를 입력합니다. `btcwebapp-551bd`는 운영 프로젝트이므로 선택하면 안 됩니다.
4. **마지막으로 GitHub 저장소를 삭제합니다.** [저장소 General 설정](https://github.com/somyun/bwa_test/settings) 하단 **Danger Zone** → **Delete this repository**를 선택하고, 확인 문구로 `somyun/bwa_test`를 입력합니다.

## 삭제 후 확인

- `https://somyun.github.io/bwa_test/`가 더 이상 제공되지 않아야 합니다.
- Firebase 콘솔의 프로젝트 선택 목록에서 `btcwebapp-test`가 삭제 예약 상태이거나 보이지 않아야 합니다. Firebase/Google Cloud 프로젝트의 완전 삭제에는 최대 30일이 걸릴 수 있습니다.
- 운영 Firebase ID `btcwebapp-551bd`와 운영 GitHub 저장소 `somyun/BTCwebapp`는 그대로 남아 있어야 합니다.

GitHub 저장소 삭제는 일부 경우 90일 안에 복구할 수 있지만, Firebase 프로젝트 ID와 번호는 삭제 후 재사용할 수 없습니다. 상세 절차는 [GitHub 공식 안내](https://docs.github.com/en/repositories/creating-and-managing-repositories/deleting-a-repository)와 [Firebase 프로젝트 안내](https://firebase.google.com/docs/projects/learn-more)를 참고하세요.
