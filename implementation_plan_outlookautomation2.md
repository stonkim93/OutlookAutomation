# Drag & Drop 및 파일명 자동 변환 구현 계획

요청하신 드래그 앤 드롭 기능과 대상 파일명 자동 변경(원본 보존 및 우회) 기능을 구현하기 위한 계획입니다.

## 🛠 구현 내용

1. **이전 코드 정리 (MAPI 잔재 제거)**
   - 이전 작업에서 일부 남아있던 MAPI 관련 불필요한 코드 조각(구조체 선언 등)을 완전히 제거하여 코드를 깔끔하게 정리합니다.

2. **Drag & Drop UI 지원 추가**
   - `MainWindow.xaml`의 `<Window>` 태그에 `AllowDrop="True"` 및 `Drop="Window_Drop"` 이벤트를 추가하여 파일을 끌어다 놓을 수 있도록 만듭니다.

3. **Window_Drop 이벤트 처리 및 자동 변환 로직**
   - 마우스로 파일을 드롭했을 때, 파일이 `_drm`으로 끝나지 않으면 다음과 같이 동작합니다:
     1. 원본 파일(예: `test.xlsx`)의 이름을 `test_drm.xlsx`로 변경(이동)합니다.
     2. 앱의 입력 소스(`SourceFile`)를 `test_drm.xlsx`로 설정합니다.
     3. 결과물 저장 대상 파일명(`TargetFileName`)을 원본 이름인 `test.xlsx`로 설정합니다.
     4. 출력 폴더(`OutputFolder`)를 원본 파일이 있는 폴더로 자동 지정합니다.
     5. 설정이 끝나면 자동으로 **자동화 실행(`BtnExecute_Click`)**을 트리거합니다.

4. **Service 로직 업데이트**
   - `OutlookAutomationService`에 `TargetFileName` 속성을 추가합니다.
   - `SaveAttachmentViaCOM` 메서드에서 파일을 추출해 저장할 때, 기존에는 첨부파일 이름(`test_drm.xlsx`)을 그대로 썼으나, 이제는 `TargetFileName`(`test.xlsx`)이 지정되어 있다면 해당 이름으로 C# File I/O를 통해 복호화된 바이트를 저장하도록 변경합니다.

---

> [!IMPORTANT]
> **사용자 확인 요청 사항 (Open Questions)**
> 1. 원본 파일을 `test_drm.xlsx`로 "이름 변경(Move)" 하는 방식을 사용하겠습니다. 만약 기존 파일명 보존을 위해 "복사(Copy)" 후 복사본을 `test_drm.xlsx`로 만들어야 한다면 알려주세요. (일반적으로 원본이 덮어씌워지지 않게 하려면 원본을 복사하는 것이 더 안전할 수 있습니다.)
> 2. 드래그 앤 드롭 시 추가적인 확인 버튼 없이 즉시 "자동화 실행"이 시작되도록 구현해도 괜찮을까요?

위 계획에 대해 승인(Proceed)해주시거나, 동작 방식(복사 vs 이동)에 대한 피드백을 주시면 즉시 작업을 진행하겠습니다!
