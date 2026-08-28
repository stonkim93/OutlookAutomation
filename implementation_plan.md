# OutlookAutomation 오류 수정 및 안정화 계획

## 배경 및 목적

현재 OutlookAutomation 앱은 빌드는 성공하지만, 실행 시 여러 문제가 발생합니다. 파일을 특정 앱으로 열고, Outlook COM 자동화를 통해 첨부 후 지정 폴더에 저장하는 기능 전반에서 오류가 발생하는 것으로 보입니다.

---

## 발견된 주요 문제점

### 1. ❌ `App.xaml.cs`의 네임스페이스 불일치 (런타임 크래시 원인)
- `MainWindow.xaml.cs`는 `OutlookAutomationApp` 네임스페이스 사용
- `App.xaml.cs`는 `OutlookAutomation` 네임스페이스 사용
- → `App.xaml`의 `StartupUri`가 올바른 클래스를 찾지 못해 **앱 시작 시 크래시** 가능

### 2. ❌ `MainWindow.xaml.cs`에 `App` 클래스 중복 정의
- `App.xaml.cs`에 이미 `App : Application`이 있는데, `MainWindow.xaml.cs` 끝에도 동일한 클래스가 또 정의됨 (317번째 줄)
- → **컴파일 시 중복 정의 오류** 또는 예상치 못한 동작 발생

### 3. ❌ 파일 열기 후 대기 시간 부족
- 파일을 열고 `Thread.Sleep(1000)` 후 바로 Outlook 첨부 시도
- Excel/Word 같은 무거운 앱은 1초 내 완전히 열리지 않음
- → `Attachments.Add(SourceFile)` 호출 시 **파일 잠금(Lock) 오류** 발생

### 4. ❌ Outlook COM 객체 생명주기 관리 불안전
- `dynamic` 타입으로 COM 객체 사용 시 RCW (Runtime Callable Wrapper) 해제 순서가 잘못되면 오류
- `_mailItem.Close(0)` 호출 후 COM 해제 전에 Outlook 앱 종료 시도 없음
- → **COM Exception** 및 프로세스가 좀비 상태로 남는 문제

### 5. ❌ 파일을 연 프로세스를 추적/종료하는 기능 없음
- `Process.Start(psi)` 후 반환된 `Process` 객체를 저장하지 않음
- → 자동화 완료 후 열었던 앱(Excel, Word 등)을 **종료할 수 없음**

### 6. ❌ Outlook Quit 호출 없음
- `Cleanup()`에서 `_outlookApp` COM 해제 전에 `_outlookApp.Quit()` 호출이 없음
- → Outlook 프로세스가 **백그라운드에 좀비 프로세스로 남음**

### 7. ⚠️ 같은 이름 파일 저장 처리 미흡
- 동일 파일명 처리가 `_saved` 접미사 하나뿐
- 여러 번 실행 시 또 덮어씀

### 8. ⚠️ UI 스레드 안전 문제
- `Task.Run()` 내부에서 실행 중 오류 발생 시 UI 업데이트가 cross-thread 예외 일으킬 수 있음

---

## 수정 계획

### [MODIFY] MainWindow.xaml.cs
- `App` 클래스 중복 정의 **제거** (310~316줄)
- `Process` 참조 저장 (열었던 앱 추적)
- Outlook `Quit()` 추가
- 대기 시간 조정 (파일 오픈 2초, 메일 표시 3초)
- 파일 저장 시 타임스탬프 기반 고유 이름 생성
- COM 객체 해제 순서 안전하게 재정렬
- 로그 메시지 개선 (단계별 진행 상황 UI 업데이트)
- UI 스레드 안전 처리 (`Dispatcher.Invoke`)

### [MODIFY] App.xaml.cs
- 네임스페이스를 `OutlookAutomationApp`으로 통일

### [MODIFY] MainWindow.xaml
- 로그/상태 출력 영역 확장 (다단계 진행 상황 표시)
- 진행 단계별 색상 표시 개선

---

## 검증 계획

### 빌드 검증
- `dotnet build` 로 컴파일 오류 없음 확인

### 수동 검증 (사용자 확인 필요)
1. .xlsx 파일 선택 후 실행 → Excel이 열리고 → Outlook이 첨부된 메일 작성 → 지정 폴더에 저장 확인
2. .pdf 파일, .hwp 파일 등 다른 확장자도 테스트

---

## 오픈 질문

> [!IMPORTANT]
> **오류 화면 이미지**: 요청에서 "첨부 그림"을 언급하셨지만 이미지가 전달되지 않았습니다. 
> 어떤 오류 메시지가 표시되었는지 알려주시면 더 정확한 수정이 가능합니다.
> 위에서 발견한 문제들을 기반으로 수정을 진행하겠습니다.

> [!NOTE]  
> MS Office 2007 (회사 PC) 및 최신 버전 둘 다 COM ProgID `"Outlook.Application"`을 통해 접근하므로 코드 변경 없이 호환됩니다.
