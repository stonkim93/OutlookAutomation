# 🚀 Outlook 파일 첨부 자동화 앱 - 설치 및 실행 가이드

## 📋 빠른 시작

### Python 버전 (빠른 테스트)

```bash
# 1. Python 설치 (3.8+)
#    https://www.python.org/downloads/

# 2. pywin32 설치
pip install pywin32

# 3. 스크립트 실행
python outlook_automation_python.py
```

### C# .NET 버전 (프로덕션)

```bash
# Visual Studio 2022에서 프로젝트 생성 후 코드 복붙
# 또는 .NET CLI 사용:

dotnet new wpf -n OutlookAutomationApp
# 코드 복붙 후:
dotnet run
```

---

## 🔧 Python 버전 - 상세 설치 가이드

### Step 1: Python 설치

```bash
# Windows에서:
1. https://www.python.org/downloads/ 방문
2. "Download Python 3.11" 클릭
3. 설치 프로그램 실행
4. ✅ "Add python.exe to PATH" 체크
5. Install Now 클릭

# 설치 확인:
python --version
# Python 3.11.x
```

### Step 2: pywin32 설치

```bash
# 명령 프롬프트(cmd) 또는 PowerShell에서:

pip install pywin32

# 설치 확인:
python -c "import win32com.client; print('✅ pywin32 설치됨')"
```

### Step 3: Outlook 설정

```
Outlook COM 자동화 보안 설정:

1. Outlook 열기
2. 파일 → 옵션 → 트러스트 센터
3. 트러스트 센터 설정 클릭
4. 프로그래머 액세스 → "신뢰할 수 있는 프로그래머 액세스를 허용" 체크

주의: Outlook이 실행 중이어야 스크립트 작동
```

### Step 4: 스크립트 실행

```bash
# outlook_automation_python.py 저장 후:
python outlook_automation_python.py

# 또는 직접 실행 (더블클릭):
# outlook_automation_python.py를 Windows 탐색기에서 더블클릭
# (Python이 PATH에 등록되어 있어야 함)
```

### 예상 동작

```
╔════════════════════════════════════════════════════════════╗
║ 🔗 Outlook 파일 첨부 자동화 시작                           ║
╚════════════════════════════════════════════════════════════╝

[Step 1] 파일 선택
==================================================
파일 선택 대화상자 표시
✅ 선택된 파일: C:\Users\user\Documents\report.xlsx
   파일 크기: 125.34 KB

[Step 2] 출력 폴더 선택
==================================================
폴더 선택 대화상자 표시
✅ 선택된 폴더: C:\Output

[Step 3] 파일을 기본 앱으로 열기
==================================================
✅ report.xlsx을 Excel(로)로 열었습니다.

[Step 4] Outlook 초기화
==================================================
✅ 기존 Outlook 인스턴스 사용

[Step 5] 이메일 메시지 생성 및 파일 첨부
==================================================
✅ 메일에 파일 첨부: report.xlsx
   첨부 파일 수: 1
✅ Outlook에서 메시지 표시

[Step 6] 첨부 파일 저장
==================================================
⚠️  사용자 확인 필요
Outlook 창이 열렸습니다.
메일의 첨부 파일을 확인하시고,
콘솔에서 엔터를 눌러 저장을 진행하세요.

[엔터 키 누르기]

✅ 첨부 파일이 저장되었습니다.
   저장 경로: C:\Output\report.xlsx
   파일 크기: 125.34 KB

════════════════════════════════════════════════════════════
✅ 모든 작업이 완료되었습니다!
════════════════════════════════════════════════════════════
저장 위치: C:\Output
```

---

## 🔧 C# .NET 버전 - 상세 설치 가이드

### Step 1: Visual Studio 2022 설치

```
1. https://visualstudio.microsoft.com/downloads/ 방문
2. "Visual Studio 2022 Community" 다운로드
3. 설치 시 다음 워크로드 선택:
   ✅ .NET desktop development
   ✅ Windows desktop development with C++

4. 설치 완료 후 실행
```

### Step 2: 프로젝트 생성

```bash
# Option A: Visual Studio GUI 사용
1. Visual Studio 열기
2. "Create a new project"
3. "WPF App (.NET Framework)" 또는 "WPF App (.NET 6)" 선택
4. 프로젝트 이름: OutlookAutomationApp
5. Create

# Option B: .NET CLI 사용
dotnet new wpf -n OutlookAutomationApp
cd OutlookAutomationApp
```

### Step 3: COM Interop 설정 (중요!)

#### .NET Framework 4.7.2 사용 시

```xml
<!-- .csproj 파일에 다음 추가 -->
<ItemGroup>
    <Reference Include="Microsoft.Office.Interop.Outlook">
        <HintPath>C:\Program Files\Microsoft Office\Office16\OUTLOOK.OLB</HintPath>
        <Embed Interop Types="False" />
    </Reference>
</ItemGroup>
```

#### .NET 6.0 이상 사용 시

```xml
<!-- .csproj 파일에 다음 추가 -->
<ItemGroup>
    <COMReference Include="Microsoft.Office.Interop.Outlook">
        <Guid>00062FFF-0000-0000-C000-000000000046</Guid>
        <VersionMajor>9</VersionMajor>
        <VersionMinor>8</VersionMinor>
        <Lcid>0</Lcid>
        <WrapperTool>primary</WrapperTool>
        <Isolated>False</Isolated>
    </COMReference>
</ItemGroup>
```

### Step 4: 코드 추가

```
1. MainWindow.xaml.cs 파일 열기
2. 제공된 C# 코드 전체 복붙
3. MainWindow.xaml 파일 열기
4. 주석의 XAML 코드 복붙
5. App.xaml.cs 업데이트

또는:

프로젝트 우클릭 → Add Reference → COM
"Microsoft Outlook Object Library" 선택
```

### Step 5: 빌드 및 실행

```bash
# 빌드
dotnet build

# 실행
dotnet run

# 또는 Visual Studio에서:
# F5 키 또는 Debug → Start Debugging
```

### 예상 UI

```
╔─────────────────────────────────────────────────────────╗
│ Outlook 파일 첨부 자동화                                  │
╠─────────────────────────────────────────────────────────╣
│                                                          │
│  Outlook 파일 첨부 자동화                                 │
│                                                          │
│  [1단계: 파일 선택]                                       │
│  [파일을 선택하세요] [파일 선택...]                       │
│                                                          │
│  [2단계: 출력 폴더 선택]                                  │
│  [폴더를 선택하세요] [폴더 선택...]                       │
│                                                          │
│  이 프로그램은 다음을 수행합니다:                         │
│  • 파일을 기본 앱으로 엽니다 (Excel, Word 등)           │
│  • Outlook을 열고 파일을 첨부합니다                      │
│  • 첨부 파일을 지정된 폴더에 저장합니다                  │
│                                                          │
│  상태: 준비 완료                                          │
│  [진행 표시줄]                                            │
│                                                          │
│  [자동화 실행]                                            │
│                                                          │
╚─────────────────────────────────────────────────────────╝
```

---

## 🐛 문제 해결

### Python: "ModuleNotFoundError: No module named 'win32com'"

```bash
# 해결책:
pip install pywin32 --upgrade

# 또는 사용자별 설치:
pip install --user pywin32
```

### Python/C#: "The specified module could not be found"

```
원인: Outlook이 설치되지 않음 또는 COM 라이브러리 미등록

해결책:
1. Outlook 설치 확인
   - Microsoft 365 또는 Office 2019 이상 필요
   - Microsoft Outlook 독립 앱 필요

2. COM 라이브러리 재등록 (C#):
   - Visual Studio → Tools → Import and Export Settings
   - Reset environment 클릭
```

### C#: "No Office found in the system"

```
해결책:
1. Office 재설치:
   - 제어판 → 프로그램 → 프로그램 제거
   - Microsoft Office 찾아 제거
   - Office 재설치

2. COM 참조 수동 추가:
   - 프로젝트 우클릭 → Add → Reference
   - COM → "Microsoft Outlook Object Library" 선택
```

### Outlook 보안 경고: "Access to Outlook is blocked"

```
해결책:
1. 첫 실행 시 사용자가 수동으로 승인
   - Outlook에서 "Allow access" 클릭

2. Outlook 보안 설정 변경:
   - Outlook 열기
   - File → Options → Trust Center
   - Trust Center Settings → Programmatic Access
   - "Trust all installed add-ins and macros" 체크

3. 또는 PowerShell로 실행:
   - PowerShell을 관리자로 실행
   - python outlook_automation_python.py 실행
```

---

## 📊 성능 및 제약사항

### Python 버전

```
장점:
  ✅ 빠른 개발/테스트
  ✅ 간단한 코드
  ✅ 스크립트 수정 용이

단점:
  ❌ Python 설치 필수 (~50MB)
  ❌ pywin32 설치 필수
  ❌ 배포 복잡 (pyinstaller 필요)
  ❌ 성능: 100ms 정도 오버헤드

권장 용도:
  - 개인용 자동화
  - 빠른 프로토타입
  - 스크립트 단계
```

### C# .NET 버전

```
장점:
  ✅ .exe 파일 하나만 필요
  ✅ 뛰어난 성능 (컴파일 언어)
  ✅ 전문적인 UI
  ✅ 배포 간단

단점:
  ❌ 개발 시간 더 필요
  ❌ .NET Runtime 필요 (또는 Self-contained 배포)
  ❌ 코드 복잡도 높음

권장 용도:
  - 프로덕션 환경
  - 사용자 배포
  - 엔터프라이즈 용도
  - MS Store 등록
```

---

## 📦 배포 방법

### Python 스크립트 배포

```bash
# 1. 단순 배포 (Python 설치 필요)
# 스크립트 파일만 배포
outlook_automation_python.py

# 2. .exe로 변환 (Python 불필요)
pip install pyinstaller
pyinstaller --onefile outlook_automation_python.py

# 배포 파일:
# dist/outlook_automation_python.exe
```

### C# 애플리케이션 배포

```bash
# 1. Self-contained 배포 (런타임 포함)
dotnet publish -c Release -r win-x64 --self-contained

# 2. Framework-dependent 배포 (런타임 별도)
dotnet publish -c Release

# 3. MSI 설치 파일 생성
# Visual Studio → Tools → Create Installer
# 또는 WiX Toolset 사용
```

---

## 🔐 보안 고려사항

### Outlook 자동화 보안

```
⚠️  주의사항:

1. Outlook 보안 경고
   - 첫 실행 시 사용자 승인 필요
   - 이메일 자동 전송 기능은 인증 필요

2. 트로이목마 위험
   - 신뢰할 수 있는 출처에서만 실행
   - 코드 검증 필수

3. 파일 접근 권한
   - 지정된 폴더만 접근
   - 파일 덮어쓰기 방지 (다른 이름 저장)

4. Outlook 데이터 보호
   - 첨부 파일만 추출 (이메일 내용 불변)
   - 메일 저장 안 함 (자동 삭제)
```

---

## ✅ 체크리스트

### 실행 전 확인

```
□ Outlook이 설치되어 있음
□ Windows 11 (또는 Windows 10 1909 이상)
□ 관리자 권한으로 실행
□ 첨부 파일 폴더에 쓰기 권한
□ Outlook이 실행 중이거나 곧 실행 가능

Python 버전:
  □ Python 3.8+ 설치
  □ pywin32 설치

C# .NET 버전:
  □ Visual Studio 2022 설치
  □ .NET 6.0+ 또는 .NET Framework 4.7.2+
  □ COM Interop 설정 완료
```

---

## 🎓 추가 기능 구현

### 자동 이메일 전송

```python
# outlook_automation_python.py 수정

# send_mail 메서드 추가
def send_mail(self, recipient_email):
    """메일 자동 전송"""
    self.mail_item.To = recipient_email
    self.mail_item.Send()  # 자동 전송
    print(f"✅ 이메일 전송됨: {recipient_email}")
```

### 여러 파일 일괄 처리

```python
# 수정된 main 함수
files = [
    "C:\\file1.xlsx",
    "C:\\file2.xlsx",
    "C:\\file3.xlsx",
]

for file_path in files:
    app.source_file = file_path
    app.run()
```

### 스케줄링 (자동 실행)

```python
# schedule 라이브러리 사용
pip install schedule

# 매일 오전 9시 실행
import schedule
schedule.every().day.at("09:00").do(app.run)

while True:
    schedule.run_pending()
    time.sleep(60)
```

---

## 📞 지원

```
문제가 발생한 경우:

1. 콘솔 로그 확인
   - 정확한 오류 메시지 복사

2. Outlook 설치 확인
   - 제어판 → 프로그램 → Microsoft Office

3. COM 라이브러리 확인
   - regsvcs.exe로 COM 재등록

4. Windows 업데이트
   - 최신 버전 설치

5. 관리자 권한 확인
   - PowerShell을 관리자로 실행
```

---

**이제 시작할 준비가 되었습니다!** 🎉
