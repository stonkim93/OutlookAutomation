<div align="center">

# 🌍 OutlookAutomation

### I'm a Color Pointer that helps to identify the keyboard input language.
(English Lower, English Upper, Korean, Pali, Japanese)

### 입력 모드에 따라 색상이 바뀌는 마우스 포인터 유틸리티

![Platform](https://img.shields.io/badge/platform-Windows-0078D4?logo=windows11&logoColor=white)
![Framework](https://img.shields.io/badge/.NET-10.0--windows-512BD4?logo=dotnet&logoColor=white)
![Language](https://img.shields.io/badge/language-C%23-239120?logo=csharp&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-yellow.svg)
![Status](https://img.shields.io/badge/status-Production--Ready-2E8B57)

</div>

Outlook 파일 첨부 자동화 앱 (OutlookAutomationApp)OutlookAutomationApp은 문서 파일(Excel, Word, PowerPoint, PDF, HWP 등)을 전용 기본 프로그램으로 열고, Outlook에 첨부하여 메일을 작성한 후, 대상 폴더 저장 및 ZIP 압축까지 한 번에 자동화해 주는 Windows 데스크톱 애플리케이션입니다.  🌟 주요 기능다양한 파일 확장자 지원: Excel, Word, PowerPoint, PDF, 한글(HWP/HWPX), 이미지, 텍스트 등 모든 형식의 파일 처리.  하이브리드 Outlook 지원 (Classic & New Outlook):클래식 Outlook (Office 2007~365): COM Interop 자동화를 이용해 백그라운드 연동 및 메일 작성[cite: 37, 39, 40].새 Outlook (New Outlook / olk.exe): COM을 지원하지 않는 최신 웹 기반 Outlook 환경에서는 Simple MAPI 연동 및 안전한 직접 복사 방식(Fallback)으로 자동 전환.  자동 ZIP 압축 기능: 지정된 출력 폴더로 원본 파일 저장 후, 해당 위치에 최적 압축률의 .zip 압축 파일 자동 생성.  스레드 안전 UI 및 실시간 상태 로깅: Dispatcher.Invoke를 적용하여 5단계 자동화 진행 상황을 UI에 실시간으로 시각화.  프로세스 및 COM 리소스 안전 정리: 파일 잠금(Lock) 방지를 위한 적절한 대기 시간 설정과 완료 후 외부 앱 프로세스 및 COM 객체 메모리 해제 처리.  🔄 동작 프로세스[1단계: 파일 선택] ──> [2단계: 출력 폴더 선택] ──> [자동화 실행]
                                                      │
 ┌────────────────────────────────────────────────────┘
 ├── 1. 지정 파일을 Windows 기본 앱으로 열기 (Excel, Word 등)[cite: 38, 39, 40]
 ├── 2. Outlook 실행 환경 및 버전 감지 (Classic COM vs New Outlook)
 ├── 3. 임시 메일 작성 및 파일 첨부 (COM 또는 Simple MAPI 호출)
 ├── 4. 작성된 메일 창 표시
 └── 5. 지정 폴더에 첨부 파일 저장 후 ZIP 압축 파일 자동 생성
🛠️ 기술 스택 및 개발 환경구분내용Framework.NET 8.0-windows[cite: 41]UI TechnologyWPF (Windows Presentation Foundation) & WinForms (폴더 선택)[cite: 38, 41]LanguageC# (Latest)[cite: 41]Architecturex64[cite: 41]InteroperabilityWindows COM Interop (Outlook.Application), Simple MAPI (MAPI32.DLL)[cite: 39, 40]CompressionSystem.IO.Compression (ZipArchive)[cite: 39]📂 프로젝트 구조OutlookAutomationApp/
├── App.xaml / App.xaml.cs          # 애플리케이션 진입점 및 공통 네임스페이스 정의
├── MainWindow.xaml                 # 메인 GUI 레이아웃 (WPF)[cite: 38]
├── MainWindow.xaml.cs              # UI 이벤트 핸들러 & OutlookAutomationService[cite: 39]
└── OutlookAutomationApp.csproj    # .NET 8.0 빌드 및 플랫폼 설정 파일[cite: 41]
🚀 시작하기요구 사항OS: Windows 10 / 11 (x64)[cite: 40, 41]개발 환경: Visual Studio 2022 (v17.8 이상) 또는 .NET 8.0 SDK[cite: 40, 41]메일 클라이언트: Microsoft Outlook (클래식 데스크톱 앱 또는 새 Outlook/New Outlook)[cite: 39, 40]빌드 및 실행저장소 클론:Bashgit clone https://github.com/your-username/OutlookAutomationApp.git
cd OutlookAutomationApp
프로젝트 빌드:Bashdotnet build --configuration Release
앱 실행:Bashdotnet run
💡 예외 처리 및 트러블슈팅파일명 중복 발생 시: 출력 폴더에 동일한 파일이 이미존재하는 경우, 타임스탬프(_yyyyMMdd_HHmmss)를 자동으로 부여하여 덮어쓰기 손실을 방지합니다.  새 Outlook(New Outlook) 호환 문제: 외부 COM 자동화가 차단된 최신 웹 기반 Outlook 실행 환경에서도 Simple MAPI 호출 및 원본 파일 복사 방식을 통해 정상 작동을 보장합니다[cite: 39].파일 점유(Lock) 오류 방지: 외부 프로그램(Excel, PowerPoint 등)이 문서를 열 때 발생하는 잠금 현상을 방지하기 위해 각 단계 간 비동기 대기 시간을 포함시켰습니다[cite: 37, 39].


## 📜 라이선스 (License)

- 이 프로젝트는 **MIT License**에 따라 자유롭게 수정 및 배포할 수 있습니다.

<br>

❤️🌍✨⚡🚀💡🎯🆕🖥️💻⌨️🔤🎨🧩🐛🔹📐📝✅🏆ℹ️❓
