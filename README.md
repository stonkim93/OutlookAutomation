# Outlook Automation App

> Windows에서 파일을 Outlook에 첨부하여 저장하는 자동화 도구

[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D4?logo=windows11&logoColor=white)](https://www.microsoft.com/windows)
[![Framework](https://img.shields.io/badge/.NET-8.0--windows-512BD4?logo=dotnet&logoColor=white)](https://dotnet.microsoft.com)
[![Language](https://img.shields.io/badge/language-C%23-239120?logo=csharp&logoColor=white)](https://docs.microsoft.com/dotnet/csharp)
[![License](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

## ✨ 주요 기능

- 📄 **다양한 파일 형식 지원** (Excel, Word, PowerPoint, PDF, 한글 등)
- 📧 **Outlook 자동 연동** (Classic Outlook COM을 통한 안정적인 백그라운드 처리)
- 💾 **안전한 DRM 복호화 및 다이렉트 ZIP 압축** (디스크 중간 저장 없이 메모리에서 바로 압축)
- 📦 **폴더명으로 압축하기 (배치 처리)** (드롭된 여러 파일을 폴더명 하나로 묶어서 압축)
- 🛡️ **개별 문서 단위 안전 종료 지원** (사용자가 엑셀 등으로 진행 중이던 기존 작업을 보존)
- 🖱️ **Drag & Drop 지원** (파일을 앱에 끌어다 놓으면 즉시 자동화 실행)

## 🚀 빠른 시작

### 요구사항

- Windows 10 / 11 (x64)
- Microsoft Outlook (Classic 또는 New Outlook)
- .NET 8.0 이상

### 설치 및 실행

```bash
# 저장소 복제
git clone https://github.com/your-username/OutlookAutomationApp.git
cd OutlookAutomationApp

# 빌드
dotnet build --configuration Release

# 실행
dotnet run
```

또는 Release 폴더에서 `.exe` 파일을 직접 실행합니다.

## 📖 사용 방법

### 1단계: 드래그 앤 드롭 (가장 쉬운 방법)
복호화할 DRM 파일(예: `document.xlsx`)을 앱 창에 **마우스로 끌어다 놓습니다.**
별도의 버튼 클릭 없이 즉시 백그라운드에서 복호화 및 압축 작업이 시작됩니다.

### 2단계: 자동 처리 및 완료
파일을 드롭하면 앱이 자동으로 다음을 수행합니다:
1. ✅ 파일을 기본 앱(Excel, Word 등)의 COM 인터페이스로 엽니다.
2. ✅ 클래식 Outlook을 실행하고 임시 메일에 파일을 첨부합니다.
3. ✅ 첨부된 파일의 복호화된 메모리 데이터를 추출하여 바로 ZIP 파일로 압축합니다.
4. ✅ 사용자가 작업 중이던 다른 문서를 건드리지 않고, 연 문서만 안전하게 닫습니다.

> **💡 옵션: 폴더명으로 압축하기**
> 실행 버튼 왼쪽의 옵션을 체크해두면, 드롭한 파일들이 들어있는 폴더 이름으로 단일 ZIP 파일이 생성되며, 여러 번 드롭해도 같은 ZIP 파일 내부에 차곡차곡 누적 압축됩니다.

## 🏗️ 프로젝트 구조

```
OutlookAutomationApp/
├── App.xaml
├── App.xaml.cs
├── MainWindow.xaml           # UI 레이아웃
├── MainWindow.xaml.cs        # 로직 및 이벤트 처리
├── OutlookAutomation.csproj  # 프로젝트 설정
└── README.md
```

## 🛠️ 기술 스택

| 항목 | 내용 |
|------|------|
| **Framework** | .NET 8.0-windows |
| **UI** | WPF (Windows Presentation Foundation) |
| **Language** | C# |
| **Architecture** | x64 |
| **Interop** | COM (Outlook.Application), MAPI (outlook.exe) |

## 🔧 기술 상세

### Outlook 호환성

```
✓ Classic Outlook (Office 2007-365)
  → COM Interop로 백그라운드 자동화
  → MAPI PropertyAccessor 훅 우회를 통한 직접 바이트 추출 지원
```

### 폴더 단위 누적 압축 지원

```
폴더명으로 압축하기 옵션 활성화 시:
C:\Work\report1.xlsx 드롭 -> C:\Work\Work.zip 에 압축됨
C:\Work\report2.xlsx 드롭 -> C:\Work\Work.zip 에 추가로 압축됨
```

### 리소스 관리 및 안전한 종료

- 파일 잠금 방지를 위한 적절한 대기 시간 설정
- Excel, Word, PPT 등 COM 객체를 정교하게 제어하여, 사용자가 기존에 작성 중이던 다른 문서를 강제 종료하는 사고(Data Loss) 방지
- 완료 후 사용된 문서 및 COM 객체만 안전하게 닫고 메모리 누수 방지

## 📋 시스템 요구사항

- **OS**: Windows 10 (21H2) 이상 또는 Windows 11
- **Outlook**: 2016 이상 (Classic) 또는 New Outlook
- **RAM**: 최소 2GB
- **Disk**: 최소 50MB 여유 공간

## ⚙️ 개발 환경 설정

### Visual Studio 2022

1. Visual Studio 2022 Community 설치
2. ".NET 데스크톱 개발" 워크로드 선택
3. 프로젝트 열기 및 `F5` 키로 실행

### .NET CLI

```bash
# .NET 8.0 SDK 설치 확인
dotnet --version

# 프로젝트 빌드
dotnet build

# 실행
dotnet run

# Release 빌드
dotnet publish -c Release -o ./publish
```

## 🐛 문제 해결

### "Outlook을 찾을 수 없음" 오류

Outlook이 설치되어 있지 않거나 실행되지 않은 경우입니다.
→ Microsoft Outlook을 설치하거나 실행 후 다시 시도하세요.

### "파일이 잠겨있음" 오류

Excel이나 Word 등의 앱이 파일을 열고 있는 경우입니다.
→ 파일을 닫은 후 다시 시도하세요.

### "저장 경로가 없음" 오류

지정한 출력 폴더가 존재하지 않습니다.
→ 유효한 폴더를 선택하세요.

## 📦 배포

### EXE 파일로 배포

```bash
# Release 빌드
dotnet publish -c Release -r win-x64 --self-contained

# 생성된 파일:
# publish/OutlookAutomationApp.exe
```

### MSI 설치 프로그램 (선택)

Visual Studio의 "Create Installer" 기능을 사용하거나 WiX Toolset을 이용합니다.

## 📝 라이선스

MIT License - 자유롭게 수정 및 배포할 수 있습니다.

## 👨‍💻 기여

Issues 및 Pull Requests를 환영합니다!

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📧 문의

문제가 있거나 제안사항이 있으시면 Issues를 통해 알려주세요.

---

**Made with ❤️ for Windows**
