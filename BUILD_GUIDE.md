# 📥 설치 및 빌드 가이드

## 🔧 개발 환경 설정

### 1. .NET 8.0 설치

```bash
# 현재 설치된 .NET 버전 확인
dotnet --version

# .NET 8.0이 없으면 설치
# https://dotnet.microsoft.com/download/dotnet/8.0
```

### 2. Visual Studio 2022 설치 (권장)

```
1. https://visualstudio.microsoft.com/download/ 접속
2. Visual Studio 2022 Community 다운로드
3. 설치 중 다음 워크로드 선택:
   ✓ .NET 데스크톱 개발
   ✓ Windows 데스크톱 개발 (C++)
4. 설치 완료
```

### 3. Outlook 설치

이 앱은 Microsoft Outlook이 필수입니다.

- **Classic Outlook**: Office 2016 이상
- **New Outlook**: Microsoft 365 구독자용

## 🚀 빌드 방법

### Visual Studio 2022 사용

```
1. Visual Studio 2022 실행
2. "프로젝트 열기" → OutlookAutomation.csproj 선택
3. 빌드 메뉴 → 솔루션 빌드 (또는 Ctrl+Shift+B)
4. 빌드 완료 후 F5 키로 실행
```

### .NET CLI 사용

#### 디버그 모드

```bash
# 프로젝트 폴더에서
dotnet build
dotnet run
```

#### Release 빌드

```bash
# 최적화된 바이너리 생성
dotnet build --configuration Release

# 실행
./bin/Release/net8.0-windows/OutlookAutomationApp.exe
```

## 📦 배포 패키지 생성

### 독립 실행형 EXE (Self-Contained)

```bash
# .NET 런타임을 포함한 완전한 패키지 (약 120MB)
dotnet publish -c Release -r win-x64 --self-contained

# 생성 경로: bin/Release/net8.0-windows/win-x64/publish/
```

### 프레임워크 종속 EXE (Framework-Dependent)

```bash
# .NET 런타임이 별도로 필요함 (약 20MB)
dotnet publish -c Release

# 생성 경로: bin/Release/net8.0-windows/publish/
# 배포 시 대상 PC에 .NET 8.0 필요
```

### 최소 크기 자체 포함 배포

```bash
# 불필요한 파일 제거하여 크기 최소화 (약 50MB)
dotnet publish -c Release -r win-x64 \
  --self-contained \
  -p:SelfContainedRid=win-x64 \
  -p:PublishTrimmed=true

# 대신 일부 리플렉션 기능이 작동하지 않을 수 있음
```

## 📋 프로젝트 설정 상세

### OutlookAutomation.csproj

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>          <!-- Windows GUI 앱 -->
    <TargetFramework>net8.0-windows</TargetFramework>  <!-- .NET 8.0 -->
    <UseWPF>true</UseWPF>                   <!-- WPF 활성화 -->
    <UseWindowsForms>true</UseWindowsForms> <!-- WinForms 활성화 (폴더 선택용) -->
    <Nullable>enable</Nullable>              <!-- Nullable 참조 타입 활성화 -->
    <LangVersion>latest</LangVersion>       <!-- 최신 C# 버전 -->
    <Platform>x64</Platform>                <!-- 64비트 플랫폼 -->
  </PropertyGroup>
</Project>
```

### 설정 설명

| 설정 | 의미 |
|------|------|
| `WinExe` | Windows GUI 응용 프로그램 (콘솔 창 없음) |
| `net8.0-windows` | Windows 전용 .NET 8.0 |
| `UseWPF` | 메인 UI (MainWindow.xaml) |
| `UseWindowsForms` | 폴더 선택 대화상자용 |
| `Nullable=enable` | Null 안전성 검사 |
| `Platform=x64` | 64비트 전용 (Outlook COM 호환성) |

## ✅ 빌드 확인

### 빌드 성공 확인

```bash
# 다음과 같은 메시지 표시
# Build succeeded.
# 0 Warning(s)
# 0 Error(s)
```

### 실행 파일 위치

**Debug:**
```
bin/Debug/net8.0-windows/OutlookAutomationApp.exe
```

**Release:**
```
bin/Release/net8.0-windows/OutlookAutomationApp.exe
```

## 🔍 디버깅

### Visual Studio에서 디버그

```
1. F5 키로 디버그 시작
2. 중단점(Breakpoint) 설정: 라인 좌측 클릭
3. 변수 확인: Debug 창 또는 마우스 호버
4. F10(Step Over) 또는 F11(Step Into)로 단계 실행
```

### 콘솔 로그 확인

코드의 `Console.WriteLine()` 또는 `Debug.WriteLine()`으로 로그를 출력합니다.

```csharp
Debug.WriteLine("디버그 메시지");  // 디버그 중에만 표시
Console.WriteLine("콘솔 메시지");   // 콘솔 창에 표시
```

## 🆘 트러블슈팅

### 빌드 실패: "WPF가 로드되지 않음"

```bash
# 문제: WPF 컴포넌트 누락
# 해결: Visual Studio 복구 또는 재설치

# Visual Studio Installer에서:
# Modify → 개별 구성 요소 → 데스크톱 개발 → 다시 설치
```

### 빌드 실패: "Outlook COM 참조 없음"

```bash
# 문제: Outlook이 설치되지 않음
# 해결: Microsoft Outlook 설치

# 또는 COM 참조 수동 추가:
# 프로젝트 우클릭 → Add Reference → COM
# "Microsoft Outlook" 선택
```

### 실행 실패: "DLL을 찾을 수 없음"

```bash
# 문제: .NET Runtime 또는 의존성 누락
# 해결:
dotnet --version  # .NET 8.0 확인

# 없으면 설치:
# https://dotnet.microsoft.com/download/dotnet/8.0
```

### 실행 시 "Outlook 자동화 실패"

```bash
# 원인 1: Outlook이 실행 중이 아님
# → Outlook을 먼저 실행

# 원인 2: Outlook이 Classic/New Outlook 혼합
# → 같은 버전 유지

# 원인 3: COM 보안 설정
# → Outlook: Tools → Trust Center → Programmer Access
# → "Trust all installed add-ins" 체크
```

## 📊 빌드 최적화

### 빠른 빌드

```bash
# 증분 빌드 (변경된 파일만 컴파일)
dotnet build
```

### 최적화된 Release 빌드

```bash
# 성능 최적화 (-O 또는 -O2)
dotnet build --configuration Release
```

### 병렬 빌드

```bash
# 최대 성능으로 병렬 컴파일
dotnet build -m
```

## 🔐 보안 서명 (선택)

배포 전에 실행 파일에 디지털 서명을 추가할 수 있습니다.

```bash
# 자체 서명 인증서 생성
# (본격적인 배포 시만 필요)
```

## 📈 성능 프로파일링

Visual Studio에서 성능 프로파일러를 실행합니다:

```
1. Debug → Performance Profiler
2. CPU 사용량, 메모리 등 분석
3. 병목 지점 식별 및 최적화
```

---

**다음 단계**: 앱 실행 후 [사용 가이드](USAGE.md)를 참고하세요.
