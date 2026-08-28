using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;  // ZIP 압축 기능
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Windows;
using Microsoft.Win32;

namespace OutlookAutomationApp
{
    [SupportedOSPlatform("windows")]
    public partial class MainWindow : Window
    {
        private readonly OutlookAutomationService _service;

        public MainWindow()
        {
            InitializeComponent();
            _service = new OutlookAutomationService();
            SetupUI();
        }

        private void SetupUI()
        {
            Title = "Outlook 파일 첨부 자동화";
            this.Width = 700;
            this.Height = 560;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "첨부할 파일을 선택하세요",
                Filter = "모든 파일 (*.*)|*.*" +
                         "|Excel 파일 (*.xlsx;*.xls)|*.xlsx;*.xls" +
                         "|Word 파일 (*.docx;*.doc)|*.docx;*.doc" +
                         "|PowerPoint 파일 (*.pptx;*.ppt)|*.pptx;*.ppt" +
                         "|PDF 파일 (*.pdf)|*.pdf" +
                         "|한글 파일 (*.hwp;*.hwpx)|*.hwp;*.hwpx" +
                         "|이미지 파일 (*.png;*.jpg;*.bmp)|*.png;*.jpg;*.bmp" +
                         "|텍스트 파일 (*.txt)|*.txt",
                FilterIndex = 1,
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _service.SourceFile = openFileDialog.FileName;
                TxtSourceFile.Text = _service.SourceFile;
                UpdateStatus($"📄 파일 선택됨: {Path.GetFileName(_service.SourceFile)}");
            }
        }

        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderDialog.Description = "저장할 폴더를 선택하세요";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _service.OutputFolder = folderDialog.SelectedPath;
                    TxtOutputFolder.Text = _service.OutputFolder;
                    UpdateStatus($"📁 폴더 선택됨: {_service.OutputFolder}");
                }
            }
        }

        private async void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs())
                return;

            // 실행 중 모든 버튼 비활성화 및 진행 표시줄 표시
            BtnExecute.IsEnabled = false;
            BtnSelectFile.IsEnabled = false;
            BtnSelectFolder.IsEnabled = false;
            ProgressBar.Visibility = Visibility.Visible;

            try
            {
                // UI 업데이트 콜백 등록 (백그라운드 스레드 → UI 스레드 안전 전달)
                _service.OnStatusUpdate = (msg) =>
                {
                    Dispatcher.Invoke(() => UpdateStatus(msg));
                };

                var result = await System.Threading.Tasks.Task.Run(() =>
                    _service.ExecuteAutomation());

                if (result.Success)
                {
                    // 어떤 방식으로 완료됐는지 메시지에 포함
                    string methodInfo = result.UsedFallback
                        ? "\n\n※ 현재 PC에 새 Outlook이 설치되어 있어\n파일을 직접 복사하는 방식으로 처리했습니다."
                        : "\n\n✅ Outlook COM 자동화 방식으로 처리했습니다.";

                    // ZIP 압축 완료 정보 추가
                    string zipInfo = !string.IsNullOrEmpty(result.ZipFilePath)
                        ? $"\n\n📦 ZIP 압축 파일:\n{result.ZipFilePath}"
                        : string.Empty;

                    UpdateStatus("✅ 모든 작업이 완료되었습니다!", "Green");
                    MessageBox.Show(
                        $"자동화가 완료되었습니다!\n\n저장 위치:\n{result.SavedFilePath}{zipInfo}{methodInfo}",
                        "성공", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    UpdateStatus($"❌ 작업 실패: {result.ErrorMessage}", "Red");
                    MessageBox.Show(
                        $"작업 중 오류가 발생했습니다:\n\n{result.ErrorMessage}",
                        "오류 상세 내용", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ 예외 발생: {ex.Message}", "Red");
                MessageBox.Show(
                    $"시스템 예외:\n\n{ex.Message}\n\n{ex.StackTrace}",
                    "예외 발생", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnExecute.IsEnabled = true;
                BtnSelectFile.IsEnabled = true;
                BtnSelectFolder.IsEnabled = true;
                ProgressBar.Visibility = Visibility.Hidden;
            }
        }

        private bool ValidateInputs()
        {
            if (string.IsNullOrEmpty(_service.SourceFile))
            {
                MessageBox.Show("파일을 선택하세요.", "입력 오류",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (string.IsNullOrEmpty(_service.OutputFolder))
            {
                MessageBox.Show("폴더를 선택하세요.", "입력 오류",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (!File.Exists(_service.SourceFile))
            {
                MessageBox.Show("선택된 파일이 존재하지 않습니다.", "파일 오류",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (!Directory.Exists(_service.OutputFolder))
            {
                MessageBox.Show("선택된 폴더가 존재하지 않습니다.", "폴더 오류",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private void UpdateStatus(string message, string color = "Black")
        {
            TextBlockStatus.Text = message;
            TextBlockStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(color));
        }
    }

    // ============================================================
    // OutlookAutomationService: 핵심 자동화 로직
    // ============================================================
    [SupportedOSPlatform("windows")]
    public class OutlookAutomationService
    {
        public string SourceFile { get; set; } = string.Empty;
        public string OutputFolder { get; set; } = string.Empty;
        public Action<string>? OnStatusUpdate { get; set; }

        // Outlook COM 관련 객체
        private dynamic? _outlookApp;
        private dynamic? _mailItem;

        // 파일을 열었던 프로세스 참조 (종료 시 사용)
        private Process? _openedProcess;

        // ─────────────────────────────────────────────────────────
        // Simple MAPI P/Invoke: 클래식/새 Outlook 모두 지원하는
        // Windows 기본 메일 API (fallback용)
        // ─────────────────────────────────────────────────────────

        /// <summary>
        /// Windows MAPI 구조체: 파일 첨부 정보
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private struct MapiFileDesc
        {
            public uint ulReserved;
            public uint flFlags;        // 0 = 일반 첨부
            public uint nPosition;      // 본문 내 위치 (0xFFFFFFFF = 끝)
            public string lpszPathName; // 파일 실제 경로
            public string lpszFileName; // 표시될 파일명
            public IntPtr lpFileType;   // null = 자동 감지
        }

        /// <summary>
        /// Windows MAPI 구조체: 메일 메시지 전체 정보
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private struct MapiMessage
        {
            public uint ulReserved;
            public string lpszSubject;     // 제목
            public string lpszNoteText;    // 본문
            public string lpszMessageType;
            public string lpszDateReceived;
            public string lpszConversationID;
            public uint flFlags;
            public IntPtr lpOriginator;
            public uint nRecipCount;
            public IntPtr lpRecips;
            public uint nFileCount;        // 첨부 파일 수
            public IntPtr lpFiles;         // MapiFileDesc 배열 포인터
        }

        // MAPI32.DLL의 MAPISendMail 함수 가져오기
        // 이 함수는 Windows에 등록된 기본 메일 클라이언트(Outlook 등)를 통해
        // 메일 작성 창을 표시하며, 클래식/새 Outlook 모두 지원
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
        private static extern uint MAPISendMail(
            IntPtr lhSession,           // MAPI 세션 (0 = 기본 세션 사용)
            IntPtr ulUIParam,           // 부모 창 핸들
            ref MapiMessage lpMessage,  // 메일 메시지 구조체
            uint flFlags,               // MAPI_DIALOG(8) | MAPI_LOGON_UI(1)
            uint ulReserved);           // 항상 0

        // MAPI 반환 코드 상수
        private const uint MAPI_SUCCESS = 0;
        private const uint MAPI_DIALOG = 8;    // 작성 창 표시
        private const uint MAPI_LOGON_UI = 1;  // 로그인 UI 허용

        /// <summary>
        /// 자동화 전체 프로세스 실행
        /// 반환: (Success, ErrorMessage, SavedFilePath, UsedFallback)
        /// </summary>
        public (bool Success, string ErrorMessage, string SavedFilePath, string ZipFilePath, bool UsedFallback) ExecuteAutomation()
        {
            string savedFilePath = string.Empty;
            string zipFilePath = string.Empty;

            try
            {
                // ── 1단계: 파일을 기본 앱으로 열기 ──────────────────
                ReportStatus("📂 [1/5] 파일을 기본 앱으로 열고 있습니다...");
                if (!OpenFileWithDefaultApp(out string openErr))
                    return (false, openErr, string.Empty, string.Empty, false);

                // 앱이 완전히 로딩될 때까지 대기 (2.5초)
                ReportStatus("⏳ [1/5] 앱이 파일을 여는 중 대기 (2.5초)...");
                System.Threading.Thread.Sleep(2500);

                // ── 2단계: Outlook 방식 선택 ────────────────────────
                ReportStatus("🔍 [2/5] Outlook 버전을 확인하고 있습니다...");

                // 클래식 Outlook COM 사용 가능 여부 확인
                bool classicOutlookAvailable = IsClassicOutlookAvailable();

                if (classicOutlookAvailable)
                {
                    // ══ 경로 A: 클래식 Outlook COM 자동화 ══════════════
                    ReportStatus("📧 [2/5] 클래식 Outlook COM 자동화를 사용합니다...");

                    if (!InitializeOutlook(out string initErr))
                        return (false, initErr, string.Empty, string.Empty, false);

                    ReportStatus("📎 [3/5] 메일 항목을 생성하고 파일을 첨부합니다...");
                    if (!CreateMailWithAttachment(out string mailErr))
                        return (false, mailErr, string.Empty, string.Empty, false);

                    ReportStatus("👁 [4/5] Outlook 메일 창을 표시합니다...");
                    DisplayMail();
                    System.Threading.Thread.Sleep(3000);

                    ReportStatus("💾 [5/5] 첨부 파일을 저장합니다...");
                    if (!SaveAttachmentViaCOM(out string saveErr, out savedFilePath))
                        return (false, saveErr, string.Empty, string.Empty, false);

                    // ── ZIP 압축 ────────────────────────────────────
                    ReportStatus("📦 [+] 저장된 파일을 ZIP으로 압축합니다...");
                    CompressToZip(savedFilePath, out zipFilePath);

                    return (true, string.Empty, savedFilePath, zipFilePath, false);
                }
                else
                {
                    // ══ 경로 B: 폴백 - 파일 직접 복사 ════════════════
                    // 새 Outlook(olk.exe)은 COM을 지원하지 않으므로
                    // Simple MAPI로 메일 창을 열어 사용자에게 표시한 뒤
                    // 파일을 직접 복사하여 동일한 결과를 달성합니다.

                    string outlookKind = DetectOutlookKind();
                    ReportStatus($"⚠️ [2/5] {outlookKind} - 직접 복사 방식으로 전환합니다...");

                    // MAPI로 메일 작성 창 열기 시도 (표시용, 사용자 확인)
                    ReportStatus("📧 [3/5] Simple MAPI로 메일 창을 표시합니다...");
                    bool mapiShown = TryShowMailViaMAPI();

                    if (!mapiShown)
                    {
                        ReportStatus("ℹ️ [3/5] MAPI 창 표시 건너뜀 - 파일 직접 복사 진행...");
                    }

                    // 잠시 대기 (메일 창이 표시된 경우 사용자 확인 시간)
                    System.Threading.Thread.Sleep(1500);

                    // 파일을 출력 폴더에 직접 복사
                    ReportStatus("💾 [5/5] 파일을 출력 폴더에 복사합니다...");
                    if (!CopyFileDirect(out string copyErr, out savedFilePath))
                        return (false, copyErr, string.Empty, string.Empty, true);

                    // ── ZIP 압축 ────────────────────────────────────
                    ReportStatus("📦 [+] 저장된 파일을 ZIP으로 압축합니다...");
                    CompressToZip(savedFilePath, out zipFilePath);

                    return (true, string.Empty, savedFilePath, zipFilePath, true);
                }
            }
            catch (Exception ex)
            {
                return (false,
                    $"자동화 프로세스 예외: {ex.Message}\n\n스택 트레이스:\n{ex.StackTrace}",
                    string.Empty, string.Empty, false);
            }
            finally
            {
                ReportStatus("🧹 정리 중... (앱 종료)");
                Cleanup();
            }
        }

        /// <summary>
        /// 클래식 Outlook(OUTLOOK.EXE)의 COM 자동화 사용 가능 여부 확인
        /// Office 2007, 2010, 2013, 2016, 2019, 365 Classic 모두 지원
        /// 새 Outlook (olk.exe / New Outlook)은 COM ProgID를 등록하지 않으므로 false 반환
        /// </summary>
        private bool IsClassicOutlookAvailable()
        {
            try
            {
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                return outlookType != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 설치된 Outlook 종류를 문자열로 반환 (로그/상태 표시용)
        /// </summary>
        private string DetectOutlookKind()
        {
            // 새 Outlook (Windows 11 기본 앱: olk.exe)이 실행 중인지 확인
            var newOutlookProcs = Process.GetProcessesByName("olk");
            if (newOutlookProcs.Length > 0)
                return "새 Outlook (New Outlook/olk.exe) 감지됨";

            // 클래식 Outlook 프로세스가 실행 중인지 확인
            var classicProcs = Process.GetProcessesByName("OUTLOOK");
            if (classicProcs.Length > 0)
                return "Outlook 프로세스 실행 중이나 COM 연결 불가";

            return "Outlook을 찾을 수 없음";
        }

        // ─────────────────────────────────────────────────────────
        // Simple MAPI를 이용해 메일 작성 창 표시 (경로 B에서 사용)
        // 새 Outlook 포함, Windows 기본 메일 클라이언트 모두 지원
        // ─────────────────────────────────────────────────────────
        private bool TryShowMailViaMAPI()
        {
            try
            {
                string fileName = Path.GetFileName(SourceFile);

                // 첨부 파일 구조체 설정
                MapiFileDesc fileDesc = new MapiFileDesc
                {
                    ulReserved = 0,
                    flFlags = 0,
                    nPosition = 0xFFFFFFFF,   // 본문 끝에 첨부
                    lpszPathName = SourceFile, // 실제 파일 경로
                    lpszFileName = fileName,   // 표시될 파일명
                    lpFileType = IntPtr.Zero
                };

                // 비관리 메모리에 MapiFileDesc 구조체 할당
                IntPtr filePtr = Marshal.AllocHGlobal(Marshal.SizeOf(fileDesc));
                Marshal.StructureToPtr(fileDesc, filePtr, false);

                // 메일 메시지 구조체 설정
                MapiMessage message = new MapiMessage
                {
                    ulReserved = 0,
                    lpszSubject = $"첨부: {fileName}",
                    lpszNoteText = $"자동화 프로그램이 생성한 임시 메일입니다.\n\n파일: {SourceFile}",
                    lpszMessageType = string.Empty,
                    lpszDateReceived = string.Empty,
                    lpszConversationID = string.Empty,
                    flFlags = 0,
                    lpOriginator = IntPtr.Zero,
                    nRecipCount = 0,
                    lpRecips = IntPtr.Zero,
                    nFileCount = 1,
                    lpFiles = filePtr
                };

                // MAPI_DIALOG | MAPI_LOGON_UI: 메일 작성 창 표시 + 로그인 허용
                uint result = MAPISendMail(
                    IntPtr.Zero, IntPtr.Zero,
                    ref message,
                    MAPI_DIALOG | MAPI_LOGON_UI,
                    0);

                // 비관리 메모리 해제
                Marshal.FreeHGlobal(filePtr);

                // MAPI_SUCCESS(0) 또는 사용자가 취소(1)해도 "표시 시도"는 성공으로 처리
                return result == MAPI_SUCCESS || result == 1;
            }
            catch
            {
                // MAPI 실패는 치명적이지 않음 - 파일 직접 복사로 계속 진행
                return false;
            }
        }

        /// <summary>
        /// 소스 파일을 출력 폴더에 직접 복사 (경로 B 폴백)
        /// Outlook attachment.SaveAsFile()과 동일한 바이트 수준 결과 생성
        /// </summary>
        /// <summary>
        /// 저장된 파일을 같은 출력 폴더 안에 ZIP으로 압축
        /// 파일명 예: Abaqus.pptx → Abaqus_20240828_143052.zip
        /// ZIP 생성 실패는 치명적이지 않으므로 예외를 삼켜서 계속 진행
        /// </summary>
        private void CompressToZip(string sourceFilePath, out string zipFilePath)
        {
            zipFilePath = string.Empty;
            try
            {
                if (!File.Exists(sourceFilePath))
                    return;

                // 원본 파일명 그대로 사용 (확장자만 .zip으로 교체)
                // 예: Abaqus.pptx → Abaqus.zip
                string nameWithoutExt = Path.GetFileNameWithoutExtension(sourceFilePath);
                string zipPath = Path.Combine(OutputFolder, $"{nameWithoutExt}.zip");

                // 동일 이름의 ZIP이 이미 존재하면 삭제 후 새로 생성
                if (File.Exists(zipPath))
                    File.Delete(zipPath);

                // ZipArchive: 새 ZIP 파일 생성
                using (ZipArchive zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                {
                    // 원본 파일명 유지하면서 ZIP 내부에 추가
                    zip.CreateEntryFromFile(
                        sourceFilePath,
                        Path.GetFileName(sourceFilePath),
                        CompressionLevel.Optimal);  // 최적 압축률
                }

                zipFilePath = zipPath;
                ReportStatus($"📦 ZIP 압축 완료: {Path.GetFileName(zipPath)}");
            }
            catch (Exception ex)
            {
                // ZIP 실패는 전체 작업을 중단시키지 않음 (경고만 표시)
                ReportStatus($"⚠️ ZIP 압축 실패 (파일 저장은 완료됨): {ex.Message}");
            }
        }

        private bool CopyFileDirect(out string error, out string savedFilePath)
        {
            error = string.Empty;
            savedFilePath = string.Empty;
            try
            {
                string fileName = Path.GetFileName(SourceFile);
                string outputPath = Path.Combine(OutputFolder, fileName);

                // 동일 파일명 존재 시 타임스탬프 기반 고유 이름 생성
                if (File.Exists(outputPath))
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    string ext = Path.GetExtension(fileName);
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    outputPath = Path.Combine(OutputFolder, $"{nameWithoutExt}_{timestamp}{ext}");
                }

                // 파일 복사 실행
                File.Copy(SourceFile, outputPath, overwrite: false);
                savedFilePath = outputPath;
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                error = $"출력 폴더에 쓰기 권한이 없습니다.\n폴더: {OutputFolder}";
                return false;
            }
            catch (Exception ex)
            {
                error = $"파일 복사 실패: {ex.Message}";
                return false;
            }
        }

        // ─────────────────────────────────────────────────────────
        // 이하 경로 A (클래식 Outlook COM) 전용 메서드들
        // ─────────────────────────────────────────────────────────

        private void ReportStatus(string message) => OnStatusUpdate?.Invoke(message);

        private bool OpenFileWithDefaultApp(out string error)
        {
            error = string.Empty;
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = SourceFile,
                    UseShellExecute = true, // 셸 실행 → 파일 형식별 기본 앱으로 오픈
                    Verb = "open"
                };
                // 프로세스 참조 저장 → Cleanup()에서 종료 가능
                _openedProcess = Process.Start(psi);
                return true;
            }
            catch (Exception ex)
            {
                error = $"기본 앱으로 파일 열기 실패: {ex.Message}\n파일: {SourceFile}";
                return false;
            }
        }

        private bool InitializeOutlook(out string error)
        {
            error = string.Empty;
            try
            {
                // "Outlook.Application" ProgID는 Office 2007 ~ 365 Classic 모두 동일
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    error = "클래식 Outlook COM ProgID를 찾을 수 없습니다.\n" +
                            "이 경로는 호출되지 않아야 합니다. 버그 보고 바랍니다.";
                    return false;
                }
                _outlookApp = Activator.CreateInstance(outlookType);
                if (_outlookApp == null)
                {
                    error = "Outlook COM 인스턴스 생성에 실패했습니다.";
                    return false;
                }
                return true;
            }
            catch (COMException comEx)
            {
                error = $"Outlook COM 초기화 실패 (0x{comEx.HResult:X8}): {comEx.Message}";
                return false;
            }
            catch (Exception ex)
            {
                error = $"Outlook 초기화 예외: {ex.Message}";
                return false;
            }
        }

        private bool CreateMailWithAttachment(out string error)
        {
            error = string.Empty;
            try
            {
                if (_outlookApp == null)
                {
                    error = "Outlook COM 객체가 초기화되지 않았습니다.";
                    return false;
                }
                // olMailItem = 0
                _mailItem = _outlookApp.CreateItem(0);
                if (_mailItem == null)
                {
                    error = "새 메일 항목 생성에 실패했습니다.";
                    return false;
                }

                string fileName = Path.GetFileName(SourceFile);
                _mailItem.Subject = $"첨부: {fileName}";
                _mailItem.Body = $"자동화 프로그램이 생성한 임시 메일입니다.\n\n파일: {SourceFile}";
                _mailItem.Attachments.Add(SourceFile);
                return true;
            }
            catch (COMException comEx)
            {
                error = $"메일 첨부 실패 (0x{comEx.HResult:X8}): {comEx.Message}\n\n" +
                        "가능한 원인:\n• 파일이 다른 프로그램에서 잠금 중\n" +
                        "• 파일 경로에 특수문자\n• 파일 접근 권한 없음";
                return false;
            }
            catch (Exception ex)
            {
                error = $"메일 작성 예외: {ex.Message}";
                return false;
            }
        }

        private void DisplayMail()
        {
            try { _mailItem?.Display(false); } catch { }
        }

        private bool SaveAttachmentViaCOM(out string error, out string savedFilePath)
        {
            error = string.Empty;
            savedFilePath = string.Empty;
            try
            {
                if (_mailItem == null || _mailItem!.Attachments.Count == 0)
                {
                    error = "저장할 첨부 파일이 존재하지 않습니다.";
                    return false;
                }

                // COM 인덱스는 1부터 시작
                dynamic attachment = _mailItem!.Attachments[1];
                string attachmentName = (string)attachment.FileName;
                string outputPath = Path.Combine(OutputFolder, attachmentName);

                // 동일 파일명 존재 시 타임스탬프 기반 고유 이름 생성
                if (File.Exists(outputPath))
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(attachmentName);
                    string ext = Path.GetExtension(attachmentName);
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    outputPath = Path.Combine(OutputFolder, $"{nameWithoutExt}_{timestamp}{ext}");
                }

                attachment.SaveAsFile(outputPath);
                savedFilePath = outputPath;
                return true;
            }
            catch (COMException comEx)
            {
                error = $"첨부 파일 저장 실패 (0x{comEx.HResult:X8}): {comEx.Message}";
                return false;
            }
            catch (Exception ex)
            {
                error = $"파일 저장 예외: {ex.Message}";
                return false;
            }
        }

        /// <summary>
        /// 자동화 완료/실패 후 모든 리소스 안전하게 정리
        /// 순서: 메일닫기 → Outlook종료 → COM해제 → GC → 편집앱종료
        /// </summary>
        private void Cleanup()
        {
            // 1) 메일 아이템 닫기 (olDiscard=1: 저장 안 함)
            if (_mailItem != null)
            {
                try { _mailItem.Close(1); } catch { }
                try { Marshal.ReleaseComObject(_mailItem); } catch { }
                _mailItem = null;
            }

            // 2) Outlook 종료 후 COM 해제 (Quit() 없으면 좀비 프로세스 남음)
            if (_outlookApp != null)
            {
                try { _outlookApp.Quit(); } catch { }
                try { Marshal.ReleaseComObject(_outlookApp); } catch { }
                _outlookApp = null;
            }

            // 3) GC로 미수집 COM 객체 정리 (권장 패턴)
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // 4) 파일을 열었던 앱(Excel, Word, 한글 등) 종료
            if (_openedProcess != null)
            {
                try
                {
                    if (!_openedProcess.HasExited)
                    {
                        // 우아한 종료 시도 (저장 다이얼로그 가능성 있음)
                        _openedProcess.CloseMainWindow();
                        bool exited = _openedProcess.WaitForExit(3000);
                        if (!exited)
                            _openedProcess.Kill(); // 3초 내 응답 없으면 강제 종료
                    }
                }
                catch { }
                finally
                {
                    _openedProcess.Dispose();
                    _openedProcess = null;
                }
            }
        }
    }
}