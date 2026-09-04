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

            _service.ZipByFolderName = ChkZipByFolder.IsChecked == true;

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

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string filePath = files[0];
                    string dir = Path.GetDirectoryName(filePath) ?? string.Empty;

                    _service.SourceFile = filePath;
                    _service.OutputFolder = dir;
                    _service.TargetFileName = Path.GetFileName(filePath); // 저장 시 zip 내부의 파일명으로 사용

                    TxtSourceFile.Text = filePath;
                    TxtOutputFolder.Text = dir;
                    UpdateStatus($"📁 드롭된 파일 처리 중: {Path.GetFileName(filePath)}");

                    // 자동 실행
                    BtnExecute_Click(this, new RoutedEventArgs());
                }
            }
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
        public string TargetFileName { get; set; } = string.Empty;
        public bool ZipByFolderName { get; set; } = true;
        public Action<string>? OnStatusUpdate { get; set; }

        // Outlook COM 관련 객체
        private dynamic? _outlookApp;
        private dynamic? _mailItem;

        // 에디터 COM 관련 객체 (안전 종료용)
        private dynamic? _excelApp;
        private dynamic? _excelWorkbook;
        private dynamic? _wordApp;
        private dynamic? _wordDocument;
        private dynamic? _pptApp;
        private dynamic? _pptPresentation;

        // 파일을 열었던 프로세스 참조 (폴백용)
        private Process? _openedProcess;



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

                    // ══ 클래식 Outlook COM 자동화 ══════════════
                    ReportStatus("📧 [2/5] 클래식 Outlook COM 자동화를 사용합니다...");

                    if (!InitializeOutlook(out string initErr))
                        return (false, initErr, string.Empty, string.Empty, false);

                    ReportStatus("📎 [3/4] 메일 항목을 생성하고 파일을 첨부합니다...");
                    if (!CreateMailWithAttachment(out string mailErr))
                        return (false, mailErr, string.Empty, string.Empty, false);

                    ReportStatus("💾 [4/4] 첨부 파일 메모리를 추출하여 바로 ZIP으로 압축합니다...");
                    if (!SaveAttachmentToZip(out string saveErr, out zipFilePath))
                        return (false, saveErr, string.Empty, string.Empty, false);

                    return (true, string.Empty, string.Empty, zipFilePath, false);
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
        /// 첨부파일 데이터를 메모리에서 추출하여 직접 ZIP 파일로 압축 저장합니다.
        /// (Fasoo DRM SaveAs 훅 우회)
        /// </summary>
        private bool SaveAttachmentToZip(out string error, out string zipFilePath)
        {
            error = string.Empty;
            zipFilePath = string.Empty;
            try
            {
                if (_mailItem == null || _mailItem!.Attachments.Count == 0)
                {
                    error = "저장할 첨부 파일이 존재하지 않습니다.";
                    return false;
                }

                // COM 인덱스는 1부터 시작
                dynamic attachment = _mailItem!.Attachments[1];
                string targetName = !string.IsNullOrEmpty(TargetFileName) ? TargetFileName : (string)attachment.FileName;
                
                string zipPath;
                if (ZipByFolderName)
                {
                    string folderName = new DirectoryInfo(OutputFolder).Name;
                    zipPath = Path.Combine(OutputFolder, $"{folderName}.zip");
                }
                else
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(targetName);
                    zipPath = Path.Combine(OutputFolder, $"{nameWithoutExt}.zip");
                    
                    // 개별 파일 압축 시 기존 파일이 있으면 덮어쓰기 위해 삭제
                    if (File.Exists(zipPath))
                        File.Delete(zipPath);
                }

                // [DRM 우회] SaveAsFile() 대신 PropertyAccessor로 메모리의 바이트 데이터를 직접 읽어옵니다.
                const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";
                byte[] attachmentData = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN);
                
                // 직접 ZIP 아카이브를 생성하거나 기존 아카이브에 추가
                ZipArchiveMode mode = File.Exists(zipPath) ? ZipArchiveMode.Update : ZipArchiveMode.Create;
                using (ZipArchive zip = ZipFile.Open(zipPath, mode))
                {
                    // Update 모드일 경우 이미 동일한 파일이 있다면 중복 방지를 위해 삭제
                    if (mode == ZipArchiveMode.Update)
                    {
                        var existingEntry = zip.GetEntry(targetName);
                        if (existingEntry != null)
                        {
                            existingEntry.Delete();
                        }
                    }

                    var entry = zip.CreateEntry(targetName, CompressionLevel.Optimal);
                    using (var entryStream = entry.Open())
                    {
                        entryStream.Write(attachmentData, 0, attachmentData.Length);
                    }
                }

                zipFilePath = zipPath;
                return true;
            }
            catch (COMException comEx)
            {
                error = $"첨부 파일 메모리 추출 실패 (0x{comEx.HResult:X8}): {comEx.Message}";
                return false;
            }
            catch (Exception ex)
            {
                error = $"파일 압축 저장 예외: {ex.Message}";
                return false;
            }
        }

        // ─────────────────────────────────────────────────────────
        // 이하 클래식 Outlook COM 전용 메서드들
        // ─────────────────────────────────────────────────────────

        private void ReportStatus(string message) => OnStatusUpdate?.Invoke(message);

        private bool OpenFileWithDefaultApp(out string error)
        {
            error = string.Empty;
            string ext = Path.GetExtension(SourceFile).ToLower();

            try
            {
                if (ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".xlsb" || ext == ".csv")
                {
                    Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                    if (excelType != null)
                    {
                        _excelApp = Activator.CreateInstance(excelType);
                        if (_excelApp != null)
                        {
                            _excelApp.Visible = true;
                            dynamic workbooks = _excelApp.Workbooks;
                            _excelWorkbook = workbooks.Open(SourceFile);
                            return true;
                        }
                    }
                }
                else if (ext == ".docx" || ext == ".doc")
                {
                    Type? wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType != null)
                    {
                        _wordApp = Activator.CreateInstance(wordType);
                        if (_wordApp != null)
                        {
                            _wordApp.Visible = true;
                            dynamic documents = _wordApp.Documents;
                            _wordDocument = documents.Open(SourceFile);
                            return true;
                        }
                    }
                }
                else if (ext == ".pptx" || ext == ".ppt")
                {
                    Type? pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (pptType != null)
                    {
                        _pptApp = Activator.CreateInstance(pptType);
                        if (_pptApp != null)
                        {
                            _pptApp.Visible = true;
                            dynamic presentations = _pptApp.Presentations;
                            _pptPresentation = presentations.Open(SourceFile);
                            return true;
                        }
                    }
                }

                // COM을 통한 제어가 불가능한 기타 파일이거나, COM 초기화 실패 시 기존 방식(폴백) 사용
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = SourceFile,
                    UseShellExecute = true,
                    Verb = "open"
                };
                // 참조는 저장해두지만, 앞으로는 Kill()로 강제종료 하지 않음.
                _openedProcess = Process.Start(psi);
                return true;
            }
            catch (Exception ex)
            {
                error = $"파일 열기 실패: {ex.Message}\n파일: {SourceFile}";
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



        /// <summary>
        /// 자동화 완료/실패 후 모든 리소스 안전하게 정리
        /// 순서: 메일닫기 → Outlook종료 → 에디터 COM 문서 닫기 → 가비지컬렉션
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

            // 3) 에디터(Excel, Word, PPT) COM을 사용해 열었던 '문서만' 닫기 (저장 안 함)
            if (_excelWorkbook != null)
            {
                try { _excelWorkbook.Close(false); } catch { }
                try { Marshal.ReleaseComObject(_excelWorkbook); } catch { }
                _excelWorkbook = null;
            }
            if (_excelApp != null)
            {
                // 열려있는 다른 통합문서가 없으면 애플리케이션도 종료
                try { if (_excelApp.Workbooks.Count == 0) _excelApp.Quit(); } catch { }
                try { Marshal.ReleaseComObject(_excelApp); } catch { }
                _excelApp = null;
            }

            if (_wordDocument != null)
            {
                try { _wordDocument.Close(false); } catch { }
                try { Marshal.ReleaseComObject(_wordDocument); } catch { }
                _wordDocument = null;
            }
            if (_wordApp != null)
            {
                try { if (_wordApp.Documents.Count == 0) _wordApp.Quit(); } catch { }
                try { Marshal.ReleaseComObject(_wordApp); } catch { }
                _wordApp = null;
            }

            if (_pptPresentation != null)
            {
                try { _pptPresentation.Close(); } catch { }
                try { Marshal.ReleaseComObject(_pptPresentation); } catch { }
                _pptPresentation = null;
            }
            if (_pptApp != null)
            {
                try { if (_pptApp.Presentations.Count == 0) _pptApp.Quit(); } catch { }
                try { Marshal.ReleaseComObject(_pptApp); } catch { }
                _pptApp = null;
            }

            // 4) 폴백 프로세스는 강제 종료(Kill)하지 않고 참조만 정리합니다.
            //    (사용자가 다른 파일을 작업 중일 때 강제종료되는 위험 방지)
            if (_openedProcess != null)
            {
                try { _openedProcess.Dispose(); } catch { }
                _openedProcess = null;
            }

            // 5) GC로 미수집 COM 객체 정리 (권장 패턴)
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}