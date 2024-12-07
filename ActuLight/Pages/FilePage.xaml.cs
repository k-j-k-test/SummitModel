using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using Microsoft.Win32;
using ActuLiteModel;
using System.Windows.Threading;
using System.Windows.Media;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Windows.Documents;
using Newtonsoft.Json;

namespace ActuLight.Pages
{
    public partial class FilePage : Page
    {
        private string latestVersion;
        private string downloadUrl;

        private FileSystemWatcher fileWatcher;
        private readonly string[] monitoringFiles = { "mp.txt", "assum.txt", "exp.txt", "out.txt" };

        private const string RecentFilesPath = "recentFiles.json";
        private ObservableCollection<RecentFile> recentFiles = new ObservableCollection<RecentFile>();
        public string currentFilePath;
        public Dictionary<string, List<List<object>>> excelData;

        public static string SelectedFolderPath { get; private set; }
        public static bool IsAutoSync = true;

        public FilePage()
        {
            InitializeComponent();
            LoadRecentFiles();
            InitializeFileWatcher();
            RecentFilesList.ItemsSource = recentFiles;
            CheckForUpdates();
        }

        private void InitializeFileWatcher()
        {
            fileWatcher = new FileSystemWatcher();
            fileWatcher.NotifyFilter = NotifyFilters.LastWrite;
            fileWatcher.Changed += OnFileChanged;
            fileWatcher.EnableRaisingEvents = false; // 초기에는 비활성화
        }

        private async void NewProjectButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "새 프로젝트 생성",
                    Filter = "Project Directory|*.smt",
                    DefaultExt = ".smt",
                    FileName = "새 프로젝트"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string projectName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                    string parentPath = Path.GetDirectoryName(saveFileDialog.FileName);
                    string projectPath = Path.Combine(parentPath, projectName);

                    // Create project directory
                    Directory.CreateDirectory(projectPath);

                    // Create subdirectories
                    string[] subDirectories = { "Inputs", "ExternalData", "Outputs", "Samples", "Scripts" };
                    foreach (string dir in subDirectories)
                    {
                        Directory.CreateDirectory(Path.Combine(projectPath, dir));
                    }

                    // Copy default txt files from Resources to ExcelData folder
                    string resourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
                    string excelDataPath = Path.Combine(projectPath, "Inputs");

                    foreach (string file in monitoringFiles)
                    {
                        string sourcePath = Path.Combine(resourcePath, file);
                        string destPath = Path.Combine(excelDataPath, file);

                        if (File.Exists(sourcePath))
                        {
                            File.Copy(sourcePath, destPath);
                        }
                        else
                        {
                            MessageBox.Show($"Resources 폴더에서 {file} 파일을 찾을 수 없습니다.", "경고", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }

                    // Create shortcut file (.smt)
                    string executablePath = Process.GetCurrentProcess().MainModule.FileName;
                    string shortcutPath = Path.Combine(projectPath, $"{projectName}.smt");

                    using (StreamWriter writer = new StreamWriter(shortcutPath))
                    {
                        writer.WriteLine(executablePath);
                    }

                    // Copy Excel template
                    string templatePath = Path.Combine(resourcePath, "ExcelData_Templete.xlsm");
                    string destinationPath = Path.Combine(projectPath, $"{projectName}.xlsm");

                    if (File.Exists(templatePath))
                    {
                        File.Copy(templatePath, destinationPath);
                        MessageBox.Show("프로젝트가 성공적으로 생성되었습니다.", "성공", MessageBoxButton.OK, MessageBoxImage.Information);

                        // Load the newly created Excel file
                        await LoadDataAsync(projectPath);
                    }
                    else
                    {
                        MessageBox.Show("Excel 템플릿 파일을 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"프로젝트 생성 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "smt files (*.smt)|*.smt|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string projectPath = Path.GetDirectoryName(openFileDialog.FileName);
                await LoadDataAsync(projectPath);
            }
        }

        private async void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                await LoadDataAsync(openFileDialog.FileName);
            }
        }

        private async void OpenSelectedFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (RecentFilesList.SelectedItem is RecentFile selectedFolder)
            {
                try
                {
                    LoadingOverlay.Visibility = Visibility.Visible;

                    // FilePage의 LoadDataAsync 메서드 호출
                    await LoadDataAsync(selectedFolder.Path);

                    // MainWindow 인스턴스 가져오기
                    var mainWindow = Application.Current.MainWindow as MainWindow;
                    if (mainWindow == null)
                    {
                        throw new InvalidOperationException("MainWindow를 찾을 수 없습니다.");
                    }

                    // ModelPointPage의 LoadDataAsync 메서드 호출
                    string ModelPointPageDir = "Pages/ModelPointPage.xaml";
                    if (mainWindow.pageCache.TryGetValue(ModelPointPageDir, out Page modelPointPage))
                    {
                        await (modelPointPage as ModelPointPage).LoadDataAsync();
                    }
                    else
                    {
                        mainWindow.pageCache[ModelPointPageDir] = new ModelPointPage();
                        await (mainWindow.pageCache[ModelPointPageDir] as ModelPointPage).LoadDataAsync();
                    }

                    // AssumptionPage의 LoadDataAsync 메서드 호출
                    string AssumptionPageDir = "Pages/AssumptionPage.xaml";
                    if (mainWindow.pageCache.TryGetValue(AssumptionPageDir, out Page assumptionPage))
                    {
                        await (assumptionPage as AssumptionPage).LoadDataAsync();
                    }
                    else
                    {
                        mainWindow.pageCache[AssumptionPageDir] = new AssumptionPage();
                        await (mainWindow.pageCache[AssumptionPageDir] as AssumptionPage).LoadDataAsync();
                    }

                    // OutputPage의 LoadDataAsync 메서드 호출
                    string OutputPageDir = "Pages/OutputPage.xaml";
                    if (mainWindow.pageCache.TryGetValue(OutputPageDir, out Page outputPage))
                    {
                        (outputPage as OutputPage).LoadData_Click(null, null);
                    }
                    else
                    {
                        mainWindow.pageCache[OutputPageDir] = new OutputPage();
                        (mainWindow.pageCache[OutputPageDir] as OutputPage).LoadData_Click(null, null);
                    }

                    // DataProcessingPage의 LoadDataAsync 메서드 호출
                    string DataProcessingPageDir = "Pages/DataProcessingPage.xaml";
                    if (mainWindow.pageCache.TryGetValue(DataProcessingPageDir, out Page dataProcessingPage))
                    {
                        (dataProcessingPage as DataProcessingPage).LoadExternalButton_Click(null, null);
                    }
                    else
                    {
                        mainWindow.pageCache[DataProcessingPageDir] = new DataProcessingPage();
                        (mainWindow.pageCache[DataProcessingPageDir] as DataProcessingPage).LoadExternalButton_Click(null, null);
                    }


                    //SpreadSheetPage의 LoadDataAsync 메서드 호출
                    string SpreadSheetPageDir = "Pages/SpreadSheetPage.xaml";
                    if (!mainWindow.pageCache.TryGetValue(SpreadSheetPageDir, out Page spreadSheetPage))
                    {
                        mainWindow.pageCache[SpreadSheetPageDir] = new SpreadSheetPage();
                        spreadSheetPage = mainWindow.pageCache[SpreadSheetPageDir];
                    }

                    string fileName = Path.GetFileNameWithoutExtension(selectedFolder.Path);
                    string autoScriptPath = Path.Combine(selectedFolder.Path, "Scripts", $"{fileName}_scripts_auto1.json");

                    if (!File.Exists(autoScriptPath))
                    {
                        string resourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "NewProject_scripts_auto1.json");
                        if (File.Exists(resourcePath))
                        {
                            File.Copy(resourcePath, autoScriptPath);
                        }
                    }

                    if (File.Exists(autoScriptPath))
                    {
                        await (spreadSheetPage as SpreadSheetPage).LoadDataAsync(autoScriptPath);
                    }

                    MessageBox.Show("모든 데이터가 성공적으로 로드되었습니다.", "성공", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"데이터 로드 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void DeleteSelectedFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (RecentFilesList.SelectedItem is RecentFile selectedFile)
            {
                recentFiles.Remove(selectedFile);
                SaveRecentFiles();
            }
        }

        private async Task LoadDataAsync(string projectPath)
        {
            LoadingOverlay.Visibility = Visibility.Visible;

            try
            {
                UpdateStatusMessage("데이터 로딩 중...", true);

                // FileWatcher 설정
                if (fileWatcher != null)
                {
                    fileWatcher.Path = Path.Combine(projectPath, "Inputs");
                    fileWatcher.EnableRaisingEvents = true;
                }

                string excelDataPath = Path.Combine(projectPath, "Inputs");
                Dictionary<string, List<List<object>>> loadedData = new Dictionary<string, List<List<object>>>();

                await Task.Run(() =>
                {
                    if (Directory.Exists(excelDataPath) && Directory.GetFiles(excelDataPath, "*.txt").Any())
                    {
                        foreach (string file in Directory.GetFiles(excelDataPath, "*.txt"))
                        {
                            string sheetName = Path.GetFileNameWithoutExtension(file);
                            List<List<object>> sheetData = new List<List<object>>();

                            string[] lines = File.ReadAllLines(file);
                            foreach (string line in lines)
                            {
                                List<object> rowData = line.Split('\t')
                                                         .Cast<object>()
                                                         .ToList();
                                sheetData.Add(rowData);
                            }

                            loadedData[sheetName] = sheetData;
                        }
                    }
                });

                excelData = loadedData;
                currentFilePath = projectPath;
                SelectedFolderPath = projectPath;

                UpdateStatusMessage($"데이터가 성공적으로 로드되었습니다. 경로: {projectPath}", true);
                UpdateDataSummary();
                AddToRecentFiles(projectPath);
                UpdateApplicationTitle(projectPath);
            }
            catch (Exception ex)
            {
                UpdateStatusMessage($"데이터 로드 중 오류 발생: {ex.Message}", false);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private async void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            if (!IsAutoSync) return;

            // 모니터링 대상 파일인지 확인
            string fileName = Path.GetFileName(e.Name);
            if (!monitoringFiles.Contains(fileName)) return;

            // 플래그 파일 경로
            string flagFilePath = Path.Combine(Path.GetDirectoryName(e.FullPath), "flag.writing");

            // 플래그 파일이 존재하면 쓰기가 진행 중이므로 대기
            while (File.Exists(flagFilePath))
            {
                await Task.Delay(100);
            }

            await DebouncerAsync.Debounce(fileName, 500, async () =>
            {
                // UI 스레드에서 실행
                await Dispatcher.InvokeAsync(async () =>
                {
                    try
                    {
                        await LoadDataAsync(currentFilePath);

                        // MainWindow 인스턴스 가져오기
                        var mainWindow = Application.Current.MainWindow as MainWindow;
                        if (mainWindow == null)
                        {
                            throw new InvalidOperationException("MainWindow를 찾을 수 없습니다.");
                        }

                        // 변경된 파일에 따른 페이지 업데이트
                        Dictionary<string, (string pageDir, Type pageType)> pageMapping = new Dictionary<string, (string, Type)>
                        {
                            { "mp.txt", ("Pages/ModelPointPage.xaml", typeof(ModelPointPage)) },
                            { "assum.txt", ("Pages/AssumptionPage.xaml", typeof(AssumptionPage)) },
                            { "exp.txt", ("Pages/AssumptionPage.xaml", typeof(AssumptionPage)) },
                            { "out.txt", ("Pages/OutputPage.xaml", typeof(OutputPage)) }
                       };

                        // 해당 파일에 대한 페이지 매핑이 있는 경우 업데이트 실행
                        if (pageMapping.TryGetValue(fileName, out var pageInfo))
                        {
                            if (mainWindow.pageCache.TryGetValue(pageInfo.pageDir, out Page page))
                            {
                                if (pageInfo.pageType == typeof(ModelPointPage))
                                {
                                    await (page as ModelPointPage).LoadDataAsync();
                                }
                                if (pageInfo.pageType == typeof(AssumptionPage) && fileName == "assum.txt")
                                {
                                    await (page as AssumptionPage).LoadAssumptionDataAsync();
                                }
                                if (pageInfo.pageType == typeof(AssumptionPage) && fileName == "exp.txt")
                                {
                                    await (page as AssumptionPage).LoadExpenseDataAsync();
                                }
                                if (pageInfo.pageType == typeof(OutputPage))
                                {
                                    (page as OutputPage).LoadData_Click(null, null);
                                }
                            }
                        }

                        // SpreadSheetPage는 항상 업데이트
                        string SpreadsheetPageDir = "Pages/SpreadSheetPage.xaml";
                        if (mainWindow.pageCache.TryGetValue(SpreadsheetPageDir, out Page spreadsheetPage))
                        {
                            (spreadsheetPage as SpreadSheetPage).UpdateInvokes();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"파일 변경 감지 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                });
            });
        }

        private async void UpdateExcelData()
        {
            LoadingOverlay.Visibility = Visibility.Visible;

            // FilePage의 LoadDataAsync 메서드 호출
            await LoadDataAsync(currentFilePath);

            // MainWindow 인스턴스 가져오기
            var mainWindow = Application.Current.MainWindow as MainWindow;
            if (mainWindow == null)
            {
                throw new InvalidOperationException("MainWindow를 찾을 수 없습니다.");
            }

            // ModelPointPage의 LoadDataAsync 메서드 호출
            string ModelPointPageDir = "Pages/ModelPointPage.xaml";
            if (mainWindow.pageCache.TryGetValue(ModelPointPageDir, out Page modelPointPage))
            {
                await(modelPointPage as ModelPointPage).LoadDataAsync();
            }
            else
            {
                mainWindow.pageCache[ModelPointPageDir] = new ModelPointPage();
                await(mainWindow.pageCache[ModelPointPageDir] as ModelPointPage).LoadDataAsync();
            }

            // AssumptionPage의 LoadDataAsync 메서드 호출
            string AssumptionPageDir = "Pages/AssumptionPage.xaml";
            if (mainWindow.pageCache.TryGetValue(AssumptionPageDir, out Page assumptionPage))
            {
                await(assumptionPage as AssumptionPage).LoadDataAsync();
            }
            else
            {
                mainWindow.pageCache[AssumptionPageDir] = new AssumptionPage();
                await(mainWindow.pageCache[AssumptionPageDir] as AssumptionPage).LoadDataAsync();
            }

            // OutputPage의 LoadDataAsync 메서드 호출
            string OutputPageDir = "Pages/OutputPage.xaml";
            if (mainWindow.pageCache.TryGetValue(OutputPageDir, out Page outputPage))
            {
                (outputPage as OutputPage).LoadData_Click(null, null);
            }
            else
            {
                mainWindow.pageCache[OutputPageDir] = new OutputPage();
                (mainWindow.pageCache[OutputPageDir] as OutputPage).LoadData_Click(null, null);
            }

            // SpreadSheetPage의 LoadDataAsync 메서드 호출
            string SpreadsheetPageDir = "Pages/SpreadSheetPage.xaml";
            if (mainWindow.pageCache.TryGetValue(SpreadsheetPageDir, out Page spreadsheetPage))
            {
                (spreadsheetPage as SpreadSheetPage).UpdateInvokes();
            }
        }

        private void UpdateApplicationTitle(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            Application.Current.MainWindow.Title = $"ActuLight - {fileName}";
        }

        private void AddToRecentFiles(string filePath)
        {
            var existingFile = recentFiles.FirstOrDefault(f => f.Path == filePath);
            if (existingFile != null)
            {
                recentFiles.Remove(existingFile);
                recentFiles.Insert(0, new RecentFile(filePath));
            }
            else
            {
                recentFiles.Insert(0, new RecentFile(filePath));
            }

            if (recentFiles.Count > 100)
            {
                recentFiles.RemoveAt(recentFiles.Count - 1);
            }
            SaveRecentFiles();
        }

        private void SaveRecentFiles()
        {
            var json = JsonConvert.SerializeObject(recentFiles);
            File.WriteAllText(RecentFilesPath, json);
        }

        private void LoadRecentFiles()
        {
            if (File.Exists(RecentFilesPath))
            {
                var json = File.ReadAllText(RecentFilesPath);
                var loadedFiles = JsonConvert.DeserializeObject<List<RecentFile>>(json);
                if (loadedFiles != null)
                {
                    recentFiles.Clear();
                    foreach (var file in loadedFiles)
                    {
                        recentFiles.Add(file);
                    }
                }
            }
        }

        private void UpdateStatusMessage(string message, bool isSuccess)
        {
            Dispatcher.Invoke(() =>
            {
                StatusMessage.Inlines.Clear();
                if (isSuccess)
                {
                    StatusMessage.Inlines.Add(new Run(message));
                }
                else
                {
                    StatusMessage.Inlines.Add(new Run(message) { Foreground = Brushes.Red });
                }
            });
        }

        private void UpdateDataSummary()
        {
            if (excelData != null)
            {
                StringBuilder summaryBuilder = new StringBuilder();
                StringBuilder warningBuilder = new StringBuilder();
                bool needsOptimization = false;

                foreach (var sheet in excelData)
                {
                    int rowCount = sheet.Value.Count;

                    if (sheet.Value.Any())
                    {
                        int columnCount = sheet.Value[0].Count;

                        summaryBuilder.AppendLine($"{sheet.Key}: {rowCount}행, {columnCount}열");

                        if (rowCount >= 100000 || columnCount >= 5000)
                        {
                            needsOptimization = true;
                            warningBuilder.AppendLine($"경고: {sheet.Key}의 데이터 크기가 매우 큽니다. 최적화가 필요할 수 있습니다.");
                        }
                    }
                }

                if (needsOptimization)
                {
                    warningBuilder.AppendLine("\n주의: 일부 데이터의 크기가 매우 큽니다. 애플리케이션의 성능 최적화가 필요할 수 있습니다.");
                }

                DataSummary.Inlines.Clear();
                DataSummary.Inlines.Add(new Run("데이터 요약:\n" + summaryBuilder.ToString()));

                if (warningBuilder.Length > 0)
                {
                    DataSummary.Inlines.Add(new Run(warningBuilder.ToString()) { Foreground = Brushes.Red });
                }
            }
            else
            {
                DataSummary.Inlines.Clear();
                DataSummary.Inlines.Add(new Run("데이터가 로드되지 않았습니다."));
            }
        }

        private void UpdateMemoryUsage()
        {
            Process currentProcess = Process.GetCurrentProcess();
            long memoryUsage = currentProcess.WorkingSet64 / (1024 * 1024); // Convert to MB
            MemoryUsage.Text = $"현재 사용 중인 메모리: {memoryUsage} MB";
        }

        private async void CheckForUpdates()
        {
            try
            {
                (latestVersion, downloadUrl) = await VersionChecker.GetLatestVersionInfo();
                if (latestVersion != null)
                {
                    if (VersionChecker.IsUpdateAvailable(App.CurrentVersion, latestVersion))
                    {
                        VersionInfoTextBlock.Text = $"Current version: {App.CurrentVersion}, Latest version: {latestVersion}";
                        UpdateLinkTextBlock.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        VersionInfoTextBlock.Text = $"Current version: {App.CurrentVersion}, Latest version: {latestVersion}";
                        UpdateLinkTextBlock.Visibility = Visibility.Collapsed;
                    }
                }
            }
            catch (Exception ex)
            {
                VersionInfoTextBlock.Text = $"Failed to check for updates: {ex.Message}";
            }
        }

        private async void UpdateLink_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Do you want to update to the latest version?", "Update Available", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    await UpdateHelper.DownloadAndExtractUpdate(downloadUrl);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Update failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    public class RecentFile
    {
        public string Path { get; set; }
        public string FileName { get; set; }
        public string LastUsed { get; set; }

        public RecentFile(string path)
        {
            Path = path;
            FileName = System.IO.Path.GetFileName(path);
            LastUsed = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        public RecentFile() { } // JSON 역직렬화용 생성자
    }
}