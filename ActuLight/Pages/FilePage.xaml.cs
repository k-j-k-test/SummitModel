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

        private const string RecentFilesPath = "recentFiles.json";
        private ObservableCollection<RecentFile> recentFiles = new ObservableCollection<RecentFile>();
        public string currentFilePath;
        private DateTime currentFileLastWriteTime;
        private DispatcherTimer fileCheckTimer;
        private bool isAutoSync = false;

        public static string SelectedFilePath { get; private set; }
        public Dictionary<string, List<List<object>>> excelData;

        public FilePage()
        {
            InitializeComponent();
            LoadRecentFiles();
            InitializeFileCheckTimer();
            RecentFilesList.ItemsSource = recentFiles;
            CheckForUpdates();
        }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private async void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                await LoadExcelFileAsync(openFileDialog.FileName);
            }
        }

        private void InitializeFileCheckTimer()
        {
            fileCheckTimer = new DispatcherTimer();
            fileCheckTimer.Tick += async (s, e) =>
            {
                await CheckForFileChangesAsync();
                UpdateMemoryUsage();
            };
            fileCheckTimer.Interval = TimeSpan.FromSeconds(2);
            fileCheckTimer.Start();
        }

        private async void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                await LoadExcelFileAsync(openFileDialog.FileName);
            }
        }

        private async void OpenSelectedFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (RecentFilesList.SelectedItem is RecentFile selectedFile)
            {
                try
                {
                    LoadingOverlay.Visibility = Visibility.Visible;

                    // FilePage의 LoadExcelFileAsync 메서드 호출
                    await LoadExcelFileAsync(selectedFile.Path);

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

                    //// SpreadSheetPage의 LoadDataAsync 메서드 호출
                    //string SpreadSheetPageDir = "Pages/SpreadSheetPage.xaml";
                    //if (mainWindow.pageCache.TryGetValue(SpreadSheetPageDir, out Page spreadSheetPage))
                    //{
                    //    await (spreadSheetPage as SpreadSheetPage).LoadDataAsync();
                    //}
                    //else
                    //{
                    //    mainWindow.pageCache[SpreadSheetPageDir] = new SpreadSheetPage();
                    //    await (mainWindow.pageCache[SpreadSheetPageDir] as SpreadSheetPage).LoadDataAsync();
                    //}

                    //// OutputPage의 LoadDataAsync 메서드 호출
                    //string OutputPageDir = "Pages/OutputPage.xaml";
                    //if (mainWindow.pageCache.TryGetValue(OutputPageDir, out Page outputPage))
                    //{
                    //    (outputPage as OutputPage).LoadData_Click(null, null);
                    //}
                    //else
                    //{
                    //    mainWindow.pageCache[OutputPageDir] = new OutputPage();
                    //    (mainWindow.pageCache[OutputPageDir] as OutputPage).LoadData_Click(null, null);
                    //}

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

        private async Task LoadExcelFileAsync(string filePath)
        {
            LoadingOverlay.Visibility = Visibility.Visible;

            try
            {
                UpdateStatusMessage("엑셀 파일 로딩 중...", true);

                await Task.Run(() =>
                {
                    excelData = ExcelImporter.ImportMultipleSheets(filePath);
                });

                currentFilePath = filePath;
                currentFileLastWriteTime = File.GetLastWriteTime(filePath);

                UpdateStatusMessage($"엑셀 파일이 성공적으로 연결되었습니다. 파일 경로: {filePath}", true);
                SelectedFilePath = filePath;

                UpdateExcelSummary();

                AddToRecentFiles(filePath);
                UpdateApplicationTitle(filePath);
            }
            catch (Exception ex)
            {
                UpdateStatusMessage($"파일 로드 중 오류 발생: {ex.Message}", false);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateApplicationTitle(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            Application.Current.MainWindow.Title = $"ActuLight - {fileName}";
        }

        private void AutoSyncCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            isAutoSync = true;
        }

        private void AutoSyncCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            isAutoSync = false;
        }

        private void DeleteSelectedFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (RecentFilesList.SelectedItem is RecentFile selectedFile)
            {
                recentFiles.Remove(selectedFile);
                SaveRecentFiles();
            }
        }

        private async Task CheckForFileChangesAsync()
        {
            if (!string.IsNullOrEmpty(currentFilePath) && File.Exists(currentFilePath))
            {
                DateTime newLastWriteTime = File.GetLastWriteTime(currentFilePath);
                if (newLastWriteTime != currentFileLastWriteTime)
                {
                    UpdateStatusMessage($"주의: 파일 {currentFilePath}의 내용이 변경되었습니다.", false);
                    currentFileLastWriteTime = newLastWriteTime;

                    if (isAutoSync)
                    {
                        await LoadExcelFileAsync(currentFilePath);
                    }
                }
            }
        }

        private void AddToRecentFiles(string filePath)
        {
            var existingFile = recentFiles.FirstOrDefault(f => f.Path == filePath);
            if (existingFile != null)
            {
                recentFiles.Remove(existingFile);
            }
            recentFiles.Insert(0, new RecentFile(filePath));
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

        private void UpdateExcelSummary()
        {
            if (excelData != null)
            {
                StringBuilder summaryBuilder = new StringBuilder();
                StringBuilder warningBuilder = new StringBuilder();
                bool needsOptimization = false;

                foreach (var sheet in excelData)
                {
                    int rowCount = sheet.Value.Count;
                    int columnCount = sheet.Value[0].Count;

                    summaryBuilder.AppendLine($"{sheet.Key}: {rowCount}행, {columnCount}열");

                    if (rowCount >= 100000 || columnCount >= 5000)
                    {
                        needsOptimization = true;
                        warningBuilder.AppendLine($"경고: {sheet.Key} 시트의 크기가 매우 큽니다. 최적화가 필요할 수 있습니다.");
                    }
                }

                if (needsOptimization)
                {
                    warningBuilder.AppendLine("\n주의: 일부 시트의 크기가 매우 큽니다. 애플리케이션의 성능 최적화가 필요할 수 있습니다.");
                }

                ExcelSummary.Inlines.Clear();
                ExcelSummary.Inlines.Add(new Run("Excel 데이터 요약:\n" + summaryBuilder.ToString()));

                if (warningBuilder.Length > 0)
                {
                    ExcelSummary.Inlines.Add(new Run(warningBuilder.ToString()) { Foreground = Brushes.Red });
                }
            }
            else
            {
                ExcelSummary.Inlines.Clear();
                ExcelSummary.Inlines.Add(new Run("Excel 데이터가 로드되지 않았습니다."));
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

        // 파라미터 없는 생성자 (JSON 역직렬화를 위해 필요)
        public RecentFile() { }
    }
}