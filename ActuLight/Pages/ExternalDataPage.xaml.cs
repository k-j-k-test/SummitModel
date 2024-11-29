using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ActuLiteModel;
using Flee.PublicTypes;
using System.ComponentModel;
using System.Windows.Threading;

namespace ActuLight.Pages
{
    public partial class ExternalDataPage : Page
    {
        private ExternalDataProcessor _processor;
        private DispatcherTimer _progressTimer;
        private LTFReader _currentReader;

        public ExternalDataPage()
        {
            InitializeComponent();
            InitializeProgressTimer();
        }

        private void InitializeProgressTimer()
        {
            _progressTimer = new DispatcherTimer();
            _progressTimer.Interval = TimeSpan.FromMilliseconds(100);
            _progressTimer.Tick += ProgressTimer_Tick;
        }

        private void ProgressTimer_Tick(object sender, EventArgs e)
        {
            if (_currentReader != null)
            {
                IndexingProgress.Value = _currentReader.Progress * 100;
            }
        }

        public async Task LoadDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;

            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;

                if (filePage?.excelData?.ContainsKey("exdata") == true)
                {
                    var exdata = filePage.excelData["exdata"];
                    _processor = new ExternalDataProcessor(FilePage.SelectedFilePath, exdata);

                    KeyDataGrid.ItemsSource = _processor.KeyConfigs.Values;
                    VariableDataGrid.ItemsSource = _processor.Configs.Values;

                    IndexingButton.IsEnabled = true;
                    UpdateIndexingStatus("준비됨");
                }
                else
                {
                    MessageBox.Show("외부 데이터 설정을 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 로드 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        private async void IndexingButton_Click(object sender, RoutedEventArgs e)
        {
            if (_processor == null) return;

            try
            {
                IndexingButton.IsEnabled = false;
                _progressTimer.Start();

                foreach (string fileName in _processor.KeyConfigs.Keys)
                {
                    UpdateIndexingStatus($"인덱싱 중: {fileName}");
                    var config = _processor.KeyConfigs[fileName];
                    config.IndexingStatus = "진행중";

                    try
                    {
                        await Task.Run(() =>
                        {
                            _currentReader = _processor.InitializeReader(fileName);
                        });
                        config.IndexingStatus = "완료";
                    }
                    catch (Exception ex)
                    {
                        config.IndexingStatus = "오류";
                        MessageBox.Show($"{fileName} 파일 인덱싱 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                        UpdateIndexingStatus("오류 발생");
                        return;
                    }
                }

                UpdateIndexingStatus("인덱싱 완료");
                MessageBox.Show("모든 파일의 인덱싱이 완료되었습니다.", "완료", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                _progressTimer.Stop();
                IndexingButton.IsEnabled = true;
                _currentReader = null;
                IndexingProgress.Value = 0;
            }
        }

        private void UpdateIndexingStatus(string status)
        {
            Dispatcher.Invoke(() => IndexingStatusText.Text = status);
        }
    }

    public class ExternalDataProcessor
    {
        private readonly ExpressionContext _context;
        private Dictionary<string, LTFReader> _readers;

        // 공개 프로퍼티로 변경
        public string BasePath { get; }
        public Dictionary<string, ExternalDataConfig> Configs { get; }
        public Dictionary<string, ExternalDataConfig> KeyConfigs { get; }

        public ExternalDataProcessor(string excelFilePath, List<List<object>> exdata)
        {
            _context = new ExpressionContext();
            _context.Imports.AddType(typeof(TextParser));
            BasePath = Path.Combine(
                Path.GetDirectoryName(excelFilePath),
                $"Data_{Path.GetFileNameWithoutExtension(excelFilePath)}",
                "ExternalData"
            );

            // 디렉토리가 없으면 생성
            if (!Directory.Exists(BasePath))
            {
                Directory.CreateDirectory(BasePath);
            }

            var allConfigs = exdata.Skip(1)
                                 .Where(row => row.Count >= 3 && row[0] != null)
                                 .Select(row => new ExternalDataConfig
                                 {
                                     FileName = row[0].ToString(),
                                     Name = row[1]?.ToString(),
                                     Selector = _context.CompileGeneric<string>(row[2]?.ToString())
                                 });

            Configs = allConfigs.Where(c => c.Name?.ToLower() != "key")
                               .ToDictionary(c => c.Name);

            KeyConfigs = allConfigs.Where(c => c.Name?.ToLower() == "key")
                                 .ToDictionary(c => c.FileName);

            _readers = new Dictionary<string, LTFReader>();
        }

        // InitializeReaders를 public 메서드로 분리
        public LTFReader InitializeReader(string fileName)
        {
            if (!KeyConfigs.TryGetValue(fileName, out var config))
            {
                throw new KeyNotFoundException($"Key not found for file: {fileName}");
            }

            string filePath = Path.Combine(BasePath, fileName);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"External data file not found: {filePath}");
            }

            var reader = new LTFReader(filePath, line =>
            {
                TextParser.SetLine(line);
                try
                {
                    return config.Selector.Evaluate();
                }
                catch
                {
                    return null;
                }
            });

            reader.LoadIndex();

            _readers[fileName] = reader;
            return reader;
        }

        public string GetValue(string name, string key)
        {
            if (!Configs.TryGetValue(name, out var config))
            {
                throw new KeyNotFoundException($"Variable {name} not found in configurations");
            }

            var reader = _readers[config.FileName];
            var line = reader.GetLine(key);

            if (line == null)
            {
                return string.Empty;
            }

            try
            {
                TextParser.SetLine(line);
                return config.Selector.Evaluate() ?? string.Empty;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error processing data for {name}: {ex.Message}");
            }
        }

        public IEnumerable<string> GetKeys(string name)
        {
            if (!KeyConfigs.TryGetValue(name, out var config))
            {
                return Enumerable.Empty<string>();
            }

            return _readers[config.FileName].GetUniqueKeys();
        }
    }

    public class ExternalDataConfig : INotifyPropertyChanged
    {
        public string FileName { get; set; }
        public string Name { get; set; }
        public IGenericExpression<string> Selector { get; set; }

        private string _indexingStatus = "미완료";
        public string IndexingStatus
        {
            get => _indexingStatus;
            set
            {
                if (_indexingStatus != value)
                {
                    _indexingStatus = value;
                    OnPropertyChanged(nameof(IndexingStatus));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}