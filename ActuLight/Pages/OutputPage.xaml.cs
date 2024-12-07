using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ActuLiteModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.IO;
using System.Linq;

namespace ActuLight.Pages
{
    public partial class OutputPage : Page
    {
        private ModelWriter _modelWriter;
        private DispatcherTimer _timer;
        private DateTime _startTime;
        private string _outputFolderPath;
        private string _selectedDelimiter = "\t";
        private Dictionary<string, DataGrid> tableGrids;
        private Dictionary<string, List<Input_output>> tableNameData;
        private bool isSearching = false;

        public OutputPage()
        {
            InitializeComponent();
            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = false;
            tableGrids = new Dictionary<string, DataGrid>();
            tableNameData = new Dictionary<string, List<Input_output>>();
        }

        public void LoadData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var filePage = ((MainWindow)Application.Current.MainWindow).pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage?.excelData == null || !filePage.excelData.ContainsKey("out"))
                {
                    throw new Exception("출력 데이터를 찾을 수 없습니다.");
                }

                string fileName = Path.GetFileNameWithoutExtension(FilePage.SelectedFolderPath);
                _outputFolderPath = Path.Combine(FilePage.SelectedFolderPath, "Outputs");
                if (!Directory.Exists(_outputFolderPath))
                {
                    Directory.CreateDirectory(_outputFolderPath);
                }

                App.ModelEngine.SetOutputs(filePage.excelData["out"]);
                _modelWriter = new ModelWriter(App.ModelEngine,
                    new DataExpander(App.ModelEngine.ModelPointInfo.Types, App.ModelEngine.ModelPointInfo.Headers));

                ClassifyAndCreateTabs();

                StartButton.IsEnabled = true;
                CancelButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"out 데이터 로드 중 오류가 발생했습니다. 수식을 다시 한 번 확인바랍니다.: \n{ex.Message}",
                    "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClassifyAndCreateTabs()
        {
            OutputTabControl.Items.Clear();
            tableGrids.Clear();
            tableNameData.Clear();

            // 테이블 타입별로 데이터 그룹화
            var groupedData = App.ModelEngine.Outputs
                .SelectMany(kv => kv.Value)
                .GroupBy(output => output.Table)
                .ToDictionary(
                    group => group.Key,
                    group => group.ToList()
                );

            // 그룹화된 데이터로 탭 생성
            foreach (var group in groupedData)
            {
                var tableName = group.Key;
                var dataGrid = CreateDataGrid();
                tableGrids[tableName] = dataGrid;
                tableNameData[tableName] = group.Value;

                var tabItem = new TabItem
                {
                    Header = tableName,
                    Content = dataGrid
                };
                OutputTabControl.Items.Add(tabItem);

                UpdateDataGrid(dataGrid, group.Value);
            }

            if (OutputTabControl.Items.Count > 0)
            {
                OutputTabControl.SelectedIndex = 0;
            }
        }

        private DataGrid CreateDataGrid()
        {
            var dataGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserSortColumns = false
            };

            // Add columns
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "Table", Binding = new System.Windows.Data.Binding(nameof(Input_output.Table)) });
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "ProductCode", Binding = new System.Windows.Data.Binding(nameof(Input_output.ProductCode)) });
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "RiderCode", Binding = new System.Windows.Data.Binding(nameof(Input_output.RiderCode)) });
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "Value", Binding = new System.Windows.Data.Binding(nameof(Input_output.Value)) });
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "Position", Binding = new System.Windows.Data.Binding(nameof(Input_output.Position)) });
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "Range", Binding = new System.Windows.Data.Binding(nameof(Input_output.Range)) });
            dataGrid.Columns.Add(new DataGridTextColumn
            { Header = "Format", Binding = new System.Windows.Data.Binding(nameof(Input_output.Format)) });

            return dataGrid;
        }

        private void UpdateDataGrid(DataGrid dataGrid, List<Input_output> data)
        {
            try
            {
                dataGrid.ItemsSource = data;
                AutoFitColumns(dataGrid);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 그리드 업데이트 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AutoFitColumns(DataGrid dataGrid)
        {
            foreach (var column in dataGrid.Columns)
            {
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            dataGrid.UpdateLayout();
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _modelWriter.IsCanceled = false;
                _modelWriter.Delimiter = _selectedDelimiter;
                _modelWriter.StatusQueue = new Queue<string>();
                _startTime = DateTime.Now;
                _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(20) };
                _timer.Tick += Timer_Tick;
                _timer.Start();
                EnableButtons(false);
                ProgressRichTextBox.AppendText($"■ 시작 {DateTime.Now} " + "\r\n");

                var selectedTab = OutputTabControl.SelectedItem as TabItem;
                if (selectedTab != null)
                {
                    var table = selectedTab.Header.ToString();
                    var productCode = ProductCodeFilterBox.Text.Trim();
                    var riderCode = RiderCodeFilterBox.Text.Trim();
                    await Task.Run(() => _modelWriter.WriteResultsAsync(_outputFolderPath, productCode, riderCode));
                }

                _timer.Stop();

                // 남은 StatusQueue 메시지 모두 출력
                while (_modelWriter.StatusQueue.Count > 0)
                {
                    ProgressRichTextBox.AppendText("- " + _modelWriter.StatusQueue.Dequeue() + "\r\n");
                }

                // 종료 메시지 출력
                var elapsed = DateTime.Now - _startTime;
                ProgressRichTextBox.AppendText($"□ 종료 {DateTime.Now}, 걸린 시간: {elapsed:hh\\:mm\\:ss} " + "\r\n" + "\r\n");
                ProgressRichTextBox.ScrollToEnd();

                // 라벨 초기화
                TimeLabel.Text = $"경과 시간: {elapsed:hh\\:mm\\:ss\\.ff}";
                ProgressLabel.Text = "작업이 완료되었습니다.";

                EnableButtons(true);
                App.ModelEngine.SetModelPoint(); //변수 초기화
            }
            catch (Exception ex)
            {
                if (_timer != null)
                {
                    _timer.Stop();
                    // 예외 발생 시에도 라벨 초기화
                    TimeLabel.Text = "경과 시간: ";
                    ProgressLabel.Text = "";
                }
                MessageBox.Show("오류가 발생했습니다. 출력 데이터가 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            _modelWriter.IsCanceled = true;
            if (_timer != null)
                _timer.Stop();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ProgressRichTextBox.Document.Blocks.Clear();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan elapsed = DateTime.Now - _startTime;
            Dispatcher.Invoke(() =>
            {
                TimeLabel.Text = $"경과 시간: {elapsed:hh\\:mm\\:ss}";
                ProgressLabel.Text = _modelWriter.StatusMessage;

                if(_modelWriter.StatusQueue.Count > 0 )
                {
                    ProgressRichTextBox.AppendText("- " + _modelWriter.StatusQueue.Dequeue() + "\r\n");
                    ProgressRichTextBox.ScrollToEnd();
                }
            });
        }

        private void DelimiterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = DelimiterComboBox.SelectedItem as ComboBoxItem;
            if (selectedItem != null)
            {
                switch (selectedItem.Content.ToString())
                {
                    case "Tab":
                        _selectedDelimiter = "\t";
                        break;
                    case "Comma (,)":
                        _selectedDelimiter = ",";
                        break;
                    case "Semicolon (;)":
                        _selectedDelimiter = ";";
                        break;
                    case "Pipe (|)":
                        _selectedDelimiter = "|";
                        break;
                    case "Space ( )":
                        _selectedDelimiter = " ";
                        break;
                    default:
                        _selectedDelimiter = ""; // None
                        break;
                }
            }
        }

        private void EnableButtons(bool enabled)
        {
            var mainWindow = (MainWindow)Application.Current.MainWindow;
            NavigationBar navigationBar = mainWindow.GlobalNavBarControl.Content as NavigationBar;
            navigationBar.SetButtonsEnabled(enabled);

            // OutputPage의 버튼들 활성화
            StartButton.IsEnabled = enabled;
            DelimiterComboBox.IsEnabled = enabled;

            //자동저장기능 중지
            FilePage.IsAutoSync = enabled;  
        }

        private async void FilterBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isSearching) return;
            isSearching = true;

            await Task.Delay(300);

            var productCode = ProductCodeFilterBox.Text.Trim();
            var riderCode = RiderCodeFilterBox.Text.Trim();

            ApplyFilter(productCode, riderCode);

            isSearching = false;
        }

        private void ApplyFilter(string productCode, string riderCode)
        {
            if (OutputTabControl.SelectedItem is not TabItem selectedTab) return;

            var currentTable = selectedTab.Header.ToString();
            var dataGrid = tableGrids[currentTable];

            if (!tableNameData.TryGetValue(currentTable, out var currentTableData))
                return;

            var filteredData = currentTableData;


            if (!string.IsNullOrEmpty(productCode) || !string.IsNullOrEmpty(riderCode))
            {
                filteredData = currentTableData
                    .Where(item =>
                        (string.IsNullOrEmpty(productCode) ||
                         item.ProductCode.Contains(productCode) ||
                         item.ProductCode == "Base") &&
                        (string.IsNullOrEmpty(riderCode) ||
                         item.RiderCode.Contains(riderCode) ||
                         string.IsNullOrEmpty(item.RiderCode))
                    )
                    .ToList();
            }

            dataGrid.ItemsSource = filteredData;
        }
    }

    public class OutTableColumnInfo : IComparable<OutTableColumnInfo>
    {
        public string Table { get; set; }
        public string ProductCode { get; set; }
        public string RiderCode { get; set; }
        public string Value { get; set; }
        public int Position { get; set; }
        public string Range { get; set; }
        public string Format { get; set; }

        public int CompareTo(OutTableColumnInfo other)
        {
            return Position.CompareTo(other.Position);
        }
    }
}