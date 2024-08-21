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

namespace ActuLight.Pages
{
    public partial class OutputPage : Page
    {
        private ModelWriter _modelWriter;
        private DispatcherTimer _timer;
        private DateTime _startTime;
        private bool _isCancelled = false;
        private string _outputFolderPath;

        public OutputPage()
        {
            InitializeComponent();
        }

        private void LoadData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var filePage = ((MainWindow)Application.Current.MainWindow).pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage?.excelData == null || !filePage.excelData.ContainsKey("out"))
                {
                    throw new Exception("출력 데이터를 찾을 수 없습니다.");
                }

                _outputFolderPath = Path.GetDirectoryName(FilePage.SelectedFilePath);
                _modelWriter = new ModelWriter(App.ModelEngine, new DataExpander(App.ModelEngine.ModelPointInfo.Types, App.ModelEngine.ModelPointInfo.Headers));
                _modelWriter.LoadTableData(filePage.excelData["out"]);
                PopulateTabControl();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 로드 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PopulateTabControl()
        {
            OutputTabControl.Items.Clear();

            foreach (var tableEntry in _modelWriter.CompiledExpressions)
            {
                var tabItem = new TabItem { Header = tableEntry.Key };
                var dataGrid = new DataGrid
                {
                    AutoGenerateColumns = false,
                    IsReadOnly = true
                };

                dataGrid.Columns.Add(new DataGridTextColumn { Header = "ColumnName", Binding = new System.Windows.Data.Binding("Key") });
                dataGrid.Columns.Add(new DataGridTextColumn { Header = "Value", Binding = new System.Windows.Data.Binding("Value") });

                var items = new List<KeyValuePair<string, string>>();
                foreach (var columnEntry in tableEntry.Value)
                {
                    items.Add(new KeyValuePair<string, string>(columnEntry.Key, columnEntry.Value.Text.ToString()));
                }
                dataGrid.ItemsSource = items;

                tabItem.Content = dataGrid;
                OutputTabControl.Items.Add(tabItem);
            }
        }

        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Column.Header.ToString() == "Value")
            {
                var editedItem = (KeyValuePair<string, string>)e.Row.Item;
                var newValue = (e.EditingElement as TextBox).Text;

                try
                {
                    var tabItem = (TabItem)OutputTabControl.SelectedItem;
                    var tableName = tabItem.Header.ToString();
                    var columnName = editedItem.Key;

                    var transformedValue = ModelEngine.TransformText(newValue, "DummyModel");
                    _modelWriter.CompiledExpressions[tableName][columnName] = App.ModelEngine.Context.CompileDynamic(transformedValue);

                    e.Row.Background = Brushes.White;
                }
                catch
                {
                    e.Row.Background = Brushes.Red;
                }
            }
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            _startTime = DateTime.Now;
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            _timer.Tick += Timer_Tick;
            _timer.Start();

            _modelWriter.TotalPoints = App.ModelEngine.ModelPoints.Count;
            _modelWriter.CompletedPoints = 0;
            _modelWriter.ErrorPoints = 0;
            _isCancelled = false;

            await Task.Run(() => WriteResults());
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            _isCancelled = true;
            UpdateProgressRichTextBox();
        }

        private void WriteResults()
        {
            foreach (var tableName in _modelWriter.CompiledExpressions.Keys)
            {
                if (_isCancelled) break;
                _modelWriter.WriteResults(_outputFolderPath, tableName);
            }

            Dispatcher.Invoke(() =>
            {
                _timer.Stop();
                UpdateProgressRichTextBox();
                MessageBox.Show(_isCancelled ? "작업이 취소되었습니다." : "작업이 완료되었습니다.");
            });
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan elapsed = DateTime.Now - _startTime;
            Dispatcher.Invoke(() =>
            {
                TimeLabel.Text = $"경과 시간: {elapsed:hh\\:mm\\:ss}";
                TotalPointsLabel.Text = $"전체 건수: {_modelWriter.TotalPoints}";
                CompletedPointsLabel.Text = $"완료 건수: {_modelWriter.CompletedPoints}";
                ErrorPointsLabel.Text = $"오류 건수: {_modelWriter.ErrorPoints}";
                ProgressLabel.Text = $"진행 상황: {((_modelWriter.CompletedPoints + _modelWriter.ErrorPoints) * 100.0 / _modelWriter.TotalPoints):F2}%";
            });
        }

        private void UpdateProgressRichTextBox()
        {
            string progressInfo = $"{ProgressLabel.Text}\n{TimeLabel.Text}\n{TotalPointsLabel.Text}\n{CompletedPointsLabel.Text}\n{ErrorPointsLabel.Text}\n\n";
            ProgressRichTextBox.AppendText(progressInfo);
            ProgressRichTextBox.ScrollToEnd();
        }

    }
}