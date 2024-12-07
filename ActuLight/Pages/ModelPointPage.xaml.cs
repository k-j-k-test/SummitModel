using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Linq;
using System.Threading.Tasks;
using ActuLiteModel;
using System.Windows.Input;
using System.ComponentModel;
using System.Windows.Threading;
using System.Linq.Expressions;

namespace ActuLight.Pages
{
    public partial class ModelPointPage : Page
    {
        private DataExpander dataExpander;
        private List<string> headers;
        private List<string> types;
        private Dictionary<string, DataGrid> tableGrids;
        private Dictionary<string, List<List<object>>> tableNameData;
        public List<object> SelectedData { get; private set; }
        private Window expandedDataWindow;
        private bool isSearching = false;

        public ModelPointPage()
        {
            InitializeComponent();
            tableGrids = new Dictionary<string, DataGrid>();
            tableNameData = new Dictionary<string, List<List<object>>>();
        }

        private async void LoadData_Click(object sender, RoutedEventArgs e)
        {
            await LoadDataAsync();
        }

        public async Task LoadDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;

            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;

                if (filePage != null && filePage.excelData != null && filePage.excelData.ContainsKey("mp"))
                {
                    var mpData = filePage.excelData["mp"];
                    types = mpData[0].Select(t => t.ToString()).ToList();
                    headers = mpData[1].Select(h => h.ToString()).ToList();
                    var allData = mpData.Skip(2).Where(x => x[0] != null).ToList();

                    dataExpander = new DataExpander(types, headers);

                    // ModelEngine의 ModelPoints 설정
                    App.ModelEngine.SetModelPoints(allData, types, headers);

                    // 각 테이블 타입에 대한 탭 생성
                    await ClassifyAndCreateTabs();

                    // 첫 번째 테이블의 첫 번째 데이터로 SelectedPoint 설정
                    var firstTablePoints = App.ModelEngine.ModelPoints.FirstOrDefault();
                    if (firstTablePoints.Value != null && firstTablePoints.Value.Any())
                    {
                        var firstPoint = firstTablePoints.Value[0];
                        var firstExpandedData = dataExpander.ExpandData(firstPoint).FirstOrDefault();
                        if (firstExpandedData != null)
                        {
                            SelectedData = firstExpandedData;
                            UpdateSelectedDataDisplay();

                            // ModelEngine의 SetModelPoint 실행
                            App.ModelEngine.SelectedPoint = firstExpandedData;
                            App.ModelEngine.SetModelPoint();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("MP 데이터를 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void UpdateDataGrid(DataGrid dataGrid, List<string> headers, List<List<object>> data)
        {
            try
            {
                dataGrid.Columns.Clear();
                for (int i = 0; i < headers.Count; i++)
                {
                    var column = new DataGridTextColumn
                    {
                        Header = headers[i],
                        Binding = new Binding($"[{i}]"),
                    };

                    if (dataExpander.GetTypes().ElementAt(i) == typeof(DateTime))
                    {
                        column.Binding.StringFormat = "yyyy-MM-dd";
                    }

                    dataGrid.Columns.Add(column);
                }

                dataGrid.ItemsSource = data;
                AutoFitColumns(dataGrid);
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"데이터 그리드 업데이트 중 오류 발생: {ex.Message}");
            }
        }

        private async Task ClassifyAndCreateTabs()
        {
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                TableTabControl.Items.Clear();
                tableGrids.Clear();
                tableNameData.Clear();

                // 테이블 타입별로 데이터 그룹화 및 저장
                var groupedData = App.ModelEngine.ModelPoints
                    .GroupBy(entry => entry.Key.Split('|')[0])
                    .ToDictionary(
                        group => group.Key,
                        group => group.SelectMany(entry => entry.Value).ToList()
                    );

                // 그룹화된 데이터로 탭 생성 및 딕셔너리에 저장
                foreach (var group in groupedData)
                {
                    var tableName = group.Key;
                    var combinedData = group.Value;

                    // 딕셔너리에 데이터 저장
                    tableNameData[tableName] = combinedData;

                    // DataGrid 생성 및 설정
                    var dataGrid = CreateDataGrid();
                    tableGrids[tableName] = dataGrid;

                    var tabItem = new TabItem
                    {
                        Header = tableName,
                        Content = dataGrid
                    };
                    TableTabControl.Items.Add(tabItem);

                    UpdateDataGrid(dataGrid, headers, combinedData);
                }

                if (TableTabControl.Items.Count > 0)
                {
                    TableTabControl.SelectedIndex = 0;
                }
            });
        }

        private DataGrid CreateDataGrid()
        {
            var dataGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserSortColumns = false
            };
            dataGrid.MouseDoubleClick += MainDataGrid_MouseDoubleClick;
            return dataGrid;
        }

        private void AutoFitColumns(DataGrid dataGrid)
        {
            foreach (var column in dataGrid.Columns)
            {
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            dataGrid.UpdateLayout();
        }

        private void ShowErrorMessage(string message)
        {
            MessageBox.Show(message, "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void MainDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (sender is DataGrid dataGrid && dataGrid.SelectedItem is List<object> selectedRow)
                {
                    if (TableTabControl.SelectedItem is TabItem selectedTab)
                    {
                        var tableType = selectedTab.Header.ToString();
                        if (tableNameData.ContainsKey(tableType))
                        {
                            var expandedData = dataExpander.ExpandData(selectedRow).ToList();
                            ShowExpandedDataWindow(expandedData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"확장된 데이터 표시 중 오류 발생: {ex.Message}");
            }
        }

        private void ShowExpandedDataWindow(List<List<object>> expandedData)
        {
            try
            {
                if (expandedDataWindow != null)
                {
                    expandedDataWindow.Close();
                }

                expandedDataWindow = new Window
                {
                    Title = "확장된 데이터",
                    Width = 800,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };

                var dataGrid = new DataGrid
                {
                    AutoGenerateColumns = false,
                    IsReadOnly = true,
                };

                dataGrid.MouseDoubleClick += ExpandedDataGrid_MouseDoubleClick;

                UpdateDataGrid(dataGrid, headers, expandedData);

                expandedDataWindow.Content = dataGrid;
                expandedDataWindow.Show();
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"확장된 데이터 창 생성 중 오류 발생: {ex.Message}");
            }
        }

        private void ExpandedDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGrid dataGrid && dataGrid.SelectedItem is List<object> selectedRow)
            {
                UpdateSelectedData(selectedRow);
            }
        }

        private void UpdateSelectedData(List<object> newSelectedData)
        {
            SelectedData = newSelectedData;
            UpdateSelectedDataDisplay();

            // ModelEngine의 SetModelPoint 실행
            App.ModelEngine.SelectedPoint = newSelectedData;
            App.ModelEngine.SetModelPoint();

            MessageBox.Show("선택된 데이터가 업데이트되었습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);

            if (expandedDataWindow != null)
            {
                expandedDataWindow.Close();
                expandedDataWindow = null;
            }
        }

        private void UpdateSelectedDataDisplay()
        {
            if (SelectedData != null)
            {
                var displayText = string.Join(", ", SelectedData.Select((value, index) => $"{headers[index]}: {value}"));
                SelectedDataTextBlock.Text = displayText;
            }
            else
            {
                SelectedDataTextBlock.Text = "선택된 데이터가 없습니다.";
            }
        }

        private async void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isSearching) return;
            isSearching = true;

            await Task.Delay(300);

            string filterText = SearchTextBox.Text;
            ApplyFilter(filterText);

            isSearching = false;
        }

        private void ApplyFilter(string filterText)
        {
            if (TableTabControl.SelectedItem is not TabItem selectedTab) return;

            var tableType = selectedTab.Header.ToString();
            var dataGrid = tableGrids[tableType];

            // 저장된 데이터 사용
            if (!tableNameData.TryGetValue(tableType, out var currentTableData))
                return;

            if (string.IsNullOrWhiteSpace(filterText))
            {
                dataGrid.ItemsSource = currentTableData;
            }
            else
            {
                var filteredData = currentTableData.Where(row =>
                    row.Any(cell =>
                        cell != null &&
                        cell.ToString().IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
                    )
                ).ToList();
                dataGrid.ItemsSource = filteredData;
            }
        }
    }

}