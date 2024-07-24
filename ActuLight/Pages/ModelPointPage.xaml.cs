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

namespace ActuLight.Pages
{
    public partial class ModelPointPage : Page
    {
        private DataExpander dataExpander;
        private List<List<object>> originalData;
        private List<string> headers;
        public List<object> SelectedData { get; private set; }
        private Window expandedDataWindow;
        private bool isSearching = false;

        public ModelPointPage()
        {
            InitializeComponent();
        }

        private async void LoadData_Click(object sender, RoutedEventArgs e)
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;

            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;

                if (filePage != null && filePage.excelData != null && filePage.excelData.ContainsKey("mp"))
                {
                    var mpData = filePage.excelData["mp"];
                    var typeNames = mpData.Headers;
                    var keys = mpData.Data[0];

                    dataExpander = new DataExpander(typeNames, keys);
                    originalData = mpData.Data.Skip(1).ToList();
                    headers = dataExpander.GetKeys().ToList();

                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        UpdateDataGrid(MainDataGrid, headers, originalData);
                        AutoFitColumns();
                    });
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
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"데이터 그리드 업데이트 중 오류 발생: {ex.Message}");
            }
        }

        private void AutoFitColumns()
        {
            foreach (var column in MainDataGrid.Columns)
            {
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
            }
            MainDataGrid.UpdateLayout();
        }

        private void ShowErrorMessage(string message)
        {
            MessageBox.Show(message, "오류", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void MainDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (MainDataGrid.SelectedItem is List<object> selectedRow)
                {
                    var expandedData = dataExpander.ExpandData(selectedRow).ToList();
                    ShowExpandedDataWindow(expandedData);
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
            App.ModelEngine.SetModelPoint(headers, SelectedData);

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
            if (string.IsNullOrWhiteSpace(filterText))
            {
                MainDataGrid.ItemsSource = originalData;
            }
            else
            {
                var filteredData = originalData.Where(row =>
                    row.Any(cell => cell?.ToString().Contains(filterText, StringComparison.OrdinalIgnoreCase) == true)
                ).ToList();

                MainDataGrid.ItemsSource = filteredData;
            }
        }
    }
}