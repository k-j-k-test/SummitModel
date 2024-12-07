using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ActuLiteModel;
using System.Windows.Media;
using System.Windows.Input;
using System.Linq;
using System.Threading.Tasks;

namespace ActuLight.Pages
{
    public partial class AssumptionPage : Page
    {
        private DataGrid detailGrid;
        private List<AssumptionDisplayData> originalAssumptionData;
        private List<ExpenseDisplayData> originalExpenseData;
        private bool isSearching = false;
        private RiskRateManagementWindow searchWindow;

        public AssumptionPage()
        {
            InitializeComponent();
        }

        private async void LoadData_Click(object sender, RoutedEventArgs e)
        {
            await LoadDataAsync();
        }

        public async Task LoadDataAsync()
        {
            await LoadAssumptionDataAsync();
            await LoadExpenseDataAsync();
        }

        public async Task LoadAssumptionDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage?.excelData != null && filePage.excelData.ContainsKey("assum"))
                {
                    var assumData = filePage.excelData["assum"];
                    var headers = assumData[0];
                    var data = assumData.Where(x => x[0] != null).Skip(1).ToList();
                    var assumList = ExcelImporter.ConvertToClassList<Input_assum>(data);
                    App.ModelEngine.SetAssumption(assumList);
                    originalAssumptionData = App.ModelEngine.Assumptions.Select(kvp => new AssumptionDisplayData
                    {
                        Key = kvp.Key,
                        Count = kvp.Value.Count
                    }).ToList();
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        AssumptionDataGrid.ItemsSource = originalAssumptionData;
                    });
                }
                else
                {
                    MessageBox.Show("Assumption 데이터를 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Assumption 데이터 로드 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        public async Task LoadExpenseDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage?.excelData != null && filePage.excelData.ContainsKey("exp"))
                {
                    var expData = filePage.excelData["exp"];
                    var headers = expData[0];
                    var data = expData.Where(x => x[0] != null).Skip(1).ToList();
                    var expList = ExcelImporter.ConvertToClassList<Input_exp>(data);
                    App.ModelEngine.SetExpense(expList);
                    originalExpenseData = App.ModelEngine.Expenses.Select(kvp => new ExpenseDisplayData
                    {
                        Key = kvp.Key,
                        ProductCode = kvp.Value.First().ProductCode,
                        RiderCode = kvp.Value.First().RiderCode,
                        Count = kvp.Value.Count
                    }).ToList();
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        ExpenseDataGrid.ItemsSource = originalExpenseData;
                    });
                }
                else
                {
                    MessageBox.Show("Expense 데이터를 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Expense 데이터 로드 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        public class AssumptionDisplayData
        {
            public string Key { get; set; }
            public int Count { get; set; }
        }

        public class ExpenseDisplayData
        {
            public string Key { get; set; }
            public string ProductCode { get; set; }
            public string RiderCode { get; set; }
            public int Count { get; set; }
        }

        private void AssumptionDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (AssumptionDataGrid.SelectedItem is AssumptionDisplayData selectedItem)
            {
                ShowDetailWindow(selectedItem.Key);
            }
        }

        private void ShowDetailWindow(string key)
        {
            if (App.ModelEngine.Assumptions.TryGetValue(key, out List<Input_assum> assumptions))
            {
                bool allZeros = assumptions.All(a => a.Rates?.All(r => r == 0) ?? true);
                if (allZeros)
                {
                    MessageBox.Show("표시할 데이터가 없습니다", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var detailWindow = new Window
                {
                    Title = "상세 가정 정보",
                    Width = 800,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };

                detailGrid = CreateDetailGrid(assumptions);
                detailWindow.Content = detailGrid;
                detailWindow.Show();
            }
        }

        private DataGrid CreateDetailGrid(List<Input_assum> assumptions)
        {
            var detailGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true
            };

            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Key1", Binding = new System.Windows.Data.Binding("Key1") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Key2", Binding = new System.Windows.Data.Binding("Key2") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Key3", Binding = new System.Windows.Data.Binding("Key3") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "조건", Binding = new System.Windows.Data.Binding("Condition") });

            int maxRatesLength = assumptions.Max(a => a.Rates?.Count ?? 0);

            for (int i = 0; i < maxRatesLength; i++)
            {
                int index = i;
                var rateColumn = new DataGridTextColumn
                {
                    Header = i.ToString(),
                    Binding = new System.Windows.Data.Binding($"Rates[{index}]")
                    {
                        StringFormat = "F6",
                        TargetNullValue = ""
                    }
                };
                detailGrid.Columns.Add(rateColumn);
            }

            detailGrid.ItemsSource = assumptions;
            return detailGrid;
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
                AssumptionDataGrid.ItemsSource = originalAssumptionData;
                ExpenseDataGrid.ItemsSource = originalExpenseData;
                return;
            }

            // Filter Assumption Data
            var filteredAssumptionData = originalAssumptionData?.Where(item =>
                item.Key.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                item.Count.ToString().IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
            ).ToList();
            AssumptionDataGrid.ItemsSource = filteredAssumptionData;

            // Filter Expense Data
            var filteredExpenseData = originalExpenseData?.Where(item =>
                item.Key.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                item.ProductCode?.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                item.RiderCode?.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                item.Count.ToString().IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
            ).ToList();
            ExpenseDataGrid.ItemsSource = filteredExpenseData;
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            if (searchWindow == null || !searchWindow.IsLoaded)
            {
                searchWindow = new RiskRateManagementWindow();
                searchWindow.Closed += (s, args) => searchWindow = null;
                searchWindow.Show();
            }
            else
            {
                searchWindow.Activate();
                if (searchWindow.WindowState == WindowState.Minimized)
                    searchWindow.WindowState = WindowState.Normal;
            }
        }

        private void ExpenseDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExpenseDataGrid.SelectedItem is ExpenseDisplayData selectedItem)
            {
                ShowExpenseDetailWindow(selectedItem.Key);
            }
        }

        private void ShowExpenseDetailWindow(string key)
        {
            if (App.ModelEngine.Expenses.TryGetValue(key, out List<Input_exp> expenses))
            {
                var detailWindow = new Window
                {
                    Title = "상세 비용 정보",
                    Width = 1000,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };

                var grid = CreateExpenseDetailGrid(expenses);
                detailWindow.Content = grid;
                detailWindow.Show();
            }
        }

        private DataGrid CreateExpenseDetailGrid(List<Input_exp> expenses)
        {
            var detailGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true
            };

            // Add columns for expense details
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "ProductCode", Binding = new System.Windows.Data.Binding("ProductCode") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "RiderCode", Binding = new System.Windows.Data.Binding("RiderCode") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Condition", Binding = new System.Windows.Data.Binding("Condition") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Alpha_P", Binding = new System.Windows.Data.Binding("Alpha_P") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Alpha_P2", Binding = new System.Windows.Data.Binding("Alpha_P2") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Alpha_S", Binding = new System.Windows.Data.Binding("Alpha_S") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Beta_P", Binding = new System.Windows.Data.Binding("Beta_P") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Beta_S", Binding = new System.Windows.Data.Binding("Beta_S") });
            detailGrid.Columns.Add(new DataGridTextColumn { Header = "Gamma", Binding = new System.Windows.Data.Binding("Gamma") });

            detailGrid.ItemsSource = expenses;
            return detailGrid;
        }
    }
}