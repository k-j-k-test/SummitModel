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
        private List<ScriptRuleDisplayData> originalScriptRuleData;
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
            await LoadScriptRuleDataAsync();
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
                        //ProductCode = kvp.Value.First().ProductCode,
                        //RiderCode = kvp.Value.First().RiderCode,
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

        public async Task LoadScriptRuleDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage?.excelData != null && filePage.excelData.ContainsKey("rule"))
                {
                    // Load script rules into ModelEngine
                    App.ModelEngine.SetScriptRule(filePage.excelData["rule"]);

                    // Prepare display data
                    var displayData = App.ModelEngine.ScriptRules
                        .GroupBy(kvp => kvp.Key)
                        .Select(g => new ScriptRuleDisplayData
                        {
                            Key = g.Key,
                            Count = 1  // 중복된 키는 덮어쓰기 때문에 항상 1
                        })
                        .ToList();

                    originalScriptRuleData = displayData;

                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        ScriptRulesDataGrid.ItemsSource = originalScriptRuleData;
                    });
                }
                else
                {
                    MessageBox.Show("Script Rule 데이터를 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Script Rule 데이터 로드 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
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
            public int Count { get; set; }
        }

        public class ScriptRuleDisplayData
        {
            public string Key { get; set; }
            public int Count { get; set; }
        }

        private void AssumptionDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (AssumptionDataGrid.SelectedItem is AssumptionDisplayData selectedItem)
            {
                ShowAssumptionDetailWindow(selectedItem.Key);
            }
        }

        private void ExpenseDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExpenseDataGrid.SelectedItem is ExpenseDisplayData selectedItem)
            {
                ShowExpenseDetailWindow(selectedItem.Key);
            }
        }

        private void ScriptRulesDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ScriptRulesDataGrid.SelectedItem is ScriptRuleDisplayData selectedItem)
            {
                ShowScriptRuleDetailWindow(selectedItem.Key);
            }
        }

        private void ShowAssumptionDetailWindow(string key)
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

                detailGrid = CreateAssumptionDetailGrid(assumptions);
                detailWindow.Content = detailGrid;
                detailWindow.Show();
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

        private void ShowScriptRuleDetailWindow(string key)
        {
            if (App.ModelEngine.ScriptRules.TryGetValue(key, out List<string> rules))
            {
                var detailWindow = new Window
                {
                    Title = $"Script Rule Details - {key}",
                    Width = 800,
                    Height = 400,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };

                var grid = new DataGrid
                {
                    AutoGenerateColumns = false,
                    IsReadOnly = true,
                    Margin = new Thickness(10)
                };

                // Get headers from the original data
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;
                var headers = filePage.excelData["rule"][0].Select(h => h?.ToString() ?? "").ToList();

                // Create columns based on headers
                for (int i = 0; i < headers.Count; i++)
                {
                    var index = i;
                    grid.Columns.Add(new DataGridTextColumn
                    {
                        Header = headers[index],
                        Binding = new System.Windows.Data.Binding($"[{index}]") { Mode = System.Windows.Data.BindingMode.OneWay }
                    });
                }

                // Set single row data
                grid.ItemsSource = new List<List<string>> { rules };

                detailWindow.Content = grid;
                detailWindow.Show();
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
                AssumptionDataGrid.ItemsSource = originalAssumptionData;
                ExpenseDataGrid.ItemsSource = originalExpenseData;
                ScriptRulesDataGrid.ItemsSource = originalScriptRuleData;
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
                item.Count.ToString().IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
            ).ToList();
            ExpenseDataGrid.ItemsSource = filteredExpenseData;

            // Filter Script Rules Data
            var filteredScriptRuleData = originalScriptRuleData?.Where(item =>
                item.Key.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                item.Count.ToString().IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0
            ).ToList();
            ScriptRulesDataGrid.ItemsSource = filteredScriptRuleData;
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

        private DataGrid CreateAssumptionDetailGrid(List<Input_assum> assumptions)
        {
            var detailGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true
            };

            var properties = typeof(Input_assum).GetProperties();

            foreach (var prop in properties)
            {
                if (prop.Name == "Rates")
                {
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
                }
                else
                {
                    detailGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = prop.Name,
                        Binding = new System.Windows.Data.Binding(prop.Name)
                    });
                }
            }

            detailGrid.ItemsSource = assumptions;
            return detailGrid;
        }

        private DataGrid CreateExpenseDetailGrid(List<Input_exp> expenses)
        {
            var detailGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true
            };

            // Input_exp의 모든 public 속성 가져오기
            var properties = typeof(Input_exp).GetProperties();

            // 각 속성에 대해 열 추가
            foreach (var prop in properties)
            {
                detailGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = prop.Name,
                    Binding = new System.Windows.Data.Binding(prop.Name)
                });
            }

            detailGrid.ItemsSource = expenses;
            return detailGrid;
        }
    }
}