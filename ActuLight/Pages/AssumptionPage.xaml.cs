using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ActuLiteModel;
using LiveCharts;
using LiveCharts.Wpf;
using System.Windows.Media;
using System.Windows.Input;
using System.Linq;
using System.Threading.Tasks;

namespace ActuLight.Pages
{
    public partial class AssumptionPage : Page
    {
        private DataGrid detailGrid;
        private CartesianChart chart;
        private List<AssumptionDisplayData> originalData;
        private bool isSearching = false;

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
            LoadingOverlay.Visibility = Visibility.Visible;

            try
            {
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                var filePage = mainWindow.pageCache["Pages/FilePage.xaml"] as FilePage;

                if (filePage != null && filePage.excelData != null && filePage.excelData.ContainsKey("assum"))
                {
                    var assumData = filePage.excelData["assum"];

                    var headers = assumData[0];
                    var data = assumData.Skip(1).ToList();

                    var assumList = ExcelImporter.ConvertToClassList<Input_assum>(data);

                    App.ModelEngine.SetAssumption(assumList);

                    originalData = App.ModelEngine.Assumptions.Select(kvp => new AssumptionDisplayData
                    {
                        Key = kvp.Key,
                        Count = kvp.Value.Count
                    }).ToList();

                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        AssumptionDataGrid.ItemsSource = originalData;
                    });
                }
                else
                {
                    MessageBox.Show("Assumption 데이터를 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
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

        public class AssumptionDisplayData
        {
            public string Key { get; set; }
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
                var detailWindow = new Window
                {
                    Title = "상세 가정 정보",
                    Width = 1200,
                    Height = 800,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };

                var grid = new Grid();
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(5) });
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

                detailGrid = CreateDetailGrid(assumptions);
                Grid.SetRow(detailGrid, 0);
                grid.Children.Add(detailGrid);

                var splitter = new GridSplitter
                {
                    Height = 5,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Center,
                    Background = Brushes.Gray
                };
                Grid.SetRow(splitter, 1);
                grid.Children.Add(splitter);

                chart = CreateChart(assumptions);
                Grid.SetRow(chart, 2);
                grid.Children.Add(chart);

                detailWindow.Content = grid;

                detailGrid.SelectionChanged += DetailGrid_SelectionChanged;

                detailWindow.Show();
            }
        }

        private void DetailGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = detailGrid.SelectedItem as Input_assum;
            if (selectedItem != null)
            {
                HighlightSeries(selectedItem);
            }
        }

        private void HighlightSeries(Input_assum selectedAssumption)
        {
            foreach (var series in chart.Series)
            {
                if (series is LineSeries lineSeries)
                {
                    if (lineSeries.Title == (selectedAssumption.Condition ?? $"Assumption {chart.Series.IndexOf(series) + 1}"))
                    {
                        lineSeries.StrokeThickness = 3;
                        lineSeries.Opacity = 1;
                        lineSeries.Fill = new SolidColorBrush(Color.FromArgb(10, ((SolidColorBrush)lineSeries.Stroke).Color.R,
                                                                                 ((SolidColorBrush)lineSeries.Stroke).Color.G,
                                                                                 ((SolidColorBrush)lineSeries.Stroke).Color.B));
                        Panel.SetZIndex(lineSeries, 1);
                    }
                    else
                    {
                        lineSeries.StrokeThickness = 1.5;
                        lineSeries.Opacity = 0.5;
                        lineSeries.Fill = Brushes.Transparent;
                        Panel.SetZIndex(lineSeries, 0);
                    }
                }
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
                        StringFormat = "F4",
                        TargetNullValue = ""
                    }
                };
                detailGrid.Columns.Add(rateColumn);
            }

            detailGrid.ItemsSource = assumptions;
            return detailGrid;
        }

        private CartesianChart CreateChart(List<Input_assum> assumptions)
        {
            var validRanges = assumptions.Select(a => GetValidDataRange(a.Rates)).ToList();
            int maxValidIndex = validRanges.Max(r => r) + 5;

            var chart = new CartesianChart();

            chart.Series = new SeriesCollection();
            chart.AxisX = new AxesCollection
            {
                new Axis
                {
                    Title = "Index",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Colors.LightGray),
                    Separator = new LiveCharts.Wpf.Separator { Stroke = new SolidColorBrush(Colors.DimGray) },
                    FontWeight = FontWeights.Bold,
                    MinValue = 0,
                    MaxValue = maxValidIndex
                }
            };
            chart.AxisY = new AxesCollection
            {
                new Axis
                {
                    Title = "Rate",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Colors.LightGray),
                    Separator = new LiveCharts.Wpf.Separator { Stroke = new SolidColorBrush(Colors.DimGray) },
                    FontWeight = FontWeights.Bold,
                    LabelFormatter = value => Math.Round(value, 4).ToString("F4")
                }
            };
            chart.LegendLocation = LegendLocation.Right;
            chart.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            chart.Foreground = new SolidColorBrush(Colors.LightGray);
            chart.DataTooltip = new DefaultTooltip { Background = new SolidColorBrush(Color.FromRgb(60, 60, 60)) };
            chart.Zoom = ZoomingOptions.X;
            chart.Pan = PanningOptions.X;
            chart.DisableAnimations = true;
            chart.AnimationsSpeed = TimeSpan.FromMilliseconds(100);

            double minY = assumptions.SelectMany(a => a.Rates).Min();
            double maxY = assumptions.SelectMany(a => a.Rates).Max();
            chart.AxisY[0].MinValue = minY;
            chart.AxisY[0].MaxValue = maxY;

            var colorList = new List<Color>
            {
                Colors.Cyan, Colors.Magenta, Colors.LimeGreen, Colors.Yellow, Colors.Orange,
                Colors.HotPink, Colors.Aqua, Colors.GreenYellow, Colors.Gold, Colors.Orchid
            };

            for (int i = 0; i < assumptions.Count; i++)
            {
                var assumption = assumptions[i];
                var validRange = validRanges[i];
                var series = new LineSeries
                {
                    Title = assumption.Condition ?? $"Assumption {i + 1}",
                    Values = new ChartValues<double>(assumption.Rates?.Take(validRange + 5) ?? new List<double>()),
                    Stroke = new SolidColorBrush(colorList[i % colorList.Count]),
                    Fill = Brushes.Transparent,
                    PointGeometry = null,
                    LineSmoothness = 0,
                    StrokeThickness = 1.5
                };
                chart.Series.Add(series);
            }

            return chart;
        }

        private int GetValidDataRange(List<double> rates)
        {
            if (rates == null || rates.Count == 0)
                return 0;

            for (int i = rates.Count - 1; i >= 0; i--)
            {
                if (rates[i] != 0)
                    return i;
            }

            return 0;
        }

        private void ShowErrorMessage(string message)
        {
            MessageBox.Show(message, "오류", MessageBoxButton.OK, MessageBoxImage.Error);
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
                AssumptionDataGrid.ItemsSource = originalData;
            }
            else
            {
                var filteredData = originalData.Where(item =>
                    item.Key.Contains(filterText, StringComparison.OrdinalIgnoreCase) ||
                    item.Count.ToString().Contains(filterText)
                ).ToList();

                AssumptionDataGrid.ItemsSource = filteredData;
            }
        }
    }
}