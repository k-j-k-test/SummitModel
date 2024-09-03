﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ActuLiteModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.IO;
using Windows.Phone.Notification.Management;

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

            StartButton.IsEnabled = false;
            CancelButton.IsEnabled = false;
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

                string excelName = Path.GetFileNameWithoutExtension(FilePage.SelectedFilePath);
                _outputFolderPath = Path.Combine(Path.GetDirectoryName(FilePage.SelectedFilePath), @$"Data_{excelName}\Outputs");
                if (!Directory.Exists(_outputFolderPath))
                {
                    Directory.CreateDirectory(_outputFolderPath);
                }

                _modelWriter = new ModelWriter(App.ModelEngine, new DataExpander(App.ModelEngine.ModelPointInfo.Types, App.ModelEngine.ModelPointInfo.Headers));
                _modelWriter.LoadTableData(filePage.excelData["out"]);
                PopulateTabControl();

                StartButton.IsEnabled = true;
                CancelButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 로드 중 오류가 발생했습니다: \n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    IsReadOnly = true,
                    CanUserSortColumns = false
                };

                dataGrid.Columns.Add(new DataGridTextColumn { Header = "ColumnName", Binding = new System.Windows.Data.Binding("ColumnName") });
                dataGrid.Columns.Add(new DataGridTextColumn { Header = "Value", Binding = new System.Windows.Data.Binding("Value") });
                dataGrid.Columns.Add(new DataGridTextColumn { Header = "RangeStart", Binding = new System.Windows.Data.Binding("RangeStart") });
                dataGrid.Columns.Add(new DataGridTextColumn { Header = "RangeEnd", Binding = new System.Windows.Data.Binding("RangeEnd") });
                dataGrid.Columns.Add(new DataGridTextColumn { Header = "Format", Binding = new System.Windows.Data.Binding("Format") });

                var items = new List<OutTableColumnInfo>();

                foreach (var columnEntry in tableEntry.Value)
                {
                    var columnInfo = new OutTableColumnInfo
                    {
                        ColumnName = columnEntry.Key,
                        Value = columnEntry.Value.Expression.Text,
                        RangeStart = columnEntry.Value.StartExpression?.Text ?? "",
                        RangeEnd = columnEntry.Value.EndExpression?.Text ?? "",
                        Format = columnEntry.Value.Format,
                    };
                    items.Add(columnInfo);
                }
                dataGrid.ItemsSource = items;

                tabItem.Content = dataGrid;
                OutputTabControl.Items.Add(tabItem);
            }

            OutputTabControl.SelectedIndex = 0;
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _modelWriter.IsCanceled = false;
                _isCancelled = false;

                _startTime = DateTime.Now;
                _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(20) };
                _timer.Tick += Timer_Tick;
                
                _timer.Start();

                EnableButtons(false);

                ProgressRichTextBox.AppendText($"■ 시작 {DateTime.Now} " + "\r\n");
                await Task.Run(() => WriteResults());
                ProgressRichTextBox.AppendText($"□ 종료 {DateTime.Now}, 걸린 시간: {(DateTime.Now - _startTime).ToString(@"hh\:mm\:ss")} " + "\r\n" + "\r\n");

                EnableButtons(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류가 발생했습니다. 출력 데이터가 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            _isCancelled = true;
            _modelWriter.IsCanceled = true;
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
            });
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan elapsed = DateTime.Now - _startTime;
            Dispatcher.Invoke(() =>
            {
                TimeLabel.Text = $"경과 시간: {elapsed:hh\\:mm\\:ss\\.ff}";
                ProgressLabel.Text = _modelWriter.StatusMessage;

                if(_modelWriter.StatusQueue.Count > 0 )
                {
                    ProgressRichTextBox.AppendText("- " + _modelWriter.StatusQueue.Dequeue() + "\r\n");
                }
            });
        }

        private void UpdateProgressRichTextBox()
        {
            ProgressLabel.Text = _modelWriter.StatusMessage;
            while (_modelWriter.StatusQueue.Count > 0)
            {
                ProgressRichTextBox.AppendText("- " + _modelWriter.StatusQueue.Dequeue() + "\r\n");
            }
            ProgressRichTextBox.ScrollToEnd();
        }

        private void EnableButtons(bool enabled)
        {
            var mainWindow = (MainWindow)Application.Current.MainWindow;
            NavigationBar navigationBar = mainWindow.GlobalNavBarControl.Content as NavigationBar;
            navigationBar.SetButtonsEnabled(enabled);

            // OutputPage의 버튼들 활성화
            StartButton.IsEnabled = enabled;

            //자동저장기능 중지
            FilePage.IsAutoSync = enabled;       
        }
    }

    public class OutTableColumnInfo
    {
        public string ColumnName { get; set; }
        public string Value { get; set; }
        public string RangeStart { get; set; }
        public string RangeEnd { get; set; }
        public string Format { get; set; }
    }
}