using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ActuLiteModel;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ICSharpCode.AvalonEdit;

namespace ActuLight.Pages
{
    public partial class SpreadSheetPage : Page
    {
        public SyntaxHighlighter SyntaxHighlighter;

        public Dictionary<string, Model> Models = App.ModelEngine.Models;
        public Dictionary<string, string> Scripts { get; set; } = new Dictionary<string, string>();

        private string selectedModel;
        private string selectedCell;

        private MatchCollection cellMatches;
        private MatchCollection invokeMatches;

        public SpreadSheetPage()
        {
            InitializeComponent();
            SyntaxHighlighter = new SyntaxHighlighter(ScriptEditor);
            ScriptEditor.TextArea.TextView.LineTransformers.Add(SyntaxHighlighter);
            UpdateSyntaxHighlighter();

            // ScriptEditor의 배경색을 진한 회색으로 설정 (Dark Theme)
            ScriptEditor.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));

            // 텍스트 색상을 밝은 회색으로 설정 (기본 텍스트 색상)
            ScriptEditor.Foreground = new SolidColorBrush(Color.FromRgb(220, 220, 220));

            // 선택 영역의 배경색 설정
            ScriptEditor.TextArea.SelectionBrush = new SolidColorBrush(Color.FromRgb(60, 60, 60));
            ScriptEditor.TextArea.SelectionForeground = new SolidColorBrush(Color.FromRgb(255, 255, 255));
        }

        private async void LoadData_Click(object sender, RoutedEventArgs e) => await LoadDataAsync();

        private async Task LoadDataAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            try
            {
                var filePage = ((MainWindow)Application.Current.MainWindow).pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage?.excelData?.ContainsKey("cell") == true)
                {
                    var cells = ExcelHelper.ConvertToClassList<Input_cell>(filePage.excelData["cell"].Data);

                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        App.ModelEngine.SetModel(cells);

                        Scripts.Clear();
                        foreach (var modelPair in Models)
                        {
                            Scripts[modelPair.Key] = string.Join(Environment.NewLine, modelPair.Value.CompiledCells.Values.Select(cell =>
                                $"{(string.IsNullOrWhiteSpace(cell.Description) ? "" : $"//{cell.Description}{Environment.NewLine}")}{cell.Name} -- {cell.Formula}{Environment.NewLine}"));
                        }

                        ModelsList.ItemsSource = Models.Keys.ToList();
                        UpdateSyntaxHighlighter();
                    });
                    MessageBox.Show("데이터를 성공적으로 불러왔습니다.", "성공", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Cell 데이터를 찾을 수 없습니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void ModelsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ModelsList.SelectedItem is string selectedItem)
            {
                if (selectedModel != null)
                {
                    Scripts[selectedModel] = ScriptEditor.Text;
                }
                UpdateCellList(selectedItem);
                ScriptEditor.Text = Scripts.TryGetValue(selectedItem, out string script) ? script : "";
                selectedModel = selectedItem;

                UpdateSyntaxHighlighter();
            }
        }

        private void CellsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CellsList.SelectedItem is TextBlock selectedTextBlock)
            {
                selectedCell = selectedTextBlock.Text;
                UpdateCellStatus();
            }
            else
            {
                selectedCell = null;
                CellStatusTextBlock.Text = "No cell selected";
            }
        }

        private async void ScriptEditor_TextChanged(object sender, EventArgs e)
        {
            await Task.Delay(250); // 250ms delay to prevent frequent updates

            if (selectedModel == null || !Models.ContainsKey(selectedModel))
            {
                return;
            }

            Model model = Models[selectedModel];

            // Update cellMatches
            var cellPattern = new Regex(@"(?://(?<description>.*)\r?\n)?(?<cellName>\w+)\s*--\s*(?<formula>.+)(\r?\n|$)");
            var newCellMatches = cellPattern.Matches(ScriptEditor.Text);

            // Check for cell changes
            bool hasCellChanges = cellMatches == null ||
                                  cellMatches.Count != newCellMatches.Count ||
                                  !Enumerable.Range(0, cellMatches.Count)
                                             .All(i => cellMatches[i].Groups["cellName"].Value == newCellMatches[i].Groups["cellName"].Value &&
                                                       cellMatches[i].Groups["formula"].Value == newCellMatches[i].Groups["formula"].Value);

            if (hasCellChanges)
            {
                cellMatches = newCellMatches;
                UpdateModelCells(model, cellMatches);
                UpdateCellList(selectedModel);
            }

            // Update invokeMatches
            var invokePattern = new Regex(@"Invoke\((\w+),\s*(\d+)\)");
            var newInvokeMatches = invokePattern.Matches(ScriptEditor.Text);

            // Check for Invoke changes
            bool hasInvokeChanges = invokeMatches == null ||
                                    invokeMatches.Count != newInvokeMatches.Count ||
                                    !Enumerable.Range(0, invokeMatches.Count)
                                               .All(i => invokeMatches[i].Groups[1].Value == newInvokeMatches[i].Groups[1].Value &&
                                                         invokeMatches[i].Groups[2].Value == newInvokeMatches[i].Groups[2].Value);

            if (hasInvokeChanges || hasCellChanges)
            {
                invokeMatches = newInvokeMatches;
                UpdateInvokes(model);
            }
        }

        private void AddSheetTab(string sheetName, Sheet sheet)
        {
            var tabItem = new TabItem
            {
                Header = sheetName
            };

            var dataGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true
            };

            var sheetData = sheet.GetAllData();
            var rows = sheetData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            if (rows.Length > 0)
            {
                var headers = rows[0].Split('\t');

                foreach (var header in headers)
                {
                    dataGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = header,
                        Binding = new System.Windows.Data.Binding($"[{header}]")
                    });
                }

                var data = new List<Dictionary<string, string>>();

                for (int i = 1; i < rows.Length; i++)
                {
                    var values = rows[i].Split('\t');
                    var rowData = new Dictionary<string, string>();
                    bool hasNonEmptyValue = false;

                    for (int j = 0; j < headers.Length && j < values.Length; j++)
                    {
                        if (!string.IsNullOrWhiteSpace(values[j]))
                        {
                            rowData[headers[j]] = values[j];
                            hasNonEmptyValue = true;
                        }
                    }

                    if (hasNonEmptyValue)
                    {
                        data.Add(rowData);
                    }
                }

                dataGrid.ItemsSource = data;
            }

            tabItem.Content = dataGrid;
            SheetTabControl.Items.Add(tabItem);
        }

        private void AddModel_Click(object sender, RoutedEventArgs e)
        {
            string newModelName = NewModelTextBox.Text.Trim();

            if (!string.IsNullOrWhiteSpace(newModelName) && !Models.ContainsKey(newModelName))
            {
                Models[newModelName] = new Model(newModelName, App.ModelEngine);
                Scripts[newModelName] = "";

                var modelsList = (List<string>)ModelsList.ItemsSource ?? new List<string>();
                modelsList.Add(newModelName);
                ModelsList.ItemsSource = null;
                ModelsList.ItemsSource = modelsList;

                NewModelTextBox.Text = "";
                ModelsList.SelectedItem = newModelName;
                UpdateCellList(newModelName);
            }
        }

        private void DeleteModel_Click(object sender, RoutedEventArgs e)
        {
            if (ModelsList.SelectedItem is string selectedModel)
            {
                Models.Remove(selectedModel);
                Scripts.Remove(selectedModel);

                var modelsList = (List<string>)ModelsList.ItemsSource;
                modelsList.Remove(selectedModel);
                ModelsList.ItemsSource = null;
                ModelsList.ItemsSource = modelsList;

                ScriptEditor.Text = "";
            }
        }

        private void UpdateCellList(string selectedModel)
        {
            if (Models.TryGetValue(selectedModel, out Model model))
            {
                var cellItems = new List<UIElement>();
                foreach (var cell in model.CompiledCells)
                {
                    var textBlock = new TextBlock { Text = cell.Key };
                    if (!cell.Value.IsCompiled)
                    {
                        textBlock.Foreground = Brushes.Red;
                        textBlock.FontWeight = FontWeights.Bold;
                    }
                    cellItems.Add(textBlock);
                }
                CellsList.ItemsSource = cellItems;
                SyntaxHighlighter.UpdateModelCells(selectedModel, model.CompiledCells.Keys);
            }
            else
            {
                CellsList.ItemsSource = null;
            }
        }

        private void UpdateCellStatus()
        {
            if (Models.TryGetValue(selectedModel, out Model model) &&
                model.CompiledCells.TryGetValue(selectedCell, out CompiledCell cell))
            {
                CellStatusTextBlock.Text = $"Formula: {cell.Formula}\n" +
                                           $"Status: {cell.CompileStatusMessage}";
            }
            else
            {
                CellStatusTextBlock.Text = "No cell selected";
            }
        }

        private void UpdateModelCells(Model model, MatchCollection cellMatches)
        {
            var newCells = new HashSet<string>();

            foreach (Match match in cellMatches)
            {
                string cellName = match.Groups["cellName"].Value;
                string description = match.Groups["description"].Value.Trim();
                string formula = match.Groups["formula"].Value.Trim();

                newCells.Add(cellName);

                if (!model.CompiledCells.TryGetValue(cellName, out CompiledCell existingCell) || existingCell.Formula != formula)
                {
                    model.ResisterCell(cellName, formula, description);
                }
                else
                {
                    existingCell.Description = description;
                }
            }

            // Remove cells that are no longer in the script
            var cellsToRemove = model.CompiledCells.Keys.Except(newCells).ToList();
            foreach (var cellToRemove in cellsToRemove)
            {
                model.CompiledCells.Remove(cellToRemove);
            }
        }

        private void UpdateInvokes(Model model)
        {
            // Clear all model sheets
            foreach (var m in Models.Values)
            {
                //Models = new Dictionary<string, Model>();
                m.Sheets.Clear();
                m.Parameter = new Parameter();
            }

            // Process each Invoke call
            foreach (Match match in invokeMatches)
            {
                string cellName = match.Groups[1].Value;
                int t = int.Parse(match.Groups[2].Value);

                try
                {
                    model.Invoke(cellName, t);
                }
                catch (Exception ex)
                {
                    model.Sheets["ErrorSheet"] = new Sheet();
                    model.Sheets["ErrorSheet"].RegisterMethod("ErrorMessage", _ => 0);
                    model.Sheets["ErrorSheet"]["ErrorMessage", 0] = 0;
                    CellStatusTextBlock.Text = $"Error during Invoke: {ex.Message}";
                }
            }

            UpdateSheets();
        }

        private void UpdateSheets()
        {
            // 현재 선택된 탭의 이름 저장
            string selectedTabName = (SheetTabControl.SelectedItem as TabItem)?.Header?.ToString();

            // Clear existing tabs
            SheetTabControl.Items.Clear();

            // First, add tabs for the current model's sheets
            if (Models.TryGetValue(selectedModel, out Model currentModel))
            {
                foreach (var sheetPair in currentModel.Sheets)
                {
                    AddSheetTab(sheetPair.Key, sheetPair.Value);
                }
            }

            // Then, add tabs for other models' sheets
            foreach (var modelPair in Models.Where(m => m.Key != selectedModel))
            {
                foreach (var sheetPair in modelPair.Value.Sheets)
                {
                    AddSheetTab(sheetPair.Key, sheetPair.Value);
                }
            }

            // 이전에 선택된 탭 또는 첫 번째 탭 선택
            if (!string.IsNullOrEmpty(selectedTabName))
            {
                var tabToSelect = SheetTabControl.Items.Cast<TabItem>().FirstOrDefault(t => t.Header.ToString() == selectedTabName);
                if (tabToSelect != null)
                {
                    tabToSelect.IsSelected = true;
                }
                else if (SheetTabControl.Items.Count > 0)
                {
                    (SheetTabControl.Items[0] as TabItem).IsSelected = true;
                }
            }
            else if (SheetTabControl.Items.Count > 0)
            {
                (SheetTabControl.Items[0] as TabItem).IsSelected = true;
            }
        }

        private void UpdateSyntaxHighlighter()
        {
            foreach (var model in Models.Values)
            {
                SyntaxHighlighter.UpdateModelCells(model.Name, model.CompiledCells.Keys);
            }

            SyntaxHighlighter.UpdateModels(Models.Keys);
            SyntaxHighlighter.UpdateContextVariables(App.ModelEngine.Context.Variables.Keys);
            SyntaxHighlighter.UpdateAssumptions(App.ModelEngine.Assumptions.Keys);
        }

    }
}