using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ActuLiteModel;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Data;
using System.IO;
using Newtonsoft.Json;
using System.Text;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit;
using Microsoft.Win32;
using System.Windows.Threading;
using System.Windows.Controls.Primitives;
using Windows.Graphics.Printing.Workflow;
using System.Collections.ObjectModel;
using System.Diagnostics;
using ICSharpCode.AvalonEdit.Editing;

namespace ActuLight.Pages
{
    public partial class SpreadSheetPage : Page
    {
        public SyntaxHighlighter SyntaxHighlighter;
        public RegionFoldingStrategy FoldingStrategy;
        public FoldingManager FoldingManager;

        public Dictionary<string, Model> Models = App.ModelEngine.Models;
        public Dictionary<string, string> Scripts { get; set; } = new Dictionary<string, string>();
        public int SignificantDigits;

        private string selectedModel;
        private string selectedCell;

        private Dictionary<string, (Sheet Sheet, bool IsLoaded)> _tabData = new Dictionary<string, (Sheet, bool)>();
        private bool _areFoldingsCollapsed = false;

        private MatchCollection cellMatches;
        private List<(string CellName, int T)> invokeList;

        private bool scriptChanged = false;
        private Dictionary<string, MatchCollection> cellMatchesDict = new Dictionary<string, MatchCollection>();

        private DispatcherTimer autoSaveTimer;
        private int autoSaveTickCounter;

        public SpreadSheetPage()
        {
            InitializeComponent();

            // 1. 구문 강조 설정
            SyntaxHighlighter = new SyntaxHighlighter(ScriptEditor);
            ScriptEditor.TextArea.TextView.LineTransformers.Add(SyntaxHighlighter);
            UpdateSyntaxHighlighter();

            // 2. 기본 설정들
            LoadSignificantDigitsSetting();
            SetupAutoSaveTimer();

            // 3. 폴딩 관련 설정
            var foldingMargin = new FoldingMargin();
            ScriptEditor.TextArea.LeftMargins.Add(foldingMargin);
            FoldingManager = FoldingManager.Install(ScriptEditor.TextArea);
            FoldingStrategy = new RegionFoldingStrategy();
            FoldingStrategy.UpdateFoldings(FoldingManager, ScriptEditor.Document);

            // 4. 에디터 옵션 설정
            ScriptEditor.Options.EnableHyperlinks = true;
            ScriptEditor.Options.EnableEmailHyperlinks = true;
            ScriptEditor.Options.ConvertTabsToSpaces = true;
            ScriptEditor.Options.ShowSpaces = false;
            ScriptEditor.Options.HighlightCurrentLine = false;
            ScriptEditor.Options.AllowScrollBelowDocument = true;
            ScriptEditor.ShowLineNumbers = true;

            // 5. 이벤트 핸들러 설정
            ScriptEditor.TextChanged += ScriptEditor_TextChanged;
            ModelsList.MouseRightButtonUp += ModelsList_MouseRightButtonUp;
            ScriptEditor.PreviewKeyDown += ScriptEditor_PreviewKeyDown_Save;
            ScriptEditor.PreviewKeyDown += ScriptEditor_PreviewKeyDown_Folding;
            ScriptEditor.PreviewKeyDown += ScriptEditor_PreviewKeyDown_Region;
            ScriptEditor.PreviewKeyDown += ScriptEditor_PreviewKeyDown_Zoom;
            ScriptEditor.PreviewKeyDown += ScriptEditor_PreviewKeyDown_Renew;
        }

        public async void LoadData_Click(object sender, RoutedEventArgs e) => await LoadDataAsync();

        public async Task LoadDataAsync()
        {
            try
            {
                LoadingOverlay.Visibility = Visibility.Visible;
                string modelsFolder = Path.Combine(FilePage.SelectedFolderPath, "Scripts");

                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = modelsFolder,
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    await LoadDataAsync(openFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 로드 중 오류 발생: 먼저 파일열기를 진행해야 합니다.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
        }

        public async Task LoadDataAsync(string filePath)
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            try
            {
                string json = await Task.Run(() => File.ReadAllText(filePath));

                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    Scripts = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                    Models.Clear();
                    cellMatches = null;

                    foreach (var kvp in Scripts)
                    {
                        Models[kvp.Key] = new Model(kvp.Key, App.ModelEngine);
                    }

                    var modelKeys = Models.Keys.ToList();
                    ModelsList.ItemsSource = modelKeys;

                    ScriptEditor.Text = string.Empty;
                    ModelsList.SelectedIndex = -1;
                    selectedModel = null;

                    for (int i = Models.Count - 1; i >= 0; i--)
                    {
                        ModelsList.SelectedIndex = i;
                        ScriptEditor.Text = Scripts.Values.ToList()[i];
                        await UpdateModelAndUIFromScript(Scripts[ModelsList.SelectedItem.ToString()]);
                    }
                });
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

        private void LoadSignificantDigitsSetting()
        {
            AppSettings settings = null;

            if (File.Exists("settings.json"))
            {
                string json = File.ReadAllText("settings.json");
                settings = JsonConvert.DeserializeObject<AppSettings>(json);
            }
            
            SignificantDigits = settings?.SignificantDigits ?? 8;
        }

        private void ModelsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ModelsList.SelectedItem is string selectedItem)
            {
                selectedModel = selectedItem;
                ScriptEditor.Text = Scripts.TryGetValue(selectedModel, out string script) ? script : "";
                UpdateCellList(selectedModel);
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
            if (ThrottlerAsync.ShouldExecute("ScriptEditor_TextChanged", 300))
            {
                await DebouncerAsync.Debounce("ScriptEditor_TextChanged", 300, async () =>
                {
                    await UpdateModelAndUIFromScript(ScriptEditor.Text);                 
                });
            }
        }

        private async Task UpdateModelAndUIFromScript(string text)
        {
            try
            {
                if (selectedModel == null || !Models.ContainsKey(selectedModel))
                {
                    return;
                }

                // 데이터 처리 및 컴파일 수행
                var (cellChanges, invokeChanges) = UpdateModel(text);

                // UI 업데이트가 필요한 경우에만 수행
                if (cellChanges || invokeChanges)
                {
                    await UpdateUIAsync();
                }

                scriptChanged = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private (bool HasCellChanges, bool HasInvokeChanges) UpdateModel(string text)
        {
            // Scripts 딕셔너리 업데이트
            Scripts[selectedModel] = text;

            Model model = Models[selectedModel];

            // Update cellMatches
            var cellPattern = new Regex(@"(?:^//(?<description>.*?)\r?\n)?^(?<cellName>\w+)\s*--\s*(?<formula>.+)$", RegexOptions.Multiline);
            var newCellMatches = cellPattern.Matches(text);

            // Check for cell changes
            bool hasCellChanges = cellMatches == null ||
                                  cellMatches.Count != newCellMatches.Count ||
                                  !Enumerable.Range(0, cellMatches.Count)
                                             .All(i => cellMatches[i].Groups["cellName"].Value == newCellMatches[i].Groups["cellName"].Value &&
                                                       cellMatches[i].Groups["formula"].Value == newCellMatches[i].Groups["formula"].Value);

            bool hasCompiledCellChanges = false;

            if (hasCellChanges)
            {
                UpdateModelCells(newCellMatches);

                // 컴파일된 셀 이름 목록 생성
                var compiledCellNames = model.CompiledCells.Values
                    .Where(cell => cell.IsCompiled)
                    .Select(cell => cell.Name)
                    .ToHashSet();

                // Check for compiled cell changes
                hasCompiledCellChanges = cellMatches == null ||
                                         cellMatches.Count != newCellMatches.Count ||
                                         !Enumerable.Range(0, cellMatches.Count)
                                                    .Where(i => compiledCellNames.Contains(cellMatches[i].Groups["cellName"].Value))
                                                    .All(i => cellMatches[i].Groups["cellName"].Value == newCellMatches[i].Groups["cellName"].Value &&
                                                              cellMatches[i].Groups["formula"].Value == newCellMatches[i].Groups["formula"].Value);

                cellMatches = newCellMatches;
                cellMatchesDict[selectedModel] = cellMatches;
            }

            // Update invokeList
            var invokePattern = new Regex(@"^[ \t]*Invoke\((\w+),\s*(\d{1,5})\).*$", RegexOptions.Multiline);
            var newInvokeMatches = invokePattern.Matches(text);

            var newInvokeList = newInvokeMatches
                .Cast<Match>()
                .Select(m => (m.Groups[1].Value, int.Parse(m.Groups[2].Value)))
                .ToList();

            // Check for Invoke changes
            bool hasInvokeChanges = invokeList == null ||
                                    invokeList.Count != newInvokeList.Count ||
                                    !invokeList.SequenceEqual(newInvokeList);

            if (hasInvokeChanges || hasCompiledCellChanges)
            {
                invokeList = newInvokeList;
                UpdateInvokes();
            }

            return (hasCellChanges, hasInvokeChanges || hasCompiledCellChanges);
        }

        private async Task UpdateUIAsync()
        {
            await Dispatcher.InvokeAsync(() =>
            {
                UpdateCellList(selectedModel);
                UpdateSyntaxHighlighter();
                FoldingStrategy.UpdateFoldings(FoldingManager, ScriptEditor.Document);
            });
        }

        private async void LoadTemplate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string folderPath = FilePage.SelectedFolderPath;
                if (string.IsNullOrEmpty(folderPath))
                {
                    MessageBox.Show("먼저 FilePage에서 파일을 선택해주세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                LoadingOverlay.Visibility = Visibility.Visible;
                string scriptsBasePath = Path.Combine(folderPath, "Scripts");

                var productCode = App.ModelEngine.Context.Variables["ProductCode"].ToString();
                var riderCode = App.ModelEngine.Context.Variables["RiderCode"].ToString();

                App.ModelEngine.ProcessScriptTemplate(productCode, riderCode, scriptsBasePath);
                string autoScriptPath = Path.Combine(scriptsBasePath, "AutoScripts", $"{productCode}_{riderCode}.json");
                await LoadDataAsync(autoScriptPath);

                await UpdateUIAsync();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"템플릿 처리 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
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
                ModelsList.SelectedItem = newModelName;
                selectedModel = newModelName;

                ScriptEditor.Text = "";
                UpdateCellList(newModelName);
            }
        }

        private void UpdateCellList(string selectedModel)
        {
            if (Models.TryGetValue(selectedModel, out Model model))
            {
                var cellItems = new List<UIElement>();

                // cellMatches를 사용하여 셀 순서 유지
                if (cellMatches != null)
                {
                    foreach (Match match in cellMatches)
                    {
                        string cellName = match.Groups["cellName"].Value;
                        if (model.CompiledCells.TryGetValue(cellName, out CompiledCell cell))
                        {
                            var textBlock = new TextBlock { Text = cellName };

                            // 컴파일 실패한 셀은 빨간색으로 표시
                            if (!cell.IsCompiled)
                            {
                                textBlock.Foreground = Brushes.Red;
                                textBlock.FontWeight = FontWeights.Bold;
                            }
                            // Invoke 패턴에 포함된 셀은 보라색으로 표시
                            else if (invokeList?.Any(invoke => invoke.CellName == cellName) == true)
                            {
                                textBlock.Foreground = Brushes.LightGreen;
                            }

                            // 컨텍스트 메뉴 이벤트 추가
                            textBlock.MouseRightButtonUp += CellItem_MouseRightButtonUp;

                            cellItems.Add(textBlock);
                        }
                    }
                }

                CellsList.ItemsSource = cellItems;
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

        private void UpdateModelCells(MatchCollection matchCollection)
        {
            if (selectedModel == null || !Models.ContainsKey(selectedModel))
            {
                return;
            }

            Model model = Models[selectedModel];

            var newCells = new HashSet<string>();

            foreach (Match match in matchCollection)
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

        private void UpdateSyntaxHighlighter()
        {
            foreach (var model in Models.Values)
            {
                SyntaxHighlighter.UpdateModelCells(model.Name, selectedCell, model.CompiledCells.Keys);
            }

            SyntaxHighlighter.UpdateModels(selectedModel, Models.Keys);
            SyntaxHighlighter.UpdateContextVariables(App.ModelEngine.Context.Variables.Keys);
            SyntaxHighlighter.UpdateAssumptions(App.ModelEngine.Assumptions.Keys);
        }

        private async void ViewInExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ViewInExcelButton.IsEnabled = false;
                ViewInExcelButton.Content = "엑셀로 추줄중..";

                string folderPath = FilePage.SelectedFolderPath;
                if (string.IsNullOrEmpty(folderPath))
                {
                    MessageBox.Show("먼저 FilePage에서 파일을 선택해주세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string fileName = Path.GetFileNameWithoutExtension(folderPath);
                string samplesDirectory = Path.Combine(folderPath, "Samples");

                if (!Directory.Exists(samplesDirectory))
                {
                    Directory.CreateDirectory(samplesDirectory);
                }

                string filePath = Path.Combine(samplesDirectory, $"Sample_{fileName}.xlsx");

                // Models가 이미 정렬되어 있으므로, 그대로 SaveSampleToExcel 메서드에 전달
                await Task.Run(() => SampleExporter.SaveSampleToExcel(filePath));
                SampleExporter.RunLatestExcelFile(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel 파일 생성 또는 열기 중 오류 발생: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ViewInExcelButton.IsEnabled = true;
                ViewInExcelButton.Content = "엑셀로 보기";
            }
        }

        private void ModelsList_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            var listBoxItem = (e.OriginalSource as FrameworkElement)?.DataContext as string;
            if (listBoxItem != null)
            {
                ModelsList.SelectedItem = listBoxItem;
                ShowModelContextMenu(listBoxItem);
            }
        }

        private void ShowModelContextMenu(string modelName)
        {
            var contextMenu = new ContextMenu();

            var deleteMenuItem = new MenuItem { Header = "모델 삭제" };
            deleteMenuItem.Click += (sender, e) => DeleteModel(modelName);
            contextMenu.Items.Add(deleteMenuItem);

            var renameMenuItem = new MenuItem { Header = "모델 이름 변경" };
            renameMenuItem.Click += (sender, e) => RenameModel(modelName);
            contextMenu.Items.Add(renameMenuItem);

            contextMenu.IsOpen = true;
        }

        private void DeleteModel(string modelName)
        {
            var result = MessageBox.Show($"정말로 '{modelName}' 모델을 삭제하시겠습니까?",
                "모델 삭제 확인", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var modelsList = ModelsList.ItemsSource as List<string>;
                int currentIndex = modelsList.IndexOf(modelName);

                Models.Remove(modelName);
                Scripts.Remove(modelName);
                modelsList.Remove(modelName);

                // UI 업데이트
                ModelsList.ItemsSource = null;
                ModelsList.ItemsSource = modelsList;

                if (modelsList.Count > 0)
                {
                    ModelsList.SelectedIndex = Math.Min(currentIndex, modelsList.Count - 1);
                    selectedModel = modelsList[ModelsList.SelectedIndex];
                }
                else
                {
                    selectedModel = null;
                    ScriptEditor.Text = "";
                    CellsList.ItemsSource = null;
                }

                UpdateSyntaxHighlighter();
            }
        }

        private void RenameModel(string oldModelName)
        {
            var modelsList = ModelsList.ItemsSource as List<string>;
            int index = modelsList.IndexOf(oldModelName);

            var textBox = new TextBox();
            textBox.Text = oldModelName;
            textBox.SelectAll();

            textBox.LostFocus += OnRenameComplete;
            textBox.KeyDown += (s, e) => {
                if (e.Key == Key.Enter)
                {
                    OnRenameComplete(s, e);
                }
                else if (e.Key == Key.Escape)
                {
                    ModelsList.ItemsSource = null;
                    ModelsList.ItemsSource = modelsList;
                    ModelsList.SelectedIndex = index;
                }
            };

            var container = ModelsList.ItemContainerGenerator.ContainerFromIndex(index) as ListBoxItem;
            container.Content = textBox;
            textBox.Focus();

            void OnRenameComplete(object sender, RoutedEventArgs e)
            {
                var newModelName = textBox.Text.Trim();

                if (!string.IsNullOrWhiteSpace(newModelName) && !Models.ContainsKey(newModelName) && newModelName != oldModelName)
                {
                    Models[newModelName] = new Model(newModelName, App.ModelEngine);
                    Models.Remove(oldModelName);

                    Scripts[newModelName] = Scripts[oldModelName];
                    Scripts.Remove(oldModelName);

                    modelsList[index] = newModelName;
                    ModelsList.ItemsSource = null;
                    ModelsList.ItemsSource = modelsList;
                    ModelsList.SelectedIndex = index;
                    selectedModel = newModelName;

                    UpdateModelAndUIFromScript(Scripts[newModelName]);
                }
                else
                {
                    ModelsList.ItemsSource = null;
                    ModelsList.ItemsSource = modelsList;
                    ModelsList.SelectedIndex = index;
                }
            }
        }

        private void MoveModelUp_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedModel(-1);
        }

        private void MoveModelDown_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedModel(1);
        }

        private void MoveSelectedModel(int direction)
        {
            bool IsUpEnd = ModelsList.SelectedIndex == 0 && direction < 0;
            bool IsDownEnd = ModelsList.SelectedIndex == ModelsList.Items.Count - 1 && direction > 0;

            if (IsUpEnd || IsDownEnd)
            {
                return;
            }

            var modelsList = Models.Keys.ToList();
            int selectedIndex = ModelsList.SelectedIndex;
            int newIndex = selectedIndex + direction;

            if (newIndex >= 0 && newIndex < modelsList.Count)
            {
                string modelToMove = modelsList[selectedIndex];
                modelsList.RemoveAt(selectedIndex);
                modelsList.Insert(newIndex, modelToMove);

                // Models와 Scripts 딕셔너리 순서 업데이트
                var orderedModels = new Dictionary<string, Model>();
                var orderedScripts = new Dictionary<string, string>();
                foreach (var modelName in modelsList)
                {
                    orderedModels[modelName] = Models[modelName];
                    orderedScripts[modelName] = Scripts[modelName];
                }
                Models = orderedModels;
                Scripts = orderedScripts;

                // UI 업데이트
                ModelsList.ItemsSource = modelsList;
                ModelsList.SelectedIndex = newIndex;

                UpdateSyntaxHighlighter();
            }
        }

        private void ToggleAllFoldings()
        {
            if (FoldingManager == null) return;

            _areFoldingsCollapsed = !_areFoldingsCollapsed;

            foreach (var folding in FoldingManager.AllFoldings)
            {
                folding.IsFolded = _areFoldingsCollapsed;
            }

            ScriptEditor.TextArea.TextView.Redraw();
        }

        private void AdjustScriptEditorFontSize(bool increase)
        {
            double currentSize = ScriptEditor.FontSize;
            double newSize = increase ? currentSize + 1 : currentSize - 1;

            // 최소 8pt, 최대 24pt로 제한
            newSize = Math.Max(8, Math.Min(24, newSize));

            ScriptEditor.FontSize = newSize;
        }

        private void ViewSelectedModelPoint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var mainWindow = Application.Current.MainWindow as MainWindow;
                var modelPointPage = mainWindow?.pageCache["Pages/ModelPointPage.xaml"] as ModelPointPage;

                if (modelPointPage?.SelectedData != null)
                {
                    ShowSelectedModelPointWindow(modelPointPage.SelectedData);
                }
            }
            catch
            {
                MessageBox.Show("선택된 모델포인트가 없습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ShowSelectedModelPointWindow(List<object> selectedData)
        {
            var window = new Window
            {
                Title = "선택된 모델포인트",
                Width = 600,
                Height = 600,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            var dataGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true
            };

            var data = Enumerable.Range(0, App.ModelEngine.ModelPointInfo.Headers.Count)
                .Select(i => new
                {
                    Header = App.ModelEngine.ModelPointInfo.Headers[i],
                    Value = selectedData[i].ToString(),
                    Type = App.ModelEngine.ModelPointInfo.Types[i]
                })
                .ToList();

            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "헤더",
                Binding = new Binding("Header")
            });
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "값",
                Binding = new Binding("Value")
            });
            dataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "타입",
                Binding = new Binding("Type")
            });

            dataGrid.ItemsSource = data;
            window.Content = dataGrid;
            window.Show();
        }


        //Save 관련
        public void SaveScripts(bool IsAuto, int AutoNum = 0)
        {
            try
            {
                if (ModelsList.Items.Count == 0) return;

                string folderPath = FilePage.SelectedFolderPath;
                if (string.IsNullOrEmpty(folderPath))
                {
                    if (!IsAuto)
                    {
                        MessageBox.Show("먼저 엑셀 파일을 선택해주세요.", "파일 선택 필요", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    return;
                }

                string modelsFolder = Path.Combine(folderPath, "Scripts");
                if (!Directory.Exists(modelsFolder))
                {
                    Directory.CreateDirectory(modelsFolder);
                }

                // ModelsList의 순서대로 새로운 Scripts 딕셔너리 생성
                var orderedScripts = new Dictionary<string, string>();
                foreach (string modelName in ModelsList.Items)
                {
                    if (Scripts.ContainsKey(modelName))
                    {
                        orderedScripts[modelName] = Scripts[modelName];
                    }
                }

                string json = JsonConvert.SerializeObject(orderedScripts, Formatting.Indented);
                string fileName;

                if (IsAuto)
                {
                    // 자동 저장 모드
                    fileName = Path.Combine(modelsFolder, $"{Path.GetFileNameWithoutExtension(folderPath)}_scripts_auto{AutoNum}.json");
                    File.WriteAllText(fileName, json);
                }
                else
                {
                    // 수동 저장 모드
                    string baseFileName = Path.GetFileNameWithoutExtension(folderPath) + "_scripts";
                    int nextVersion = GetNextVersion(modelsFolder, baseFileName);
                    string defaultFileName = $"{baseFileName}_v{nextVersion}.json";

                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        InitialDirectory = modelsFolder,
                        Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                        FilterIndex = 1,
                        RestoreDirectory = true,
                        FileName = defaultFileName
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        fileName = saveFileDialog.FileName;
                        File.WriteAllText(fileName, json);
                        MessageBox.Show($"모델이 성공적으로 저장되었습니다: {fileName}", "저장 성공", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // 사용자가 저장을 취소한 경우
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!IsAuto)
                {
                    MessageBox.Show($"모델 저장 중 오류 발생: {ex.Message}", "저장 오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                // 자동 저장 모드에서는 오류 메시지를 표시하지 않습니다.
            }

            int GetNextVersion(string folder, string baseFileName)
            {
                int maxVersion = 0;
                string[] files = Directory.GetFiles(folder, $"{baseFileName}_v*.json");

                foreach (string file in files)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    if (fileName.EndsWith(".json"))
                    {
                        fileName = fileName.Substring(0, fileName.Length - 5);
                    }
                    string versionStr = fileName.Substring(fileName.LastIndexOf('v') + 1);
                    if (int.TryParse(versionStr, out int version))
                    {
                        maxVersion = Math.Max(maxVersion, version);
                    }
                }

                return maxVersion + 1;
            }
        }

        private void SetupAutoSaveTimer()
        {
            autoSaveTimer = new DispatcherTimer();
            autoSaveTimer.Interval = TimeSpan.FromSeconds(5);
            autoSaveTimer.Tick += AutoSaveTimer_Tick;
            autoSaveTimer.Start();
        }

        private void AutoSaveTimer_Tick(object sender, EventArgs e)
        {
            autoSaveTickCounter++;

            if (autoSaveTickCounter % 60 == 0)
            {
                SaveScripts(true, 2);
            }
            else
            {
                if (scriptChanged) SaveScripts(true, 1);
            }
            scriptChanged = false;
        }


        //키 다운 이벤트
        private void ScriptEditor_PreviewKeyDown_Save(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                try
                {
                    e.Handled = true;
                    SaveScripts(false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"모델 저장 중 오류 발생: {ex.Message}", "저장 오류", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ScriptEditor_PreviewKeyDown_Region(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.R && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                try
                {
                    e.Handled = true;
                    var document = ScriptEditor.Document;
                    var offset = ScriptEditor.CaretOffset;
                    var line = document.GetLineByOffset(offset);
                    var lineText = document.GetText(line.Offset, line.Length);
                    var indent = new string(' ', lineText.TakeWhile(char.IsWhiteSpace).Count());

                    var region = $"{indent}#region New Region\n{indent}\n{indent}#endregion";
                    document.Insert(line.Offset, region);

                    ScriptEditor.Select(line.Offset + indent.Length + 8, 10); // Select "New Region"
                }
                catch
                {

                }
            }
        }

        private void ScriptEditor_PreviewKeyDown_Folding(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.M && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                try
                {
                    e.Handled = true;
                    ToggleAllFoldings();
                }
                catch
                {

                }
            }
        }

        private void ScriptEditor_PreviewKeyDown_Zoom(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                try
                {
                    switch (e.Key)
                    {
                        case Key.OemPlus:
                        case Key.Add:
                            e.Handled = true;
                            AdjustScriptEditorFontSize(true);
                            break;

                        case Key.OemMinus:
                        case Key.Subtract:
                            e.Handled = true;
                            AdjustScriptEditorFontSize(false);
                            break;
                    }
                }
                catch
                {

                }

            }
        }

        private void ScriptEditor_PreviewKeyDown_Renew(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5)
            {
                try
                {
                    UpdateInvokes();
                }
                catch
                {

                }
            }
        }


        //탭 업데이트 
        public void UpdateInvokes()
        {
            if (Models != null && selectedModel != null && Models.TryGetValue(selectedModel, out var model))
            {
                try
                {
                    // Clear all model sheets
                    foreach (var m in Models.Values)
                    {
                        m.Clear();
                    }

                    // Process each Invoke call
                    foreach (var (cellName, t) in invokeList)
                    {
                        model.Invoke(cellName, t);
                    }
                }
                catch (Exception ex)
                {
                    // 에러 발생 시 모든 시트를 지우고 ErrorSheet만 생성
                    model.Sheets.Clear();
                    var errorSheet = new Sheet();
                    errorSheet.ResisterExpression("Error", App.ModelEngine.Context.CompileGeneric<double>("0"));
                    errorSheet.SetValue("Error", 0, 0);  // 에러 표시를 위한 더미 데이터

                    // 에러 메시지를 별도의 메서드로 저장
                    errorSheet.ResisterExpression("ErrorMessage", App.ModelEngine.Context.CompileGeneric<double>("0"));
                    errorSheet.SetValue("ErrorMessage", 0, 0);

                    model.Sheets["ErrorSheet"] = errorSheet;

                    // CellStatusTextBlock 업데이트
                    CellStatusTextBlock.Text = $"Error during Invoke: {ex.Message}";
                }
                finally
                {
                    SortSheets();
                    UpdateSheets();
                }
            }
            else
            {
                return;
            }
        }

        public void UpdateSheets()
        {
            string selectedTabName = (SheetTabControl.SelectedItem as TabItem)?.Header?.ToString();
            var currentTabs = new Dictionary<string, TabItem>();
            foreach (TabItem tab in SheetTabControl.Items)
            {
                currentTabs[tab.Header.ToString()] = tab;
            }

            var updatedTabs = new List<TabItem>();
            _tabData.Clear();

            // 현재 모델의 시트들을 먼저 처리
            if (Models.TryGetValue(selectedModel, out Model currentModel))
            {
                foreach (var sheetPair in currentModel.Sheets)
                {
                    string tabName = sheetPair.Key.Replace(";", ":");
                    _tabData[tabName] = (sheetPair.Value, false);

                    if (currentTabs.TryGetValue(tabName, out TabItem existingTab))
                    {
                        // 탭은 유지하되 내용은 지연 로딩을 위해 초기화
                        InitializeEmptyTab(tabName, existingTab);
                        updatedTabs.Add(existingTab);
                        currentTabs.Remove(tabName);
                    }
                    else
                    {
                        var newTab = new TabItem();
                        InitializeEmptyTab(tabName, newTab);
                        updatedTabs.Add(newTab);
                    }
                }
            }

            // 다른 모델의 시트들 처리
            foreach (var modelPair in Models.Where(m => m.Key != selectedModel))
            {
                foreach (var sheetPair in modelPair.Value.Sheets)
                {
                    string tabName = sheetPair.Key;
                    _tabData[tabName] = (sheetPair.Value, false);

                    if (currentTabs.TryGetValue(tabName, out TabItem existingTab))
                    {
                        InitializeEmptyTab(tabName, existingTab);
                        updatedTabs.Add(existingTab);
                        currentTabs.Remove(tabName);
                    }
                    else
                    {
                        var newTab = new TabItem();
                        InitializeEmptyTab(tabName, newTab);
                        updatedTabs.Add(newTab);
                    }
                }
            }

            Application.Current.Dispatcher.Invoke(() =>
            {
                foreach (var oldTab in currentTabs.Values)
                {
                    SheetTabControl.Items.Remove(oldTab);
                }

                for (int i = 0; i < updatedTabs.Count; i++)
                {
                    if (!SheetTabControl.Items.Contains(updatedTabs[i]))
                    {
                        SheetTabControl.Items.Add(updatedTabs[i]);
                    }
                    if (SheetTabControl.Items.IndexOf(updatedTabs[i]) != i)
                    {
                        SheetTabControl.Items.Remove(updatedTabs[i]);
                        SheetTabControl.Items.Insert(i, updatedTabs[i]);
                    }
                }

                // 이전 선택 탭 복원
                var tabToSelect = SheetTabControl.Items.Cast<TabItem>()
                    .FirstOrDefault(t => t.Header.ToString() == selectedTabName);
                if (tabToSelect != null)
                {
                    tabToSelect.IsSelected = true;
                    LoadTabData(tabToSelect); // 선택된 탭의 데이터 로드
                }
                else if (SheetTabControl.Items.Count > 0)
                {
                    (SheetTabControl.Items[0] as TabItem).IsSelected = true;
                    LoadTabData(SheetTabControl.Items[0] as TabItem);
                }
            });

            UpdateParameterGroups();
            HighlightHeaders();
        }

        private void InitializeEmptyTab(string tabName, TabItem tab)
        {
            tab.Header = tabName;
            var loadingGrid = new Grid();
            loadingGrid.Children.Add(new TextBlock
            {
                Text = "Click to load data",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            });
            tab.Content = loadingGrid;
        }

        private void LoadTabData(TabItem tab)
        {
            string tabName = tab.Header.ToString();
            if (!_tabData.TryGetValue(tabName, out var tabInfo) || tabInfo.IsLoaded)
                return;

            var dataGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserSortColumns = false
            };

            AddSheetTab(tabName, tabInfo.Sheet, tab);
            _tabData[tabName] = (tabInfo.Sheet, true);
        }

        private void SheetTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && e.AddedItems[0] is TabItem newTab)
            {
                LoadTabData(newTab);
            }
        }

        private void AddSheetTab(string sheetName, Sheet sheet, TabItem tabItem)
        {
            DataGrid dataGrid;
            if (tabItem.Content is DataGrid existingGrid)
            {
                dataGrid = existingGrid;
                dataGrid.Columns.Clear();
                dataGrid.ItemsSource = null;
            }
            else
            {
                dataGrid = new DataGrid
                {
                    AutoGenerateColumns = false,
                    IsReadOnly = true,
                    CanUserSortColumns = false,
                    FrozenColumnCount = 1
                };
                tabItem.Content = dataGrid;
            }

            if (sheetName == "ErrorSheet")
            {
                // ErrorSheet의 경우 에러 메시지를 표시
                dataGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = "Error Message",
                    Binding = new System.Windows.Data.Binding("[ErrorMessage]")
                });

                var errorMessage = CellStatusTextBlock.Text;
                var data = new List<Dictionary<string, string>>
                {
                    new Dictionary<string, string> { ["ErrorMessage"] = errorMessage }
                };

                dataGrid.ItemsSource = data;
            }
            else
            {
                var sheetData = sheet.GetAllData();
                var rows = sheetData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                if (rows.Length > 0)
                {
                    var headers = rows[0].Split('\t');

                    foreach (var header in headers)
                    {
                        var column = new DataGridTextColumn()
                        {
                            Header = new TextBlock { Text = header },
                            Binding = new Binding($"[{header}]")
                        };

                        var binding = new Binding($"[{header}]");
                        binding.Converter = new SignificantDigitsConverter(SignificantDigits);
                        column.Binding = binding;

                        dataGrid.Columns.Add(column);
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
            }
            App.ApplyTheme();
            tabItem.Header = sheetName;
        }

        private void SortSheets()
        {
            var sortOption = App.SettingsManager.CurrentSettings.DataGridSortOption;

            foreach (var model in Models.Values)
            {
                foreach (Sheet sheet in model.Sheets.Values)
                {
                    switch (sortOption)
                    {
                        case DataGridSortOption.CellDefinitionOrder:
                            SortSheetByCellDefinition(model.Name, sheet);
                            break;
                        case DataGridSortOption.Alphabetical:
                            sheet.SortCache(key => key);
                            break;
                            // Default case: 기존 순서 유지
                    }
                }
            }
        }

        private void SortSheetByCellDefinition(string modelName, Sheet sheet)
        {
            if (!cellMatchesDict.ContainsKey(modelName))
            {
                return;
            }

            var cellOrder = cellMatchesDict[modelName].Cast<Match>()
                .Select(m => m.Groups["cellName"].Value)
                .ToList();

            sheet.SortCache(key =>
            {
                int index = cellOrder.IndexOf(key);
                return index >= 0 ? index : int.MaxValue;
            });
        }

        private void HighlightHeaders()
        {
            if (SheetTabControl.SelectedItem is TabItem selectedTab &&
                selectedTab.Content is DataGrid dataGrid &&
                invokeList != null && invokeList.Any())
            {
                var cellNamesToHighlight = invokeList.Select(invoke => invoke.CellName).ToHashSet();
                DataGridColumn firstHighlightedColumn = null;

                foreach (var column in dataGrid.Columns)
                {
                    if (column is DataGridTextColumn textColumn)
                    {
                        var header = textColumn.Header as TextBlock;
                        if (header != null && cellNamesToHighlight.Contains(header.Text))
                        {
                            header.FontWeight = FontWeights.Bold;
                            column.CellStyle = new Style(typeof(DataGridCell))
                            {
                                Setters =
                        {
                            new Setter(DataGridCell.FontWeightProperty, FontWeights.Bold),
                            new Setter(DataGridCell.ForegroundProperty, Brushes.Red)
                        }
                            };
                            firstHighlightedColumn = firstHighlightedColumn ?? column;
                        }
                    }
                }

                if (firstHighlightedColumn != null && dataGrid.Items.Count > 0)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                        dataGrid.ScrollIntoView(dataGrid.Items[0], firstHighlightedColumn)));
                }
            }
        }


        // Invoke 패턴 추가 메서드
        private void CellItem_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            var textBlock = sender as TextBlock;
            if (textBlock == null) return;

            var contextMenu = new ContextMenu();
            var cellName = textBlock.Text;

            // 추가 메뉴 아이템
            var addMenuItem = new MenuItem { Header = "Invoke 추가" };
            addMenuItem.Click += (s, args) => AddInvokePattern(cellName);
            contextMenu.Items.Add(addMenuItem);

            // 삭제 메뉴 아이템 (Invoke 패턴에 있는 경우에만 활성화)
            if (invokeList?.Any(invoke => invoke.CellName == cellName) == true)
            {
                var removeMenuItem = new MenuItem { Header = "Invoke 삭제" };
                removeMenuItem.Click += (s, args) => RemoveInvokePattern(cellName);
                contextMenu.Items.Add(removeMenuItem);
            }

            contextMenu.IsOpen = true;
            e.Handled = true;
        }

        private void AddInvokePattern(string cellName)
        {
            var document = ScriptEditor.Document;
            string invokePattern = $"Invoke({cellName}, 0)";

            // 마지막 줄에 개행 문자가 없으면 추가
            if (!document.Text.EndsWith(Environment.NewLine))
            {
                document.Insert(document.TextLength, Environment.NewLine);
            }

            // Invoke 패턴 추가
            document.Insert(document.TextLength, invokePattern + Environment.NewLine);
        }

        private void RemoveInvokePattern(string cellName)
        {
            var document = ScriptEditor.Document;
            var text = document.Text;
            var lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            // Invoke 패턴에 맞는 라인 찾기
            var invokePattern = new Regex($@"^[ \t]*Invoke\({cellName},\s*\d+\).*$");
            var newLines = lines.Where(line => !invokePattern.IsMatch(line)).ToList();

            // 문서 업데이트
            document.Text = string.Join(Environment.NewLine, newLines);
        }


        //탭 컨트롤 필터
        private void ParameterGroupComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ParameterGroupComboBox.SelectedItem == null) return;

            string selectedGroup = ParameterGroupComboBox.SelectedItem.ToString();
            UpdateTabVisibility(selectedGroup);
        }

        private void UpdateTabVisibility(string selectedGroup)
        {
            foreach (TabItem tab in SheetTabControl.Items)
            {
                string tabName = tab.Header.ToString();
                string group = GetParameterGroup(tabName);

                if (selectedGroup == "All")
                {
                    tab.Visibility = Visibility.Visible;
                }
                else if (selectedGroup == "Base")
                {
                    tab.Visibility = group == "Base" ? Visibility.Visible : Visibility.Collapsed;
                }
                else
                {
                    tab.Visibility = group == selectedGroup ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }

        private void UpdateParameterGroups()
        {
            var groups = new HashSet<string> { "All", "Base" };

            foreach (TabItem tab in SheetTabControl.Items)
            {
                string tabName = tab.Header.ToString();
                string group = GetParameterGroup(tabName);
                if (!string.IsNullOrEmpty(group) && group != "Base")
                {
                    groups.Add(group);
                }
            }

            // 현재 선택된 항목 저장
            var currentSelection = ParameterGroupComboBox.SelectedItem?.ToString();

            // ComboBox 아이템 업데이트 (SelectionChanged 이벤트 발생 방지)
            ParameterGroupComboBox.SelectionChanged -= ParameterGroupComboBox_SelectionChanged;

            var orderedGroups = groups.OrderBy(g => g == "All" ? 0 : g == "Base" ? 1 : 2)
                                     .ThenBy(g => g.Split(':')[0])
                                     .ThenBy(g => g.Replace(g.Split(':')[0] + ":", "").PadLeft(10, '0'));

            ParameterGroupComboBox.ItemsSource = orderedGroups;

            // 선택 상태 처리
            if (!string.IsNullOrEmpty(currentSelection))
            {
                if (currentSelection == "All" || currentSelection == "Base")
                {
                    ParameterGroupComboBox.SelectedItem = currentSelection;
                    UpdateTabVisibility(currentSelection); // 명시적으로 탭 가시성 업데이트
                }
                else if (groups.Contains(currentSelection))
                {
                    ParameterGroupComboBox.SelectedItem = currentSelection;
                    UpdateTabVisibility(currentSelection);
                }
                else
                {
                    ParameterGroupComboBox.SelectedItem = "Base";
                    UpdateTabVisibility("Base");
                }
            }
            else
            {
                ParameterGroupComboBox.SelectedItem = "Base";
                UpdateTabVisibility("Base");
            }

            // SelectionChanged 이벤트 핸들러 다시 연결
            ParameterGroupComboBox.SelectionChanged += ParameterGroupComboBox_SelectionChanged;
        }

        private string GetParameterGroup(string tabName)
        {
            var match = Regex.Match(tabName, @"{([^}]+)}");
            if (!match.Success) return "Base";

            // 전체 파라미터를 그룹으로 사용
            return match.Groups[1].Value.Trim();
        }
    }

    public class SignificantDigitsConverter : IValueConverter
    {
        private readonly int _significantDigits;

        public SignificantDigitsConverter(int significantDigits)
        {
            _significantDigits = significantDigits;
        }

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is string stringValue && double.TryParse(stringValue, out double doubleValue))
            {
                return FormatToSignificantDigits(doubleValue, _significantDigits);
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        private string FormatToSignificantDigits(double value, int significantDigits)
        {
            if (double.IsNaN(value))
                return "NaN";

            if (value == 0)
                return "0";

            if (double.IsInfinity(value))
                return value > 0 ? "Infinity" : "-Infinity";

            // 값의 절대값을 취하고 지수 표기법으로 변환
            string scientificNotation = Math.Abs(value).ToString($"E{significantDigits - 1}");
            string[] parts = scientificNotation.Split('E');
            double coefficient = double.Parse(parts[0]);
            int exponent = int.Parse(parts[1]);

            // 정수 부분의 자릿수 계산
            int integerPartDigits = Math.Max(1, exponent + 1);
            if (integerPartDigits >= significantDigits)
            {
                // 정수 부분이 유효숫자 이상인 경우, 정수 부분만 표시
                return Math.Round(value, 0).ToString("F0");
            }
            else
            {
                // 소수점 이하 자릿수 계산
                int decimalPlaces = Math.Max(0, significantDigits - integerPartDigits);
                // 반올림 후 문자열로 변환
                string roundedValue = Math.Round(value, decimalPlaces).ToString($"F{decimalPlaces}");
                // 끝의 불필요한 0 제거
                roundedValue = roundedValue.TrimEnd('0');
                if (roundedValue.EndsWith("."))
                    roundedValue = roundedValue.TrimEnd('.');
                return roundedValue;
            }
        }
    }

    public class RegionFoldingStrategy
    {
        private static readonly Regex regionStartRegex = new Regex(@"^\s*#region\s+(.*)$", RegexOptions.Compiled);
        private static readonly Regex regionEndRegex = new Regex(@"^\s*#endregion", RegexOptions.Compiled);

        public void UpdateFoldings(FoldingManager manager, TextDocument document)
        {
            IEnumerable<NewFolding> newFoldings = CreateNewFoldings(document, out int firstErrorOffset);
            manager.UpdateFoldings(newFoldings, firstErrorOffset);
        }

        public IEnumerable<NewFolding> CreateNewFoldings(TextDocument document, out int firstErrorOffset)
        {
            firstErrorOffset = -1;
            List<NewFolding> newFoldings = new List<NewFolding>();

            Stack<int> startOffsets = new Stack<int>();
            Stack<string> names = new Stack<string>();

            foreach (DocumentLine line in document.Lines)
            {
                string text = document.GetText(line);
                Match startMatch = regionStartRegex.Match(text);
                if (startMatch.Success)
                {
                    startOffsets.Push(line.Offset);
                    names.Push(startMatch.Groups[1].Value.Trim());
                }
                else
                {
                    Match endMatch = regionEndRegex.Match(text);
                    if (endMatch.Success && startOffsets.Count > 0)
                    {
                        int startOffset = startOffsets.Pop();
                        string name = names.Pop();
                        newFoldings.Add(new NewFolding(startOffset, line.EndOffset) { Name = name });
                    }
                }
            }

            newFoldings.Sort((a, b) => a.StartOffset.CompareTo(b.StartOffset));
            return newFoldings;
        }
    }
}