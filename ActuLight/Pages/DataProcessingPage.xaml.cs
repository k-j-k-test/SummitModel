using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ActuLiteModel;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Documents;
using Flee.PublicTypes;
using System.Collections.Specialized;
using System.Threading.Tasks;
using System.Web;
using System.Diagnostics;
using ModernWpf.Controls;
using System.Text.RegularExpressions;

namespace ActuLight.Pages
{
    public partial class DataProcessingPage : System.Windows.Controls.Page
    {
        private ObservableCollection<FileEntry> fileEntries;
        public Dictionary<string, LTFReader> readers;
        private Dictionary<string, FileCache> fileCache;
        private ExpressionContext context;
        private bool isProcessing;

        //타이머 관련 필드
        private System.Windows.Threading.DispatcherTimer processTimer;
        private LTFReader currentReader;
        private int currentFileIndex;
        private int totalFiles;

        public DataProcessingPage()
        {
            InitializeComponent();
            InitializeCollections();
            InitializeExpressionContext();
            InitializeContextMenu();
            SetupTextBoxHandlers();

            FleeFunc.Readers = readers;
        }

        private void InitializeCollections()
        {
            EmptyMessage.Visibility = Visibility.Visible;

            fileEntries = new ObservableCollection<FileEntry>();
            readers = new Dictionary<string, LTFReader>();
            fileCache = new Dictionary<string, FileCache>(); // Initialize cache
            FilesGrid.ItemsSource = fileEntries;

            fileEntries.CollectionChanged += (s, e) =>
            {
                EmptyMessage.Visibility = fileEntries.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            };

            // FileEntry의 Key 속성 변경 이벤트 구독
            ((INotifyCollectionChanged)FilesGrid.Items).CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null)
                {
                    foreach (FileEntry entry in e.NewItems)
                    {
                        entry.PropertyChanged += (sender, args) =>
                        {
                            if (args.PropertyName == nameof(FileEntry.Key) && sender is FileEntry fileEntry)
                            {
                                UpdateSampleKeys(fileEntry);
                            }
                        };
                    }
                }
            };

            processTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };

            processTimer.Tick += ProcessTimer_Tick;
        }

        private void InitializeExpressionContext()
        {
            context = new ExpressionContext();
            context.Imports.AddType(typeof(TextParser));

            // a1~a8 변수를 string으로 초기화
            for (int i = 1; i <= 8; i++)
            {
                context.Variables[$"a{i}"] = string.Empty;
            }
        }

        private void InitializeContextMenu()
        {
            var contextMenu = new ContextMenu();

            var openFolderMenuItem = new MenuItem { Header = "폴더로 이동" };
            openFolderMenuItem.Click += (s, e) => {
                if (FilesGrid.SelectedItem is FileEntry selectedEntry)
                {
                    string baseDirectory = LTFProcessor.GetBaseDirectory(selectedEntry.Path);
                    if (Directory.Exists(baseDirectory))
                    {
                        System.Diagnostics.Process.Start("explorer.exe", baseDirectory);
                    }
                    else
                    {
                        MessageBox.Show("처리 폴더가 아직 생성되지 않았습니다.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            };

            var deleteMenuItem = new MenuItem { Header = "삭제" };
            deleteMenuItem.Click += (s, e) => {
                if (FilesGrid.SelectedItems.Count > 0)
                {
                    var selectedEntries = FilesGrid.SelectedItems.Cast<FileEntry>().ToList();
                    foreach (var entry in selectedEntries)
                    {
                        fileEntries.Remove(entry);
                        readers.Remove(Path.GetFileName(entry.Path));
                        fileCache.Remove(Path.GetFileName(entry.Path));
                    }

                    //update file number
                    for (int i = 0; i < fileEntries.Count; i++)
                    {
                        fileEntries[i].Number = i + 1;
                    }

                    // SampleLinesTextBox 내용 지우기
                    SampleLinesTextBox.Document = new FlowDocument();
                }
            };

            var deleteAllMenuItem = new MenuItem { Header = "전체 삭제" };
            deleteAllMenuItem.Click += (s, e) => {
                fileEntries.Clear();
                readers.Clear();
                fileCache.Clear();

                // SampleLinesTextBox 내용 지우기
                SampleLinesTextBox.Document = new FlowDocument();
            };

            contextMenu.Items.Add(openFolderMenuItem);
            contextMenu.Items.Add(deleteMenuItem);
            contextMenu.Items.Add(deleteAllMenuItem);

            FilesGrid.ContextMenu = contextMenu;
        }

        private void SetupTextBoxHandlers()
        {
            ProcessTypeComboBox.SelectedIndex = 0;

            for (int i = 1; i <= 8; i++)
            {
                var textBox = FindName($"A{i}TextBox") as TextBox;
                if (textBox != null)
                {
                    textBox.TextChanged += (s, e) =>
                    {
                        if (FilesGrid.SelectedItem is FileEntry selectedEntry)
                        {
                            UpdateSampleLines(selectedEntry.Path);
                        }
                    };
                }
            }

            SampleLinesTextBox.PreviewMouseLeftButtonDown += UpdateCaretPosition;
            SampleLinesTextBox.SelectionChanged += UpdateCaretPosition;
        }

        private void ProcessTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (SearchTextBox != null)
            //{
            //    var comboBox = sender as ComboBox;
            //    var selectedItem = comboBox?.SelectedItem as ComboBoxItem;

            //    // Enable SearchTextBox only when "Index" is selected
            //    SearchTextBox.IsEnabled = selectedItem?.Content.ToString() == "Index";

            //    // Clear the search text when switching away from Index
            //    if (!SearchTextBox.IsEnabled)
            //    {
            //        SearchTextBox.Text = string.Empty;
            //        SearchTextBox.ItemsSource = null;
            //    }
            //}
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (FilesGrid.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select files to process.", "No Files Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                isProcessing = true;
                StartButton.IsEnabled = false;
                CancelButton.IsEnabled = true;
                StatusTextBox.Text = "0%";

                var filesToProcess = FilesGrid.SelectedItems.Cast<FileEntry>().ToList();
                totalFiles = filesToProcess.Count;
                processTimer.Start();

                for (currentFileIndex = 0; currentFileIndex < filesToProcess.Count; currentFileIndex++)
                {
                    if (!isProcessing) break;

                    var entry = filesToProcess[currentFileIndex];
                    if (string.IsNullOrWhiteSpace(entry.Key))
                    {
                        throw new Exception($"Key is required for file: {entry.Path}");
                    }

                    FilesGrid.SelectedItems.Clear();
                    FilesGrid.SelectedItem = entry;

                    string fileName = Path.GetFileName(entry.Path);
                    if (!readers.TryGetValue(fileName, out var reader))
                    {
                        reader = new LTFReader(entry.Path);
                        readers[fileName] = reader;
                    }

                    reader.IsCanceled = false;
                    reader.Progress = 0;
                    currentReader = reader;
                    var processType = ProcessTypeComboBox.Text;
                    var expression = context.CompileGeneric<string>(entry.Key);

                    // KeySelector 설정
                    reader.KeySelector = line => GetExpressionValue(line, expression);

                    // Index 처리인 경우에만 KeySelectorExpression 설정
                    if (processType == "Index")
                    {
                        string keyExpression = entry.Key;
                        for (int i = 1; i <= 8; i++)
                        {
                            var textBox = FindName($"A{i}TextBox") as TextBox;
                            if (textBox != null && !string.IsNullOrWhiteSpace(textBox.Text))
                            {
                                string pattern = $@"\ba{i}\b";
                                string value = textBox.Text;

                                if (value.Contains(",") || int.TryParse(value, out int fieldIndex))
                                {
                                    keyExpression = Regex.Replace(keyExpression, pattern, $"sub({value})");
                                }
                            }
                        }
                        reader.KeySelectorExpression = keyExpression;
                    }

                    await Task.Run(() =>
                    {
                        switch (processType)
                        {
                            case "Split":
                                LTFProcessor.Split(reader, line => GetExpressionValue(line, expression));
                                break;
                            case "Count":
                                LTFProcessor.Count(reader, line => GetExpressionValue(line, expression));
                                break;
                            case "Distinct":
                                LTFProcessor.Distinct(reader, line => GetExpressionValue(line, expression));
                                break;
                            case "Index":
                                reader.IndexPath = reader.GetIndexPath();
                                reader.FileLastWriteTime = File.GetLastWriteTime(reader.FilePath);
                                reader.KeySelectorHash = reader.GetKeySelectorHash(reader.KeySelector);
                                reader.LoadIndex();
                                break;
                        }
                    });
                }

                if (isProcessing)
                {
                    StatusTextBox.Text = $"{currentFileIndex}/{filesToProcess.Count} - Completed";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during processing: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusTextBox.Text = "Error";
            }
            finally
            {
                isProcessing = false;
                StartButton.IsEnabled = true;
                CancelButton.IsEnabled = false;
                processTimer.Stop();
                currentReader = null;
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            isProcessing = false;
            foreach (var reader in readers.Values)
            {
                reader.IsCanceled = true;
            }
            StatusTextBox.Text = $"Cancelled at {currentFileIndex + 1}/{totalFiles} - {(currentReader.Progress * 100):F1}%";
        }

        public async void LoadExternalButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LoadExternalButton.IsEnabled = false;
                string externalDataFolder = Path.Combine(FilePage.SelectedFolderPath, "ExternalData");

                if (!Directory.Exists(externalDataFolder))
                {
                    MessageBox.Show($"외부 데이터 폴더가 존재하지 않습니다: {externalDataFolder}",
                        "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 폴더 내의 모든 파일 가져오기 (하위 폴더 제외)
                var files = Directory.GetFiles(externalDataFolder);

                foreach (string file in files)
                {
                    // 이미 로드된 파일인지 확인
                    if (readers.ContainsKey(file))
                        continue;

                    await Task.Run(() => LoadExternalFile(file));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"외부 데이터 로드 중 오류 발생: {ex.Message}",
                    "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                LoadExternalButton.IsEnabled = true;
            }
        }

        public void LoadExternalFile(string file)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    string fileName = Path.GetFileName(file);

                    // 이미 같은 파일명이 있는지 확인
                    if (readers.ContainsKey(fileName))
                    {
                        return;
                    }

                    var reader = new LTFReader(file);
                    readers[fileName] = reader;

                    // 인덱스 파일 경로 설정 및 로드 시도
                    reader.IndexPath = reader.GetIndexPath();
                    reader.FileLastWriteTime = File.GetLastWriteTime(file);

                    if (reader.LoadIndexWithoutKeySelector())
                    {
                        if (!string.IsNullOrEmpty(reader.KeySelectorExpression))
                        {
                            var expression = context.CompileGeneric<string>(reader.KeySelectorExpression);
                            reader.KeySelector = line => GetExpressionValue(line, expression);
                        }
                    }

                    var sampleLines = File.ReadLines(file).Take(100).ToList();
                    fileCache[fileName] = new FileCache(file, sampleLines, reader.Delimiter);

                    var entry = new FileEntry
                    {
                        Number = fileEntries.Count + 1,
                        Path = file,
                        Key = reader.KeySelectorExpression ?? string.Empty,
                    };
                    fileEntries.Add(entry);

                    DelimiterTextBox.Text = FormatDelimiter(reader.Delimiter);
                    FilesGrid.SelectedItem = entry;
                    UpdateSampleLines(fileName);
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"파일 '{Path.GetFileName(file)}' 처리 중 오류 발생: {ex.Message}",
                        "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private void FilesGrid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files.OrderBy(x => x))
                {
                    if (File.Exists(file))
                    {
                        try
                        {
                            string fileName = Path.GetFileName(file);

                            // 이미 같은 파일명이 있는지 확인
                            if (readers.ContainsKey(fileName))
                            {
                                MessageBox.Show($"이미 같은 이름의 파일이 있습니다: {fileName}",
                                    "중복 파일", MessageBoxButton.OK, MessageBoxImage.Warning);
                                continue;
                            }

                            var reader = new LTFReader(file);
                            readers[fileName] = reader;

                            // Read and cache sample lines
                            var sampleLines = File.ReadLines(file).Take(100).ToList();
                            fileCache[fileName] = new FileCache(file, sampleLines, reader.Delimiter);

                            var entry = new FileEntry
                            {
                                Number = fileEntries.Count + 1,
                                Path = file,
                                Key = "a1"
                            };
                            fileEntries.Add(entry);

                            DelimiterTextBox.Text = FormatDelimiter(reader.Delimiter);
                            FilesGrid.SelectedItem = entry;
                            UpdateSampleLines(fileName);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error processing file {file}: {ex.Message}",
                                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
        }

        private void FilesGrid_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void FilesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FilesGrid.SelectedItem is FileEntry selectedEntry)
            {
                string fileName = Path.GetFileName(selectedEntry.Path);
                if (readers.TryGetValue(fileName, out var reader))
                {
                    DelimiterTextBox.Text = FormatDelimiter(reader.Delimiter);
                    SkipTextBox.Text = reader.SkipLines.ToString();
                }
                UpdateSampleLines(fileName);
                UpdateSampleKeys(selectedEntry);
            }
        }

        private void FilesGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // Enter 키 이벤트를 처리했다고 표시
                e.Handled = true;

                // 현재 셀의 편집을 완료
                DataGrid grid = (DataGrid)sender;
                TextBox cell = Keyboard.FocusedElement as TextBox;

                grid.CommitEdit();

                Dispatcher.InvokeAsync(() =>
                {
                    if (cell != null)
                    {
                        // 또는 ContentPresenter를 통해 값 가져오기
                        var selectedItem = (FileEntry)grid.SelectedItem;
                        selectedItem.Key = cell.Text;
                        UpdateSampleKeys(selectedItem);
                    }             
                });
            }
        }

        private void SkipTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (FilesGrid.SelectedItem is FileEntry selectedEntry)
            {
                string fileName = Path.GetFileName(selectedEntry.Path);
                if (readers.TryGetValue(fileName, out var reader))
                {
                    if (int.TryParse(SkipTextBox.Text, out int skipLines) && skipLines >= 0)
                    {
                        reader.SkipLines = skipLines;
                        UpdateSampleLines(fileName);
                    }
                }
            }
        }

        private void SearchTextBox_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
        {
            if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            {
                try
                {
                    if (FilesGrid.SelectedItem is FileEntry selectedEntry)
                    {
                        string fileName = Path.GetFileName(selectedEntry.Path);
                        if (readers.TryGetValue(fileName, out var reader))
                        {
                            var searchText = sender.Text.ToLower();
                            if (string.IsNullOrEmpty(searchText))
                            {
                                sender.ItemsSource = null;
                                return;
                            }

                            var sortedKeys = reader.Index.Keys.ToList();

                            // 이진 탐색으로 시작 위치 찾기
                            int index = sortedKeys.BinarySearch(searchText);
                            if (index < 0)
                                index = ~index;

                            var suggestions = new List<string>();
                            while (index < sortedKeys.Count && suggestions.Count < 100)
                            {
                                string key = sortedKeys[index];
                                if (!key.ToLower().StartsWith(searchText))
                                    break;

                                suggestions.Add(key);
                                index++;
                            }

                            sender.ItemsSource = suggestions;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error during autocomplete: {ex.Message}");
                }
            }
        }

        private void SearchTextBox_SuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
        {
            if (args.SelectedItem != null)
            {
                sender.Text = args.SelectedItem.ToString();
            }
        }

        private void SearchTextBox_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
        {
            string searchText;
            if (args.ChosenSuggestion != null)
            {
                // 사용자가 제안된 항목을 선택한 경우
                searchText = args.ChosenSuggestion.ToString();
            }
            else
            {
                // 사용자가 Enter를 누른 경우
                searchText = args.QueryText;
            }

            UpdateSearchResults(searchText);
        }

        private void UpdateCaretPosition(object sender, RoutedEventArgs e)
        {
            try
            {
                var caretPosition = SampleLinesTextBox.CaretPosition;
                if (caretPosition != null)
                {
                    var text = new TextRange(SampleLinesTextBox.Document.ContentStart, caretPosition).Text;

                    var lineNumber = text.Count(c => c == '\n') + 1;

                    var lastNewLine = text.LastIndexOf('\n');
                    var columnNumber = lastNewLine == -1 ? text.Length : text.Length - lastNewLine - 1;

                    CursorPositionText.Text = $"Line: {lineNumber}, Column: {columnNumber}";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating caret position: {ex.Message}");
                CursorPositionText.Text = "Line: -, Column: -";
            }
        }

        private void UpdateSearchResults(string searchText)
        {
            if (FilesGrid.SelectedItem is FileEntry selectedEntry &&
                !string.IsNullOrWhiteSpace(searchText))
            {
                string fileName = Path.GetFileName(selectedEntry.Path);
                if (readers.TryGetValue(fileName, out var reader))
                {
                    try
                    {
                        var searchResults = reader.GetLines(searchText, 100);

                        var document = new FlowDocument
                        {
                            PageWidth = searchResults.Count > 0 ?
                                searchResults.Max(x => x.Length) * 8 + searchResults.First().Count(x => x == '\t') * 32 : 400,
                            PagePadding = new Thickness(0)
                        };

                        var paragraph = new Paragraph();
                        document.Blocks.Add(paragraph);

                        var highlights = GetHighlightRules();
                        var hasDelimiter = reader.Delimiter != "None" && !string.IsNullOrEmpty(reader.Delimiter);

                        // Sample Keys 문서 준비
                        var keysDocument = new FlowDocument
                        {
                            PageWidth = 2000,
                            PagePadding = new Thickness(0)
                        };
                        var keysParagraph = new Paragraph();
                        keysDocument.Blocks.Add(keysParagraph);

                        // Key 식을 컴파일
                        IGenericExpression<string> expression = null;
                        if (!string.IsNullOrWhiteSpace(selectedEntry.Key))
                        {
                            expression = context.CompileGeneric<string>(selectedEntry.Key);
                        }

                        foreach (var line in searchResults)
                        {
                            // SampleLines 업데이트
                            AddHighlightedLine(paragraph, line, highlights, hasDelimiter, reader.Delimiter);
                            paragraph.Inlines.Add(new Run(Environment.NewLine));

                            // SampleKeys 업데이트
                            if (expression != null)
                            {
                                string keyValue = GetExpressionValue(line, expression);
                                keysParagraph.Inlines.Add(new Run(keyValue));
                                keysParagraph.Inlines.Add(new Run(Environment.NewLine));
                            }
                        }

                        if (searchResults.Count == 0)
                        {
                            paragraph.Inlines.Add(new Run("No results found."));
                            keysParagraph.Inlines.Add(new Run("No results found."));
                        }
                        else if (searchResults.Count == 100)
                        {
                            paragraph.Inlines.Add(new Run("Showing first 100 results..."));
                            keysParagraph.Inlines.Add(new Run("Showing first 100 results..."));
                        }

                        SampleLinesTextBox.Document = document;
                        SampleKeysTextBox.Document = keysDocument;
                    }
                    catch (Exception ex)
                    {
                        var errorMessage = $"Error during search: {ex.Message}";

                        // SampleLinesTextBox용 document
                        var errorDocument1 = new FlowDocument();
                        errorDocument1.Blocks.Add(new Paragraph(new Run(errorMessage)));
                        SampleLinesTextBox.Document = errorDocument1;

                        // SampleKeysTextBox용 document
                        var errorDocument2 = new FlowDocument();
                        errorDocument2.Blocks.Add(new Paragraph(new Run(errorMessage)));
                        SampleKeysTextBox.Document = errorDocument2;
                    }
                }
            }
        }

        private void UpdateSampleLines(string fileName)
        {
            try
            {
                if (!fileCache.TryGetValue(Path.GetFileName(fileName), out var cache))
                {
                    throw new Exception("File cache not found");
                }

                var document = new FlowDocument
                {
                    PageWidth = fileCache[Path.GetFileName(fileName)].SampleLines.Max(x => x.Length) * 8 + fileCache[Path.GetFileName(fileName)].SampleLines.First().Count(x => x == '\t') * 32,
                    PagePadding = new Thickness(0)
                };
                var paragraph = new Paragraph();
                document.Blocks.Add(paragraph);

                var highlights = GetHighlightRules();
                var hasDelimiter = cache.Delimiter != "None" && !string.IsNullOrEmpty(cache.Delimiter);

                // 현재 Skip 값 가져오기
                int skipLines = 0;
                if (readers.TryGetValue(Path.GetFileName(fileName), out var reader))
                {
                    skipLines = reader.SkipLines;
                }

                // Skip된 라인 표시
                if (skipLines > 0)
                {
                    for (int i = 0; i < Math.Min(skipLines, cache.SampleLines.Count); i++)
                    {
                        var skippedRun = new Run(cache.SampleLines[i])
                        {
                            Foreground = System.Windows.Media.Brushes.Gray,
                            TextDecorations = TextDecorations.Strikethrough
                        };
                        paragraph.Inlines.Add(skippedRun);
                        paragraph.Inlines.Add(new Run(Environment.NewLine));
                    }
                }

                // Skip 이후의 라인 표시
                foreach (var line in cache.SampleLines.Skip(skipLines))
                {
                    AddHighlightedLine(paragraph, line, highlights, hasDelimiter, cache.Delimiter);
                    paragraph.Inlines.Add(new Run(Environment.NewLine));
                }

                SampleLinesTextBox.Document = document;
            }
            catch (Exception ex)
            {
                var errorDocument = new FlowDocument();
                errorDocument.Blocks.Add(new Paragraph(new Run($"Error reading file: {ex.Message}")));
                SampleLinesTextBox.Document = errorDocument;
            }
        }

        private void UpdateSampleKeys(FileEntry entry)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(entry.Key))
                {
                    SampleKeysTextBox.Document = new FlowDocument();
                    return;
                }

                string fileName = Path.GetFileName(entry.Path);
                if (!fileCache.TryGetValue(fileName, out var cache))
                {
                    throw new Exception("File cache not found");
                }

                var document = new FlowDocument
                {
                    PageWidth = 2000,
                    PagePadding = new Thickness(0)
                };
                var paragraph = new Paragraph();
                document.Blocks.Add(paragraph);

                IGenericExpression<string> expression = context.CompileGeneric<string>(entry.Key);

                foreach (var line in cache.SampleLines)
                {
                    string keyValue = GetExpressionValue(line, expression);
                    paragraph.Inlines.Add(new Run(keyValue));
                    paragraph.Inlines.Add(new Run(Environment.NewLine));
                }

                SampleKeysTextBox.Document = document;
            }
            catch (Exception ex)
            {
                var errorDocument = new FlowDocument();
                errorDocument.Blocks.Add(new Paragraph(new Run($"Error calculating keys: {ex.Message}")));
                SampleKeysTextBox.Document = errorDocument;
            }
        }

        private void ProcessTimer_Tick(object sender, EventArgs e)
        {
            if (!isProcessing || currentReader == null)
            {
                processTimer.Stop();
                return;
            }

            StatusTextBox.Text = $"{currentFileIndex + 1}/{totalFiles} - {(currentReader.Progress * 100):F1}%";
        }

        private List<(string text, string style)> GetHighlightRules()
        {
            var rules = new List<(string text, string style)>();

            for (int i = 1; i <= 8; i++)
            {
                var textBox = FindName($"A{i}TextBox") as TextBox;
                if (textBox != null && !string.IsNullOrWhiteSpace(textBox.Text))
                {
                    rules.Add((textBox.Text, $"HighlightStyle{i}"));
                }
            }

            return rules;
        }

        private void AddHighlightedLine(Paragraph paragraph, string line, List<(string text, string style)> highlights, bool hasDelimiter, string delimiter)
        {
            // Run 인스턴스 목록을 관리하여 스타일 변경 가능
            var inlines = new List<Run>();

            // 라인의 전체 내용을 한 번에 추가
            var baseRun = new Run(line);
            inlines.Add(baseRun);

            // 하이라이트 규칙 처리
            foreach (var (text, style) in highlights)
            {
                if (hasDelimiter && int.TryParse(text, out int fieldIndex))
                {
                    // 구분자가 있는 경우, 필드의 특정 영역 선택
                    var parts = line.Split(new[] { delimiter }, StringSplitOptions.None);
                    if (fieldIndex >= 0 && fieldIndex <= parts.Length)
                    {
                        int startPos = 0;
                        for (int j = 0; j < fieldIndex ; j++)
                        {
                            startPos += parts[j].Length + delimiter.Length;
                        }

                        int length = parts[fieldIndex ].Length;

                        // 해당 위치의 텍스트에 스타일 적용
                        ApplyStyleToRange(inlines, startPos, length, style);
                    }
                }
                else if (text.Contains(","))
                {
                    // 구분자가 없고 범위 기반 규칙인 경우
                    var parts = text.Split(',');
                    if (parts.Length == 2 &&
                        int.TryParse(parts[0], out int start) &&
                        int.TryParse(parts[1], out int length) &&
                        start >= 0 && start + length <= line.Length)
                    {
                        ApplyStyleToRange(inlines, start, length, style);
                    }
                }
                else if (int.TryParse(text, out int position))
                {
                    // 특정 위치에 단일 문자 강조
                    if (position >= 0 && position < line.Length)
                    {
                        ApplyStyleToRange(inlines, position, 1, style);
                    }
                }
            }

            // 최종적으로 모든 `Run` 객체를 Paragraph에 추가
            foreach (var run in inlines)
            {
                paragraph.Inlines.Add(run);
            }
        }

        private void ApplyStyleToRange(List<Run> inlines, int start, int length, string style)
        {
            int currentPos = 0;

            for (int i = 0; i < inlines.Count; i++)
            {
                var run = inlines[i];
                int runLength = run.Text.Length;

                if (currentPos + runLength <= start)
                {
                    // 현재 Run은 하이라이트 범위 이전에 위치
                    currentPos += runLength;
                    continue;
                }

                if (currentPos >= start + length)
                {
                    // 현재 Run은 하이라이트 범위를 지난 후
                    break;
                }

                // 하이라이트 범위와 Run의 겹치는 부분 계산
                int localStart = Math.Max(0, start - currentPos);
                int localEnd = Math.Min(runLength, start + length - currentPos);
                int localLength = localEnd - localStart;

                // 현재 Run을 하이라이트 범위 기준으로 분할
                if (localStart > 0)
                {
                    var beforeRun = new Run(run.Text.Substring(0, localStart))
                    {
                        Style = run.Style
                    };
                    inlines.Insert(i, beforeRun);
                    i++;
                }

                var highlightedRun = new Run(run.Text.Substring(localStart, localLength))
                {
                    Style = SampleLinesTextBox.Resources[style] as Style
                };
                inlines.Insert(i, highlightedRun);
                i++;

                if (localEnd < runLength)
                {
                    var afterRun = new Run(run.Text.Substring(localEnd))
                    {
                        Style = run.Style
                    };
                    inlines.Insert(i, afterRun);
                    i++;
                }

                // 원래 Run 제거
                inlines.RemoveAt(i);
                i--;

                currentPos += runLength;
            }
        }

        private void SetVariablesForLine(string line)
        {
            if (!(FilesGrid.SelectedItem is FileEntry selectedFile))
                return;

            var reader = readers[Path.GetFileName(selectedFile.Path)];
            var hasDelimiter = !string.IsNullOrEmpty(reader.Delimiter);

            for (int i = 1; i <= 8; i++)
            {
                var textBox = FindName($"A{i}TextBox") as TextBox;
                if (textBox != null && !string.IsNullOrWhiteSpace(textBox.Text))
                {
                    string varName = $"a{i}";
                    string value = "";

                    string textBoxContent = textBox.Text;

                    if (textBoxContent.Contains(","))
                    {
                        // Substring 처리
                        var range = textBoxContent.Split(',');
                        if (range.Length == 2 &&
                            int.TryParse(range[0], out int start) &&
                            int.TryParse(range[1], out int length) &&
                            start >= 0 && length > 0 && start + length <= line.Length)
                        {
                            value = line.Substring(start, length).Trim();
                        }
                    }
                    else if (int.TryParse(textBoxContent, out int fieldOrPosition))
                    {
                        if (hasDelimiter)
                        {
                            // Delimiter 기반 처리
                            var parts = line.Split(new[] { reader.Delimiter }, StringSplitOptions.None);
                            if (fieldOrPosition >= 0 && fieldOrPosition <= parts.Length)
                            {
                                value = parts[fieldOrPosition].Trim();
                            }
                        }
                    }

                    context.Variables[varName] = value;
                }
            }
        }

        private string GetExpressionValue(string line, IGenericExpression<string> expression)
        {
            // 현재 라인에 대해 변수 설정
            Dispatcher.Invoke(() =>
            {
                var selectedFile = FilesGrid.SelectedItem as FileEntry;
                var delimiter = readers[Path.GetFileName(selectedFile.Path)].Delimiter;
                TextParser.SetLine(line, delimiter);
                SetVariablesForLine(line);
            });

            try
            {
                var result = expression.Evaluate();
                return result?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                return $"Expression Error: {ex.Message}";
            }
        }

        private string FormatDelimiter(string delimiter)
        {
            return delimiter switch
            {
                "\t" => "Tab",
                "," => "Comma (,)",
                "|" => "Pipe (|)",
                ";" => "Semicolon (;)",
                " " => "Space",
                "" => "None",
                null => "None",
                _ => $"Custom ({delimiter})"
            };
        }
    }

    public class FileEntry : INotifyPropertyChanged
    {
        private int number;
        private string path;
        private string key;

        public int Number
        {
            get => number;
            set
            {
                number = value;
                OnPropertyChanged(nameof(Number));
            }
        }

        public string Path
        {
            get => path;
            set
            {
                path = value;
                OnPropertyChanged(nameof(Path));
            }
        }

        public string Key
        {
            get => key;
            set
            {
                key = value;
                OnPropertyChanged(nameof(Key));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class FileCache
    {
        public string FilePath { get; }
        public List<string> SampleLines { get; }
        public string Delimiter { get; }

        public FileCache(string filePath, List<string> sampleLines, string delimiter)
        {
            FilePath = filePath;
            SampleLines = sampleLines;
            Delimiter = delimiter;
        }
    }
}