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
using System.Text.RegularExpressions;
using System.Security.Policy;

namespace ActuLight.Pages
{
    public partial class DataProcessingPage : Page
    {
        private ObservableCollection<FileEntry> fileEntries;
        private Dictionary<string, LTFReader> readers;
        private Dictionary<string, FileCache> fileCache;
        private bool isProcessing;
        private ExpressionContext context;

        public DataProcessingPage()
        {
            InitializeComponent();
            InitializeCollections();
            InitializeExpressionContext();
            SetupInitialState();
            SetupTextBoxHandlers();
        }

        private void InitializeCollections()
        {
            fileEntries = new ObservableCollection<FileEntry>();
            readers = new Dictionary<string, LTFReader>();
            fileCache = new Dictionary<string, FileCache>(); // Initialize cache
            FilesGrid.ItemsSource = fileEntries;
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

        private void SetupTextBoxHandlers()
        {
            for (int i = 1; i <= 8; i++)
            {
                var textBox = FindName($"A{i}TextBox") as TextBox;
                if (textBox != null)
                {
                    textBox.TextChanged += (s, e) =>
                    {
                        UpdateVariableValue((TextBox)s);
                        if (FilesGrid.SelectedItem is FileEntry selectedEntry)
                        {
                            UpdateSampleLines(selectedEntry.Path);
                        }
                    };
                }
            }
        }

        private void SetupInitialState()
        {
            ProcessTypeComboBox.SelectedIndex = 0;
            UpdateProcessControls();
        }

        private void ProcessTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateProcessControls();
        }

        private void UpdateProcessControls()
        {
            ProcessSpecificControls.Children.Clear();
            var label = new Label { Content = GetProcessSpecificLabel() };
            Grid.SetColumn(label, 0);

            var textBox = new TextBox { Width = 300, HorizontalAlignment = HorizontalAlignment.Left };
            Grid.SetColumn(textBox, 1);

            ProcessSpecificControls.Children.Add(label);
            ProcessSpecificControls.Children.Add(textBox);
        }

        private void UpdateVariableValue(TextBox textBox)
        {
            var match = Regex.Match(textBox.Name, @"A(\d)TextBox");
            if (match.Success)
            {
                string varName = $"a{match.Groups[1].Value}";
                if (int.TryParse(textBox.Text, out int value))
                {
                    context.Variables[varName] = value;
                }
                else if (textBox.Text.Contains(","))
                {
                    var parts = textBox.Text.Split(',');
                    if (parts.Length == 2 &&
                        int.TryParse(parts[0], out int start) &&
                        int.TryParse(parts[1], out int length))
                    {
                        context.Variables[varName] = new[] { start, length };
                    }
                }
            }
        }

        private string GetProcessSpecificLabel()
        {
            if (ProcessTypeComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                switch (selectedItem.Content.ToString())
                {
                    case "Split": return "Split By:";
                    case "Filter": return "Filter Expression:";
                    case "Count": return "Count By:";
                    case "Distinct": return "Distinct By:";
                    default: return "Expression:";
                }
            }
            return "Expression:";
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (fileEntries.Count == 0)
            {
                MessageBox.Show("Please add files to process.", "No Files", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                isProcessing = true;
                StartButton.IsEnabled = false;
                CancelButton.IsEnabled = true;

                foreach (var entry in fileEntries)
                {
                    if (!readers.TryGetValue(entry.Path, out var reader))
                    {
                        reader = new LTFReader(entry.Path);
                        readers[entry.Path] = reader;
                    }

                    var processType = (ProcessTypeComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                    var expressionStr = (ProcessSpecificControls.Children[1] as TextBox).Text;

                    if (string.IsNullOrWhiteSpace(expressionStr))
                    {
                        throw new Exception("수식을 입력해주세요");
                    }

                    IGenericExpression<string> expression = context.CompileGeneric<string>((ProcessSpecificControls.Children[1] as TextBox).Text);
                    
                    switch (processType)
                    {
                        case "Split":
                            LTFProcessor.Split(reader, line => GetExpressionValue(line, expression));
                            break;
                        case "Filter":
                            LTFProcessor.Filter(reader, line => EvaluateFilterExpression(line, expression));
                            break;
                        case "Count":
                            LTFProcessor.Count(reader, line => GetExpressionValue(line, expression));
                            break;
                        case "Distinct":
                            LTFProcessor.Distinct(reader, line => GetExpressionValue(line, expression));
                            break;
                    }

                    if (!isProcessing) break;
                }

                if (isProcessing)
                {
                    MessageBox.Show("Processing completed successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during processing: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                isProcessing = false;
                StartButton.IsEnabled = true;
                CancelButton.IsEnabled = false;
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            isProcessing = false;
            foreach (var reader in readers.Values)
            {
                reader.IsCanceled = true;
            }
        }

        private void FilesGrid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    if (File.Exists(file))
                    {
                        try
                        {
                            var reader = new LTFReader(file);
                            readers[file] = reader;

                            // Read and cache sample lines
                            var sampleLines = File.ReadLines(file).Take(100).ToList();
                            fileCache[file] = new FileCache(file, sampleLines, reader.Delimiter);

                            var entry = new FileEntry
                            {
                                Number = fileEntries.Count + 1,
                                Path = file,
                                Key = Path.GetFileNameWithoutExtension(file)
                            };
                            fileEntries.Add(entry);

                            DelimiterTextBox.Text = FormatDelimiter(reader.Delimiter);
                            FilesGrid.SelectedItem = entry;
                            UpdateSampleLines(file);
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
                // 선택된 파일의 구분자 표시
                if (readers.TryGetValue(selectedEntry.Path, out var reader))
                {
                    DelimiterTextBox.Text = reader.Delimiter;
                }
                UpdateSampleLines(selectedEntry.Path);
            }
        }

        private void UpdateSampleLines(string filePath)
        {
            try
            {
                if (!fileCache.TryGetValue(filePath, out var cache))
                {
                    throw new Exception("File cache not found");
                }

                var document = new FlowDocument
                {
                    PageWidth = fileCache[filePath].SampleLines.Max(x => x.Length) * 8,
                    PagePadding = new Thickness(0)
                };
                var paragraph = new Paragraph();
                document.Blocks.Add(paragraph);

                var highlights = GetHighlightRules();
                var hasDelimiter = cache.Delimiter != "None" && !string.IsNullOrEmpty(cache.Delimiter);

                foreach (var line in cache.SampleLines)
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
                    if (fieldIndex > 0 && fieldIndex <= parts.Length)
                    {
                        int startPos = 0;
                        for (int j = 0; j < fieldIndex - 1; j++)
                        {
                            startPos += parts[j].Length + delimiter.Length;
                        }

                        int length = parts[fieldIndex - 1].Length;

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

            var reader = readers[selectedFile.Path];
            var hasDelimiter = !string.IsNullOrEmpty(reader.Delimiter);

            for (int i = 1; i <= 8; i++)
            {
                var textBox = FindName($"A{i}TextBox") as TextBox;
                if (textBox != null && !string.IsNullOrWhiteSpace(textBox.Text))
                {
                    string varName = $"a{i}";
                    string value = "";

                    if (hasDelimiter)
                    {
                        if (int.TryParse(textBox.Text, out int fieldIndex))
                        {
                            var parts = line.Split(new[] { reader.Delimiter }, StringSplitOptions.None);
                            if (fieldIndex > 0 && fieldIndex <= parts.Length)
                            {
                                value = parts[fieldIndex - 1].Trim();
                            }
                        }
                    }
                    else
                    {
                        if (int.TryParse(textBox.Text, out int position) &&
                            position > 0 && position <= line.Length)
                        {
                            value = line.Substring(position - 1, 1);
                        }
                    }

                    context.Variables[varName] = value;
                }
            }
        }

        private string GetExpressionValue(string line, IGenericExpression<string> expression)
        {
            // 현재 라인에 대해 변수 설정
            SetVariablesForLine(line);

            try
            {
                // 표현식 평가
                return EvaluateExpression(expression);
            }
            catch (Exception ex)
            {
                return $"Expression Error: {ex.Message}";
            }
        }

        private bool EvaluateFilterExpression(string line, IGenericExpression<string> expression)
        {
            // Implement filter expression evaluation logic here
            // This is a simplified example
            return true;
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

        private string EvaluateExpression(IGenericExpression<string> expression)
        {
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