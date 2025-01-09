using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.CodeCompletion;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Editing;
using ICSharpCode.AvalonEdit.Rendering;
using ActuLiteModel;
using System.IO;
using ModernWpf;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Windows.Media.Animation;
using System.Windows.Documents;

namespace ActuLight
{
    public class SyntaxHighlighter : DocumentColorizingTransformer
    {
        public static Dictionary<string, (Color Dark, Color Light, Color HighContrast)> ColorSchemes;

        private string _currentModel;
        private string _currentCell;

        private CustomCompletionWindow _completionWindow;
        private readonly TextEditor _textEditor;

        private Dictionary<string, HashSet<string>> _modelCells = new Dictionary<string, HashSet<string>>();
        private HashSet<string> _models = new HashSet<string>();
        private HashSet<string> _contextVariables = new HashSet<string>();
        private HashSet<string> _assumptions = new HashSet<string>();
        private HashSet<string> _expenses = typeof(Input_exp)
            .GetProperties()
            .Where(p => p.Name.StartsWith("Alpha") ||
                        p.Name.StartsWith("Beta") ||
                        p.Name.StartsWith("Gamma") ||
                        p.Name.StartsWith("Ce") ||
                        p.Name.StartsWith("Refund") ||
                        p.Name.StartsWith("Etc"))
            .Select(p => p.Name)
            .ToHashSet();

        private List<string> _functions = new List<string>();
        private Dictionary<string, string> _cellCompletions = new Dictionary<string, string>();

        public SyntaxHighlighter(TextEditor textEditor)
        {
            LoadAutoCompletionSet();

            _textEditor = textEditor;
            _textEditor.TextChanged += TextArea_TextEntered;

            // 테마 변경 이벤트 구독
            ThemeManager.Current.ActualApplicationThemeChanged += Current_ActualApplicationThemeChanged;

            // TextEditor 스타일 설정
            SetTextEditorStyle();

            ColorSchemes = new Dictionary<string, (Color Dark, Color Light, Color HighContrast)>
            {
                ["Comment"] = (Colors.LightGreen, Colors.Green, Colors.LimeGreen),
                ["Separator"] = (Colors.LightGray, Colors.Gray, Colors.DarkGray),
                ["Function"] = (Colors.Khaki, Colors.DarkKhaki, Colors.PaleGoldenrod),
                ["Number"] = (Colors.LightSeaGreen, Colors.MediumSeaGreen, Colors.Aquamarine),
                ["String"] = (Colors.LightSalmon, Colors.Salmon, Colors.Coral),
                ["Boolean"] = (Colors.SkyBlue, Colors.DeepSkyBlue, Colors.LightSkyBlue),
                ["Model"] = (Colors.PaleTurquoise, Colors.MediumTurquoise, Colors.Turquoise),
                ["Cell"] = (Colors.Plum, Colors.MediumPurple, Colors.Violet),
                ["ContextVariable"] = (Colors.PaleGreen, Colors.MediumAquamarine, Colors.LightGreen)
            };
        }

        //Show AutoCompletion Window
        public void TextArea_TextEntered(object sender, EventArgs e)
        {
            if (!(sender is TextEditor textEditor)) return;

            var textArea = textEditor.TextArea;
            var (currentLine, textBeforeCaret, currentWord) = GetCurrentLineInfo(textArea);

            var completionData = GetCompletionData(textBeforeCaret, currentLine, currentWord);

            if (completionData.Any())
            {
                ShowCompletionWindow(textArea, completionData);
            }
            else
            {
                _completionWindow?.Close();
                _completionWindow = null;
            }
        }

        private List<CustomCompletionData> GetCompletionData(string textBeforeCaret, string currentLine, string currentWord)
        {
             var completionData = new List<CustomCompletionData>();

            if (currentLine.Contains("--"))
            {
                var parts = currentLine.Split(new[] { "--" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2 && parts[1].Trim() == "")
                {
                    if (_cellCompletions.TryGetValue(parts[0].Trim(), out string completion))
                    {
                        completionData.Add(new CustomCompletionData(completion, CompletionType.CellAutoCompletion, parts[0]));
                    }
                }
            }

            if (string.IsNullOrEmpty(currentWord) || !currentLine.Contains("--"))
            {
                return completionData;
            }

            // Check if we're inside Assum function
            var assumMatch = Regex.Match(textBeforeCaret, @"Assum\([""']([^""']*)$");
            if (assumMatch.Success)
            {
                // We're inside Assum function, offer assumption completions
                completionData.AddRange(GetFilteredCompletionData(_assumptions, currentWord, CompletionType.Assumption));
            }

            // Check if we're inside Exp function
            var expMatch = Regex.Match(textBeforeCaret, @"Exp\([""']([^""']*)$");
            if (expMatch.Success)
            {
                // We're inside Exp function, offer expense completions
                completionData.AddRange(GetFilteredCompletionData(_expenses, currentWord, CompletionType.Expense));
            }

            else
            {
                // Check if we're after a model name and a dot
                var modelMatch = Regex.Match(textBeforeCaret, @"(\w+)({[^}]*})?\.([\w]*)$");
                if (modelMatch.Success)
                {
                    string modelName = modelMatch.Groups[1].Value;
                    string parameters = modelMatch.Groups[2].Value; // This will be empty if no parameters
                    string cellPrefix = modelMatch.Groups[3].Value;
                    if (_modelCells.TryGetValue(modelName, out var cells))
                    {
                        completionData.AddRange(GetFilteredCompletionData(cells, cellPrefix, CompletionType.CellReference));
                    }
                }
                else
                {
                    // Add cell completions
                    if (!string.IsNullOrEmpty(_currentModel) && _modelCells.TryGetValue(_currentModel, out var currentModelCells))
                    {
                        completionData.AddRange(GetFilteredCompletionData(currentModelCells, currentWord, CompletionType.CellReference));
                    }

                    // Add function completions
                    completionData.AddRange(GetFilteredCompletionData(_functions, currentWord, CompletionType.Function));

                    // Add model completions
                    completionData.AddRange(GetFilteredCompletionData(_models, currentWord, CompletionType.Model));

                    // Add context variable completions
                    completionData.AddRange(GetFilteredCompletionData(_contextVariables, currentWord, CompletionType.ContextVariable));
                }
            }

            return completionData.OrderBy(item => item.Text.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase) ? 0 : 1).ThenBy(item => item.Text).ToList();
        }

        private List<CustomCompletionData> GetFilteredCompletionData(IEnumerable<string> items, string filter, CompletionType type)
        {
            var filteredItems = items
                .Where(item => item.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                .Take(10)
                .Select(item => new CustomCompletionData(item, type, filter))
                .ToList();

            return filteredItems;
        }

        private void ShowCompletionWindow(TextArea textArea, IEnumerable<ICompletionData> completionData)
        {
            if (_completionWindow != null)
            {
                _completionWindow.Close();
            }

            _completionWindow = new CustomCompletionWindow(textArea);

            var (textBeforeCaret, currentLine, currentWord) = GetCurrentLineInfo(textArea);

            var dataList = _completionWindow.CompletionList.CompletionData;

            if (Regex.Match(currentLine, @"(\w+)\s*--\s*$").Success)
            {
                _completionWindow.StartOffset = textArea.Caret.Offset - currentWord.Length - 1;
            }
            else
            {
                _completionWindow.StartOffset = textArea.Caret.Offset - currentWord.Length;
            }


            foreach (var data in completionData)
            {
                dataList.Add(data);
            }

            if (dataList.Count > 0)
            {
                _completionWindow.Show();

                _completionWindow.Closed += (sender, args) => _completionWindow = null;
            }
            else
            {
                _completionWindow.Close();
                _completionWindow = null;
            }
        }

        private (string CurrentLine, string TextBeforeCaret, string CurrentWord) GetCurrentLineInfo(TextArea textArea)
        {
            var document = textArea.Document;
            var currentLine = document.GetLineByOffset(textArea.Caret.Offset);
            var lineStartOffset = currentLine.Offset;
            var lineEndOffset = currentLine.EndOffset;
            var caretOffsetInLine = textArea.Caret.Offset - lineStartOffset;

            // 현재 라인 전체 텍스트
            string currentLineText = document.GetText(lineStartOffset, currentLine.Length);

            // 캐럿 이전의 텍스트
            string textBeforeCaret = document.GetText(lineStartOffset, caretOffsetInLine);

            // 현재 단어 (캐럿 이전의 마지막 단어)
            string currentWord = "";
            var match = Regex.Match(textBeforeCaret, @"[\w.]+$");
            if (match.Success)
            {
                currentWord = match.Value;
            }

            return (currentLineText, textBeforeCaret, currentWord);
        }

        //Colorize
        protected override void ColorizeLine(DocumentLine line)
        {
            string text = CurrentContext.Document.GetText(line);

            // 주석 처리 먼저 적용
            int commentIndex = text.IndexOf("//");
            if (commentIndex >= 0)
            {
                ApplyCommentColorRule(line, commentIndex, text.Length - commentIndex);
                // 주석 이후 부분은 다른 규칙 적용하지 않음
                text = text.Substring(0, commentIndex);
            }

            ApplyColorRule(line, text, RegexPatterns.Comment, "Comment");
            ApplyColorRule(line, text, RegexPatterns.Separator, "Separator");
            ApplyColorRule(line, text, RegexPatterns.Number, "Number");
            ApplyColorRule(line, text, RegexPatterns.String, "String");
            ApplyColorRule(line, text, RegexPatterns.Boolean, "Boolean");

            // 동적으로 생성되는 정규표현식 사용
            ApplyColorRule(line, text, RegexPatterns.GetFunctionRegex(_functions), "Function");
            ApplyColorRule(line, text, RegexPatterns.GetModelRegex(_models), "Model");
            ApplyColorRule(line, text, RegexPatterns.GetContextVariableRegex(_contextVariables), "ContextVariable");

            ApplyCurrentModelCellColorRule(line, text);
            ApplyOtherModelCellColorRule(line, text);
            ApplyCellReferenceValidation(line, text);
        }

        private void ApplyColorRule(DocumentLine line, string text, Regex regex, string colorKey)
        {
            var brush = GetThemeColor(colorKey);

            foreach (Match match in regex.Matches(text))
            {
                ChangeLinePart(
                    line.Offset + match.Index,
                    line.Offset + match.Index + match.Length,
                    element => element.TextRunProperties.SetForegroundBrush(brush)
                );
            }
        }

        private void ApplyCurrentModelCellColorRule(DocumentLine line, string text)
        {
            if (string.IsNullOrEmpty(_currentModel) || !_modelCells.TryGetValue(_currentModel, out var currentModelCells))
                return;

            var cellBrush = GetThemeColor("Cell");
            string pattern = $@"\b({string.Join("|", currentModelCells.Select(Regex.Escape))})\b";
            var regex = new Regex(pattern);

            foreach (Match match in regex.Matches(text))
            {
                int startOffset = line.Offset + match.Index;
                int endOffset = startOffset + match.Length;
                ChangeLinePart(startOffset, endOffset, element => element.TextRunProperties.SetForegroundBrush(cellBrush));
            }
        }

        private void ApplyOtherModelCellColorRule(DocumentLine line, string text)
        {
            var cellBrush = GetThemeColor("Cell");

            foreach (var modelEntry in _modelCells.Where(entry => entry.Key != _currentModel))
            {
                string modelName = modelEntry.Key;
                HashSet<string> cells = modelEntry.Value;

                string pattern = $@"\b{Regex.Escape(modelName)}\.({string.Join("|", cells.Select(Regex.Escape))})\b";
                var regex = new Regex(pattern);

                foreach (Match match in regex.Matches(text))
                {
                    int startOffset = line.Offset + match.Index + modelName.Length + 1; // +1 for the dot
                    int endOffset = startOffset + match.Groups[1].Length;
                    ChangeLinePart(startOffset, endOffset, element => element.TextRunProperties.SetForegroundBrush(cellBrush));
                }
            }
        }

        private void ApplyCellReferenceValidation(DocumentLine line, string text)
        {
            var errorBrush = new SolidColorBrush(Colors.Red);
            errorBrush.Freeze();
            var cellBrush = GetThemeColor("Cell");
            var modelBrush = GetThemeColor("Model");

            foreach (Match match in RegexPatterns.Cell.Matches(text))
            {
                string fullMatch = match.Value;
                int modelEndIndex = fullMatch.IndexOf('{');
                if (modelEndIndex == -1) modelEndIndex = fullMatch.IndexOf('.');
                if (modelEndIndex == -1) modelEndIndex = fullMatch.IndexOf('[');

                string modelName = fullMatch.Substring(0, modelEndIndex);
                string remaining = fullMatch.Substring(modelEndIndex);

                string parameters = "";
                string cellName = "";

                if (remaining.StartsWith("{"))
                {
                    int paramEndIndex = remaining.IndexOf('}');
                    parameters = remaining.Substring(0, paramEndIndex + 1);
                    remaining = remaining.Substring(paramEndIndex + 1);
                }

                if (remaining.StartsWith("."))
                {
                    cellName = remaining.Substring(1, remaining.Length - 2);
                }

                int modelStartIndex = match.Index;
                int modelLength = modelName.Length + parameters.Length;

                // 여기서부터는 기존 로직과 동일
                if (string.IsNullOrEmpty(cellName)) // 현재 모델의 셀 참조
                {
                    cellName = modelName;
                    if (_currentModel != null && _modelCells.TryGetValue(_currentModel, out var currentModelCells))
                    {
                        if (currentModelCells.Contains(cellName))
                        {
                            ApplyBrush(line, modelStartIndex, modelLength, cellBrush);
                        }
                        else
                        {
                            ApplyErrorStyle(line, modelStartIndex, modelLength, errorBrush);
                        }
                    }
                }
                else // 다른 모델의 셀 참조
                {
                    if (_models.Contains(modelName))
                    {
                        ApplyBrush(line, modelStartIndex, modelLength, modelBrush);

                        int cellStartIndex = modelStartIndex + modelLength + 1; // +1 for the dot
                        if (_modelCells.TryGetValue(modelName, out var modelCells) && modelCells.Contains(cellName))
                        {
                            ApplyBrush(line, cellStartIndex, cellName.Length, cellBrush);
                        }
                        else
                        {
                            ApplyErrorStyle(line, cellStartIndex, cellName.Length, errorBrush);
                        }
                    }
                    else
                    {
                        ApplyErrorStyle(line, modelStartIndex, modelLength + 1 + cellName.Length, errorBrush);
                    }
                }
            }
        }

        private void ApplyBrush(DocumentLine line, int startIndex, int length, SolidColorBrush brush)
        {
            int startOffset = line.Offset + startIndex;
            int endOffset = startOffset + length;
            ChangeLinePart(startOffset, endOffset, element => element.TextRunProperties.SetForegroundBrush(brush));
        }

        private void ApplyErrorStyle(DocumentLine line, int startIndex, int length, SolidColorBrush brush)
        {
            int startOffset = line.Offset + startIndex;
            int endOffset = startOffset + length;
            ChangeLinePart(startOffset, endOffset, element =>
            {
                element.TextRunProperties.SetTextDecorations(TextDecorations.Underline);
                element.TextRunProperties.SetForegroundBrush(brush);
            });
        }

        private void ApplyCommentColorRule(DocumentLine line, int startOffset, int length)
        {
            ChangeLinePart(
                line.Offset + startOffset,
                line.Offset + startOffset + length,
                element =>
                {
                    element.TextRunProperties.SetForegroundBrush(GetThemeColor("Comment"));
                    // 주석 부분에 대해 다른 스타일 적용 (예: 기울임꼴)
                    element.TextRunProperties.SetTypeface(new Typeface(
                        element.TextRunProperties.Typeface.FontFamily,
                        FontStyles.Italic,
                        element.TextRunProperties.Typeface.Weight,
                        element.TextRunProperties.Typeface.Stretch
                    ));
                }
            );
        }

        private SolidColorBrush GetThemeColor(string colorKey)
        {
            var theme = ThemeManager.Current.ActualApplicationTheme;
            var (dark, light, highContrast) = ColorSchemes[colorKey];

            return theme switch
            {
                ApplicationTheme.Dark => new SolidColorBrush(dark),
                ApplicationTheme.Light => new SolidColorBrush(light),
                _ => new SolidColorBrush(highContrast)
            };
        }

        //AutoCompletion Update
        public void UpdateModels(string modelName, IEnumerable<string> models)
        {
            _currentModel = modelName;
            _models = new HashSet<string>(models);
            _textEditor.TextArea.TextView.Redraw();
        }

        public void UpdateModelCells(string modelName, string cellName, IEnumerable<string> cells)
        {         
            if (!_modelCells.ContainsKey(modelName))
            {
                _modelCells[modelName] = new HashSet<string>();
            }
            _modelCells[modelName] = new HashSet<string>(cells);
            _textEditor.TextArea.TextView.Redraw();

            _currentModel = modelName;
            _currentCell = cellName;
        }

        public void UpdateContextVariables(IEnumerable<string> variables)
        {
            _contextVariables = new HashSet<string>(variables);
            _textEditor.TextArea.TextView.Redraw();
        }

        public void UpdateAssumptions(IEnumerable<string> variables)
        {
            _assumptions = new HashSet<string>(variables);
            _textEditor.TextArea.TextView.Redraw();
        }

        //Setting
        private void SetTextEditorStyle()
        {
            _textEditor.FontFamily = new FontFamily("Consolas");
            _textEditor.FontSize = 12;
            _textEditor.ShowLineNumbers = true;
            _textEditor.Options.EnableHyperlinks = true;
            _textEditor.Options.EnableEmailHyperlinks = true;

            // 배경색과 전경색을 테마에 따라 동적으로 설정
            _textEditor.SetResourceReference(TextEditor.BackgroundProperty, "SystemControlBackgroundAltHighBrush");
            _textEditor.SetResourceReference(TextEditor.ForegroundProperty, "SystemControlForegroundBaseHighBrush");

            // 선택 영역 스타일 설정
            _textEditor.TextArea.SelectionBrush = new SolidColorBrush(Color.FromArgb(60, 51, 153, 255));
            _textEditor.TextArea.SelectionForeground = null;  // 선택된 텍스트의 전경색은 변경하지 않음

            // 현재 줄 강조 설정
            _textEditor.TextArea.TextView.CurrentLineBackground = new SolidColorBrush(Color.FromArgb(40, 0, 0, 0));
            _textEditor.TextArea.TextView.CurrentLineBorder = new Pen(new SolidColorBrush(Color.FromArgb(40, 0, 0, 0)), 2);
        }

        private void LoadAutoCompletionSet()
        {
            string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoCompletion.json");
            if (File.Exists(jsonPath))
            {
                string json = File.ReadAllText(jsonPath);
                var autoCompletionSet = JsonConvert.DeserializeObject<AutoCompletionSet>(json);
                _functions = autoCompletionSet.Functions;
                _cellCompletions = autoCompletionSet.CellCompletions;
            }
        }
        
        private void Current_ActualApplicationThemeChanged(ThemeManager sender, object args)
        {
            _textEditor.TextArea.TextView.Redraw();
        }
    }

    public static class RegexPatterns
    {
        // Matches comments starting with '//' until the end of the line
        // Example: "// This is a comment"
        public static readonly Regex Comment = new Regex(@"//.*");

        // Matches the separator '--' used to separate cell names from formulas
        // Example: "cellName -- formula"
        public static readonly Regex Separator = new Regex(@"--");

        // Matches whole numbers or decimal numbers
        // Examples: "42", "3.14"
        public static readonly Regex Number = new Regex(@"\b\d+(\.\d+)?\b");

        // Matches strings enclosed in double quotes, allowing for escaped quotes
        // Example: "Hello, \"world\"!"
        public static readonly Regex String = new Regex(@"""(?:\\.|[^""\\])*""");

        // Matches boolean values (true or false)
        // Examples: "true", "false"
        public static readonly Regex Boolean = new Regex(@"\b(true|false)\b");

        // Matches cell references, optionally with model name and parameters
        // Examples: "cellName[", "modelName.cellName[", "modelName{param1:value,param2:value}.cellName["
        public static readonly Regex Cell = new Regex(@"\b\w+(?:{[^{}]*})?(?:\.\w+)?\[", RegexOptions.Compiled);

        // Matches cell definitions, including optional description comments
        // Example: "// CellReference description\ncellName -- formula"
        public static readonly Regex CellDefinition = new Regex(@"(?://(?<description>.*)\r?\n)?(?<cellName>\w+)\s*--\s*(?<formula>.+)(\r?\n|$)", RegexOptions.Compiled);

        // Matches Invoke function calls
        // Example: "Invoke(cellName, 0)"
        public static readonly Regex Invoke = new Regex(@"Invoke\((\w+),\s*(\d+)\)");

        // Matches the inside of an Assum function call, used for auto-completion
        // Example: "Assum("assumption
        public static readonly Regex AssumInside = new Regex(@"Assum\([""']([^""']*)$");

        // Matches the inside of an Exp function call, used for auto-completion
        // Example: "Exp("expense
        public static readonly Regex ExpInside = new Regex(@"Exp\([""']([^""']*)$");

        // Matches model and cell references for auto-completion
        // Example: "modelName.cellName"
        public static readonly Regex ModelDotCell = new Regex(@"(\w+)({[^}]*})?\.([\w]*)$");

        // Matches model names with optional parameters
        // Examples: "modelName", "modelName{param:value}"
        public static readonly Regex ModelAndParams = new Regex(@"(\w+({[^}]*})?\.[\w.]*)$");

        // Generates a regex for matching function names based on a provided list
        // Example: If functions = ["Sum", "Prd"], it generates: \b(Sum|Prd)\b
        public static Regex GetFunctionRegex(IEnumerable<string> functions)
        {
            string pattern = $@"\b({string.Join("|", functions)})\b";
            return new Regex(pattern);
        }

        // Generates a regex for matching model names based on a provided list
        // Example: If models = ["Model1", "Model2"], it generates: \b(Model1|Model2)\b
        public static Regex GetModelRegex(IEnumerable<string> models)
        {
            string pattern = $@"\b({string.Join("|", models)})\b";
            return new Regex(pattern);
        }

        // Generates a regex for matching context variable names based on a provided list
        // Example: If contextVariables = ["var1", "var2"], it generates: \b(var1|var2)\b
        public static Regex GetContextVariableRegex(IEnumerable<string> contextVariables)
        {
            string pattern = $@"\b({string.Join("|", contextVariables)})\b";
            return new Regex(pattern);
        }
    }

    public class AutoCompletionSet
    {
        public List<string> Functions { get; set; }
        public Dictionary<string, string> CellCompletions { get; set; }
    }

    public enum CompletionType
    {
        None,
        CellReference,
        Function,
        Model,
        ContextVariable,
        Assumption,
        Expense,
        CellAutoCompletion
    }

    public class CustomCompletionData : ICompletionData
    {
        public CustomCompletionData(string text, CompletionType type, string highlightText)
        {
            Text = text;
            CompletionType = type;
            HighlightText = highlightText;
        }

        public ImageSource Image => null;
        public string Text { get; }
        public string HighlightText { get; }
        public CompletionType CompletionType { get; }
        public object Content => CreateTextBlock();
        public object Description => null;
        public double Priority => 0;

        private TextBlock CreateTextBlock()
        {
            var theme = ThemeManager.Current.ActualApplicationTheme;
            var colorKey = CompletionType switch
            {
                CompletionType.CellReference => "Cell",
                CompletionType.Function => "Function",
                CompletionType.Model => "Model",
                CompletionType.ContextVariable => "ContextVariable",
                CompletionType.Assumption => "ContextVariable",
                CompletionType.CellAutoCompletion => "Cell",
                CompletionType.Expense => "ContextVariable",
                _ => "Default"
            };

            var color = GetThemeColor(colorKey, theme);

            var textBlock = new TextBlock
            {
                FontSize = 11
            };

            // 입력된 부분을 찾아 볼드 처리
            string inputText = HighlightText;
            if (!string.IsNullOrEmpty(inputText))
            {
                int index = Text.IndexOf(inputText, StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    if (index > 0)
                        textBlock.Inlines.Add(new Run(Text.Substring(0, index)) { Foreground = new SolidColorBrush(color) });

                    textBlock.Inlines.Add(new Run(Text.Substring(index, inputText.Length))
                    {
                        Foreground = new SolidColorBrush(color),
                        FontWeight = FontWeights.Bold,
                        FontSize = 12, // 폰트 크기를 키움
                        FontFamily = new FontFamily("Segoe UI") // 폰트 변경
                    });

                    if (index + inputText.Length < Text.Length)
                        textBlock.Inlines.Add(new Run(Text.Substring(index + inputText.Length)) { Foreground = new SolidColorBrush(color) });
                }
                else
                {
                    textBlock.Inlines.Add(new Run(Text) { Foreground = new SolidColorBrush(color) });
                }
            }
            else
            {
                textBlock.Inlines.Add(new Run(Text) { Foreground = new SolidColorBrush(color) });
            }

            return textBlock;
        }

        private Color GetThemeColor(string colorKey, ApplicationTheme theme)
        {
            if (SyntaxHighlighter.ColorSchemes.TryGetValue(colorKey, out var colors))
            {
                return theme switch
                {
                    ApplicationTheme.Dark => colors.Dark,
                    ApplicationTheme.Light => colors.Light,
                    _ => colors.HighContrast
                };
            }
            return Colors.Black; // 기본 색상
        }

        public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs)
        {
            try
            {
                string lineText = GetCurrentLineText(textArea, completionSegment);
                string textBeforeCaret = lineText.Substring(0, completionSegment.EndOffset - textArea.Document.GetLineByOffset(completionSegment.Offset).Offset);
                string textAfterCaret = lineText.Substring(completionSegment.EndOffset - textArea.Document.GetLineByOffset(completionSegment.Offset).Offset);

                string insertText = Text;

                switch (CompletionType)
                {
                    case CompletionType.CellReference:
                        insertText = CompleteCellReference(textBeforeCaret, textAfterCaret);
                        break;
                    case CompletionType.Function:
                        insertText = CompleteFunctionCall();
                        break;
                    case CompletionType.Model:
                        insertText = CompleteModelReference(textBeforeCaret, textAfterCaret);
                        break;
                    case CompletionType.Assumption:
                        insertText = CompleteAssumption(Text);
                        break;
                    case CompletionType.CellAutoCompletion:
                        insertText = CompleteCellDefinition(lineText);
                        break;
                }

                ReplaceText(textArea, completionSegment, insertText);

                // 커서 위치 조정
                if (CompletionType == CompletionType.Function)
                {
                    if (Text == "Assum")
                    {
                        textArea.Caret.Offset -= 5;
                    }
                    else if(Text == "Exp")
                    {
                        textArea.Caret.Offset -= 2;
                    }
                    else
                    {
                        textArea.Caret.Offset -= 1;
                    }
                }

                if (CompletionType == CompletionType.ContextVariable || CompletionType == CompletionType.Assumption)
                {
                    //캐럿 초기화로 자동완성창 다시실행 차단
                    textArea.Caret.Offset -= lineText.Length-1;
                    textArea.Caret.Offset += lineText.Length-1;
                }
            }
            catch
            {

            }

        }

        private string GetCurrentLineText(TextArea textArea, ISegment completionSegment)
        {
            var currentLine = textArea.Document.GetLineByOffset(completionSegment.Offset);
            return textArea.Document.GetText(currentLine.Offset, currentLine.Length);
        }

        private string CompleteCellReference(string textBeforeCaret, string textAfterCaret)
        {
            // 전체 입력 텍스트 구성
            string fullText = textBeforeCaret + textAfterCaret;

            // 모델 이름 (있는 경우)과 현재 입력된 셀 이름 분리
            var modelMatch = Regex.Match(textBeforeCaret, @"(\w+({[^}]*})?\.)?(\w*)$");
            string modelPrefix = modelMatch.Groups[1].Value;
            string currentInput = modelMatch.Groups[3].Value;

            // 인덱스 확인 (대괄호 안의 내용)
            var indexMatch = Regex.Match(fullText, @"\[([^\]]*)\]");
            string existingIndex = indexMatch.Success ? $"[{indexMatch.Groups[1].Value}]" : "";

            // {param:value} 패턴이 있는지 확인
            bool hasParameters = modelPrefix.Contains("{") && modelPrefix.Contains("}");

            string result = hasParameters ? $".{Text}" : $"{modelPrefix}{Text}";

            // 캐럿이 '[' 바로 앞에 있는 경우
            if (!textAfterCaret.StartsWith("["))
            {
                result += "[t]";
            }

            return result;
        }

        private string CompleteFunctionCall()
        {
            if (Text == "Assum")
            {
                return $"{Text}(\"\")[t]";
            }
            if (Text == "Exp")
            {
                return $"{Text}(\"\")";
            }
            else
            {
                return $"{Text}()";
            }
        }

        private string CompleteAssumption(string text)
        {
            return text.Replace("|", @""", """);
        }

        private string CompleteModelReference(string textBeforeCaret, string textAfterCaret)
        {
            if (textAfterCaret.StartsWith("."))
            {
                return Text;
            }
            else
            {
                return $"{Text}.";
            }
        }

        private string CompleteCellDefinition(string lineText)
        {
            var parts = lineText.Split(new[] { "--" }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0)
            {
                var cellName = parts[0].Trim();
                return $"{cellName} -- {Text}";
            }
            else
            {
                return Text;
            }
        }

        private void ReplaceText(TextArea textArea, ISegment completionSegment, string insertText)
        {
            if (CompletionType == CompletionType.CellAutoCompletion)
            {
                var currentLine = textArea.Document.GetLineByOffset(completionSegment.Offset);
                textArea.Document.Replace(currentLine.Offset, currentLine.Length, insertText);
            }
            else
            {
                textArea.Document.Replace(completionSegment, insertText);
            }
        }
    }

    public class CustomCompletionWindow : CompletionWindow
    {
        private bool _isUpdating = false;
        private Border _mainBorder;

        public CustomCompletionWindow(TextArea textArea) : base(textArea)
        {
            this.ResizeMode = ResizeMode.NoResize;
            this.CloseWhenCaretAtBeginning = true;
            this.CloseAutomatically = true;
            this.Width = 240;
            this.MaxHeight = 120;

            // CompletionList의 필터링 비활성화
            this.CompletionList.IsFiltering = false;
            // Window 배경을 투명하게 설정
            this.AllowsTransparency = true;
            this.Background = Brushes.Transparent;

            // 메인 Border 생성 및 설정
            _mainBorder = new Border
            {
                CornerRadius = new CornerRadius(3),
                ClipToBounds = true
            };

            // CompletionList를 _mainBorder의 자식으로 설정
            if (this.Content is UIElement content)
            {
                this.Content = null;
                _mainBorder.Child = content;
                this.Content = _mainBorder;
            }

            // CompletionList의 필터링 비활성화
            this.CompletionList.IsFiltering = false;


            // 키 입력 이벤트 핸들러 추가
            this.TextArea.TextEntered += TextArea_TextEntered;
            this.Closed += CustomCompletionWindow_Closed;

            // 테마 변경 이벤트 구독
            ThemeManager.Current.ActualApplicationThemeChanged += Current_ActualApplicationThemeChanged;

            // 초기 테마 적용
            ApplyTheme();
        }

        public new void Show()
        {
            this.Opacity = 0;
            base.Show();
            UpdateListBox();
            StartAnimation(0, 1);
        }

        public new void Close()
        {
            StartAnimation(1, 0, () => base.Close());
        }

        private void StartAnimation(double from, double to, Action completedAction = null)
        {
            var animation = new DoubleAnimation
            {
                From = from,
                To = to,
                Duration = TimeSpan.FromMilliseconds(200)
            };

            if (completedAction != null)
            {
                animation.Completed += (s, e) => completedAction();
            }

            this.BeginAnimation(UIElement.OpacityProperty, animation);
        }

        private void TextArea_TextEntered(object sender, TextCompositionEventArgs e)
        {
            if (!_isUpdating)
            {
                _isUpdating = true;
                UpdateListBox();
                _isUpdating = false;
            }
        }

        private void UpdateListBox()
        {
            var listBox = this.CompletionList.ListBox;
            var allItems = this.CompletionList.CompletionData.ToList();

            if (listBox.Items.Count > 0)
            {
                listBox.SelectedIndex = 0;
            }

            // Force layout update
            listBox.UpdateLayout();
            this.UpdateLayout();
        }

        private void CustomCompletionWindow_Closed(object sender, EventArgs e)
        {
            this.TextArea.TextEntered -= TextArea_TextEntered;
            ThemeManager.Current.ActualApplicationThemeChanged -= Current_ActualApplicationThemeChanged;
        }

        private void ApplyTheme()
        {
            var theme = ThemeManager.Current.ActualApplicationTheme;

            if (theme == ApplicationTheme.Dark)
            {
                this.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                this.Foreground = new SolidColorBrush(Colors.White);
                this.BorderBrush = new SolidColorBrush(Color.FromRgb(60, 60, 60));
            }
            else // Light theme
            {
                this.Background = new SolidColorBrush(Colors.White);
                this.Foreground = new SolidColorBrush(Colors.Black);
                this.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200));
            }

            // CompletionList 스타일 업데이트
            if (this.CompletionList.ListBox != null)
            {
                this.CompletionList.ListBox.Background = this.Background;
                this.CompletionList.ListBox.Foreground = this.Foreground;
                this.CompletionList.ListBox.BorderBrush = this.BorderBrush;
            }
        }

        private void Current_ActualApplicationThemeChanged(ThemeManager sender, object args)
        {
            ApplyTheme();
        }
    }

}
