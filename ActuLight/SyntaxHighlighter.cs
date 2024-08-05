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
using System.Security.RightsManagement;

namespace ActuLight
{
    public class SyntaxHighlighter : DocumentColorizingTransformer
    {
        public static Dictionary<string, (Color Dark, Color Light, Color HighContrast)> ColorSchemes;
        public static int Delay = 300;

        private string _currentModel;
        private string _currentCell;

        private CompletionWindow _completionWindow;
        private readonly TextEditor _textEditor;

        private Dictionary<string, HashSet<string>> _modelCells = new Dictionary<string, HashSet<string>>();
        private HashSet<string> _models = new HashSet<string>();
        private HashSet<string> _contextVariables = new HashSet<string>();
        private HashSet<string> _assumptions = new HashSet<string>();

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
        public async void TextArea_TextEntered(object sender, EventArgs  e)
        {
            await Debouncer.Debounce("TextAreaUpdate", 10, async () =>
            {
                if (!(sender is TextEditor textEditor)) return;

                var textArea = textEditor.TextArea;

                // UI 스레드에서 필요한 정보를 가져옵니다.
                var textBeforeCaret = await Application.Current.Dispatcher.InvokeAsync(() =>
                    textArea.Document.GetText(0, textArea.Caret.Offset));
                var currentLine = await Application.Current.Dispatcher.InvokeAsync(() =>
                    GetCurrentLine(textArea));
                var currentWord = await Application.Current.Dispatcher.InvokeAsync(() =>
                    GetCurrentWord(textArea));

                var completionData = await Task.Run(() => GetCompletionData(textBeforeCaret, currentLine, currentWord));

                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    if (completionData.Any())
                    {
                        ShowCompletionWindow(textArea, completionData);
                    }
                    else
                    {
                        _completionWindow?.Close();
                        _completionWindow = null;
                    }
                });
            });
        }

        private List<ICompletionData> GetCompletionData(string textBeforeCaret, string currentLine, string currentWord)
        {
            var completionData = new List<ICompletionData>();

            if (currentLine.Contains("--"))
            {
                var parts = currentLine.Split(new[] { "--" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2 && parts[1].Trim() == "")
                {
                    if (_cellCompletions.TryGetValue(parts[0].Trim(), out string completion))
                    {
                        completionData.Add(new CustomCompletionData(completion, CompletionType.CellCompletion));
                        return completionData;
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
                completionData.AddRange(_assumptions
                    .Where(assum => assum.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase))
                    .Select(assumption => new CustomCompletionData(assumption, CompletionType.Assumption)));
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
                        completionData.AddRange(cells
                            .Where(cell => cell.StartsWith(cellPrefix, StringComparison.OrdinalIgnoreCase))
                            .Select(cell => new CustomCompletionData(cell, CompletionType.Cell)));
                    }
                }
                else
                {
                    // Add cell completions
                    if (!string.IsNullOrEmpty(_currentModel) && _modelCells.TryGetValue(_currentModel, out var currentModelCells))
                    {
                        completionData.AddRange(currentModelCells
                            .Where(cell => cell.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase))
                            .Select(cell => new CustomCompletionData(cell, CompletionType.Cell)));
                    }

                    // Add function completions
                    completionData.AddRange(_functions
                        .Where(func => func.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase))
                        .Select(func => new CustomCompletionData(func, CompletionType.Function)));

                    // Add model completions
                    completionData.AddRange(_models
                        .Where(model => model.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase))
                        .Select(model => new CustomCompletionData(model, CompletionType.Model)));

                    // Add context variable completions
                    completionData.AddRange(_contextVariables
                        .Where(var => var.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase))
                        .Select(var => new CustomCompletionData(var, CompletionType.ContextVariable)));
                }
            }

            return completionData;
        }

        private void ShowCompletionWindow(TextArea textArea, IEnumerable<ICompletionData> completionData)
        {
            if (_completionWindow != null)
            {
                _completionWindow.Close();
            }

            _completionWindow = new CompletionWindow(textArea)
            {
                ResizeMode = ResizeMode.NoResize,
                CloseWhenCaretAtBeginning = true,
                CloseAutomatically = true,
                Opacity = 0, // 시작 시 투명하게 설정
                Width = 240, // 너비 고정
                MaxHeight = 120 // 최대 높이 설정
            };
            
            var currentWord = GetCurrentWord(textArea);
            _completionWindow.StartOffset = textArea.Caret.Offset - currentWord.Length;

            var dataList = _completionWindow.CompletionList.CompletionData;
            foreach (var data in completionData)
            {
                dataList.Add(data);
            }

            if (dataList.Count > 0)
            {
                _completionWindow.Show();
                _completionWindow.CompletionList.SelectedItem = dataList[0];
                _completionWindow.Closed += (sender, args) => _completionWindow = null;

                // 애니메이션 효과 추가
                var animation = new DoubleAnimation
                {
                    From = 0,
                    To = 1,
                    Duration = TimeSpan.FromMilliseconds(0)
                };
                _completionWindow.BeginAnimation(UIElement.OpacityProperty, animation);
            }
            else
            {
                _completionWindow.Close();
                _completionWindow = null;
            }
        }

        private static string GetCurrentLine(TextArea textArea)
        {
            var caretPosition = textArea.Caret.Offset;
            var currentLine = textArea.Document.GetLineByOffset(caretPosition);
            return textArea.Document.GetText(currentLine.Offset, currentLine.Length);
        }

        private static string GetCurrentWord(TextArea textArea)
        {
            var caretPosition = textArea.Caret.Offset;
            var lineStart = textArea.Document.GetLineByOffset(caretPosition).Offset;
            var text = textArea.Document.GetText(lineStart, caretPosition - lineStart);
            var match = Regex.Match(text, @"[\w.]+$");
            return match.Success ? match.Value : string.Empty;
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
                string modelName = match.Groups[1].Value;
                string parameters = match.Groups[2].Value;
                string cellName = match.Groups[3].Value;

                int modelStartIndex = match.Index;
                int modelLength = modelName.Length + parameters.Length;

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
        public static readonly Regex Cell = new Regex(@"(\w+)({[^}]*})?\.?([\w.]+)?\[");

        // Matches cell definitions, including optional description comments
        // Example: "// Cell description\ncellName -- formula"
        public static readonly Regex CellDefinition = new Regex(@"(?://(?<description>.*)\r?\n)?(?<cellName>\w+)\s*--\s*(?<formula>.+)(\r?\n|$)");

        // Matches Invoke function calls
        // Example: "Invoke(cellName, 0)"
        public static readonly Regex Invoke = new Regex(@"Invoke\((\w+),\s*(\d+)\)");

        // Matches the inside of an Assum function call, used for auto-completion
        // Example: "Assum("assumption
        public static readonly Regex AssumInside = new Regex(@"Assum\([""']([^""']*)$");

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
        Cell,
        Function,
        Model,
        ContextVariable,
        Assumption,
        CellCompletion
    }

    public class CustomCompletionData : ICompletionData
    {
        public CustomCompletionData(string text, CompletionType type)
        {
            Text = text;
            CompletionType = type;
        }

        public ImageSource Image => null;
        public string Text { get; }
        public CompletionType CompletionType { get; }
        public object Content => CreateTextBlock();
        public object Description => null;
        public double Priority => 0;

        private TextBlock CreateTextBlock()
        {
            var theme = ThemeManager.Current.ActualApplicationTheme;
            var colorKey = CompletionType switch
            {
                CompletionType.Cell => "Cell",
                CompletionType.Function => "Function",
                CompletionType.Model => "Model",
                CompletionType.ContextVariable => "ContextVariable",
                CompletionType.Assumption => "ContextVariable",
                CompletionType.CellCompletion => "Cell",
                _ => "Default"
            };

            var color = GetThemeColor(colorKey, theme);

            return new TextBlock
            {
                Text = Text,
                Foreground = new SolidColorBrush(color),
                FontSize = 11
            };
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
            string lineText = GetCurrentLineText(textArea, completionSegment);
            string textBeforeCaret = lineText.Substring(0, completionSegment.EndOffset - textArea.Document.GetLineByOffset(completionSegment.Offset).Offset);
            string textAfterCaret = lineText.Substring(completionSegment.EndOffset - textArea.Document.GetLineByOffset(completionSegment.Offset).Offset);

            string insertText = Text;

            switch (CompletionType)
            {
                case CompletionType.Cell:
                    insertText = CompleteCellReference(textBeforeCaret, textAfterCaret);
                    break;
                case CompletionType.Function:
                    insertText = CompleteFunctionCall();
                    break;
                case CompletionType.Model:
                    insertText = CompleteModelReference(textBeforeCaret, textAfterCaret);
                    break;
                case CompletionType.Assumption:
                    insertText = Text;
                    break;
                case CompletionType.CellCompletion:
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
                else
                {
                    textArea.Caret.Offset -= 1;
                }
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

            string result;

            // 캐럿이 '[' 바로 앞에 있는 경우
            if (textAfterCaret.StartsWith("["))
            {
                // 모델 접두사가 있으면 유지하고, 없으면 그대로 진행
                result = string.IsNullOrEmpty(modelPrefix) ? Text : $"{modelPrefix}{Text}";
            }
            // 캐럿이 '[' 앞이 아닌 경우
            else
            {
                // 현재 입력이 비어있거나 완전한 셀 이름이 아닌 경우
                if (string.IsNullOrEmpty(currentInput) || currentInput != Text)
                {
                    result = $"{modelPrefix}{Text}";
                }
                // 현재 입력이 이미 완전한 셀 이름인 경우
                else
                {
                    result = $"{modelPrefix}{Text}";
                }

                // 인덱스가 없는 경우에만 [t] 추가
                if (string.IsNullOrEmpty(existingIndex))
                {
                    result += "[t]";
                }
            }

            return result;
        }

        private string CompleteFunctionCall()
        {
            if (Text == "Assum")
            {
                return $"{Text}(\"\")[t]";
            }
            else
            {
                return $"{Text}()";
            }
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
            if (CompletionType == CompletionType.CellCompletion)
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
}
