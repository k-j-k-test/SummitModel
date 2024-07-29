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
using System.Text.Json;
using ActuLight.Properties;
using ModernWpf;
using System.Windows.Data;

namespace ActuLight
{
    public class SyntaxHighlighter : DocumentColorizingTransformer
    {
        public static Dictionary<string, (Color Dark, Color Light, Color HighContrast)> ColorSchemes;

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
            _textEditor.TextArea.TextEntered += TextArea_TextEntered;

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
        private async void TextArea_TextEntered(object sender, TextCompositionEventArgs e)
        {
            await Task.Delay(200);

            if (!(sender is TextArea textArea)) return;

            string currentLine = GetCurrentLine(textArea);
            string currentWord = GetCurrentWord(textArea);

            var completionData = new List<ICompletionData>();

            if (currentLine.Contains("--"))
            {
                var parts = currentLine.Split(new[] { "--" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2 && parts[1].Trim() == "")
                {
                    if (_cellCompletions.TryGetValue(parts[0].Trim(), out string completion))
                    {
                        completionData = new List<ICompletionData>
                        {
                            new CustomCompletionData(completion, CompletionType.CellCompletion)
                        };
                        ShowCompletionWindow(textArea, completionData, textArea.Caret.Offset);
                    }
                    return;
                }
            }

            if (string.IsNullOrEmpty(currentWord) || !currentLine.Contains("--") || currentLine.IndexOf("--") >= textArea.Caret.Column)
            {
                _completionWindow?.Close();
                return;
            }

            // Check if we're inside Assum function
            var textBeforeCaret = textArea.Document.GetText(0, textArea.Caret.Offset);
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

            if (completionData.Any())
            {
                ShowCompletionWindow(textArea, completionData);
            }
            else
            {
                _completionWindow?.Close();
            }
        }

        private void ShowCompletionWindow(TextArea textArea, IEnumerable<ICompletionData> completionData, int? startOffset = null)
        {
            _completionWindow?.Close();

            _completionWindow = new CompletionWindow(textArea)
            {
                ResizeMode = ResizeMode.NoResize,
                CloseWhenCaretAtBeginning = true,
                CloseAutomatically = true,
            };

            if (startOffset.HasValue)
            {
                _completionWindow.StartOffset = startOffset.Value;
            }
            else
            {
                var currentWord = GetCurrentWord(textArea);
                _completionWindow.StartOffset = textArea.Caret.Offset - currentWord.Length;
            }

            var dataList = _completionWindow.CompletionList.CompletionData;
            foreach (var data in completionData)
            {
                dataList.Add(data);
            }

            _completionWindow.Show();

            if (dataList.Count > 0)
            {
                _completionWindow.CompletionList.SelectedItem = dataList[0];
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
            var theme = ThemeManager.Current.ActualApplicationTheme;

            ApplyColorRule(line, text, @"//.*", "Comment");
            ApplyColorRule(line, text, @"--", "Separator");
            ApplyColorRule(line, text, @"\b\d+(\.\d+)?\b", "Number");
            ApplyColorRule(line, text, @"""(?:\\.|[^""\\])*""", "String");
            ApplyColorRule(line, text, @"\b(true|false)\b", "Boolean");
            ApplyColorRule(line, text, $@"\b({string.Join("|", _functions)})\b", "Function");
            ApplyColorRule(line, text, $@"\b({string.Join("|", _models)})\b", "Model");
            ApplyColorRule(line, text, $@"\b({string.Join("|", _contextVariables)})\b", "ContextVariable");

            ApplyCurrentModelCellColorRule(line, text);
            ApplyOtherModelCellColorRule(line, text);
            ApplyCellReferenceValidation(line, text);

        }

        private void ApplyColorRule(DocumentLine line, string text, string pattern, string colorKey)
        {
            var regex = new Regex(pattern);
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

            // 모든 셀 참조 패턴 (현재 모델 및 다른 모델, 파라미터 포함)
            string pattern = @"(\w+)({[^}]+})?\.?(\w+)?\[";
            var regex = new Regex(pattern);

            foreach (Match match in regex.Matches(text))
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
                var autoCompletionSet = JsonSerializer.Deserialize<AutoCompletionSet>(json);
                _functions = autoCompletionSet.Functions;
                _cellCompletions = autoCompletionSet.CellCompletions;
            }
        }

        private void Current_ActualApplicationThemeChanged(ThemeManager sender, object args)
        {
            _textEditor.TextArea.TextView.Redraw();
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
                CompletionType.Assumption => "ContextVariable", // 가정도 ContextVariable과 같은 색상 사용
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
            string insertText;

            // Get the text from the start of the line to the completion segment
            var lineStart = textArea.Document.GetLineByOffset(completionSegment.Offset).Offset;
            var textBeforeSegment = textArea.Document.GetText(completionSegment.Offset, completionSegment.Length);

            if (CompletionType == CompletionType.Cell)
            {
                // Check for model name and parameters
                var modelMatch = Regex.Match(textBeforeSegment, @"(\w+({[^}]*})?\.[\w.]*)$");
                if (modelMatch.Success)
                {
                    string modelPrefix = modelMatch.Groups[1].Value;
                    // Remove the last dot if it exists
                    if (modelPrefix.EndsWith("."))
                    {
                        modelPrefix = modelPrefix.Substring(0, modelPrefix.Length - 1);
                    }
                    insertText = $"{modelPrefix}.{Text}[t]";
                }
                else
                {
                    if (textBeforeSegment.Contains("."))
                    {
                        insertText = $".{Text}[t]";
                    }
                    else
                    {
                        insertText = $"{Text}[t]";
                    }                 
                }
            }
            else if (CompletionType == CompletionType.Function)
            {
                if (Text == "Assum")
                {
                    insertText = $"{Text}(\"\")[t]";
                }
                else
                {
                    insertText = $"{Text}()";
                }
            }
            else if (CompletionType == CompletionType.Model)
            {
                insertText = $"{Text}.";
            }
            else if (CompletionType == CompletionType.Assumption)
            {
                insertText = Text;
            }
            else if (CompletionType == CompletionType.CellCompletion)
            {
                var currentLine = textArea.Document.GetLineByOffset(completionSegment.Offset);
                var lineText = textArea.Document.GetText(currentLine.Offset, currentLine.Length);
                var parts = lineText.Split(new[] { "--" }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length > 0)
                {
                    var cellName = parts[0].Trim();
                    insertText = $"{cellName} -- {Text}";
                    textArea.Document.Replace(currentLine.Offset, currentLine.Length, insertText);
                    return;
                }
                else
                {
                    insertText = Text;
                }
            }
            else
            {
                insertText = Text;
            }

            textArea.Document.Replace(completionSegment, insertText);

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
    }

}