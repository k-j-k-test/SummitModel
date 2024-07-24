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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Windows.Documents;

namespace ActuLight
{
    public class SyntaxHighlighter : DocumentColorizingTransformer
    {
        private List<(Regex Regex, SolidColorBrush Brush)> _rules;
        private CompletionWindow _completionWindow;
        private readonly TextEditor _textEditor;

        private Dictionary<string, HashSet<string>> _modelCells = new Dictionary<string, HashSet<string>>();
        private HashSet<string> _models = new HashSet<string>();
        private HashSet<string> _contextVariables = new HashSet<string>();
        private HashSet<string> _assumptions = new HashSet<string>();

        // Functions는 정적으로 유지
        private static readonly string[] Functions = { "Invoke", "Sum", "Prd", "Assum", "If" };

        public SyntaxHighlighter(TextEditor textEditor)
        {
            _textEditor = textEditor;
            InitializeRules();

            _textEditor.TextArea.TextEntered += TextArea_TextEntered;
        }

        private void InitializeRules()
        {
            _rules = new List<(Regex Regex, SolidColorBrush Brush)>
            {
                (new Regex(@"//.*"), new SolidColorBrush(Colors.LightGreen)),           // 주석
                (new Regex(@"--"), new SolidColorBrush(Colors.Gray)),                   // 구분자
                (new Regex($@"\b({string.Join("|", Functions)})\b"), new SolidColorBrush(Colors.LightSkyBlue)),  // 함수
                (new Regex(@"\b\d+(\.\d+)?\b"), new SolidColorBrush(Colors.Gold)),      // 숫자
                (new Regex(@"""(?:\\.|[^""\\])*"""), new SolidColorBrush(Colors.LightGoldenrodYellow)),  // 문자열 (큰따옴표 안의 텍스트)
                (new Regex(@"\b(true|false)\b"), new SolidColorBrush(Colors.Orange))    // boolean 값
            };
        }

        protected override void ColorizeLine(DocumentLine line)
        {
            string text = CurrentContext.Document.GetText(line);

            // 기존 규칙 적용
            foreach (var (regex, brush) in _rules)
            {
                foreach (Match match in regex.Matches(text))
                {
                    ChangeLinePart(
                        line.Offset + match.Index,
                        line.Offset + match.Index + match.Length,
                        element => element.TextRunProperties.SetForegroundBrush(brush)
                    );
                }
            }

            // 모델 하이라이팅
            var modelRegex = new Regex($@"\b({string.Join("|", _models)})\b");
            foreach (Match match in modelRegex.Matches(text))
            {
                ChangeLinePart(
                    line.Offset + match.Index,
                    line.Offset + match.Index + match.Length,
                    element => element.TextRunProperties.SetForegroundBrush(new SolidColorBrush(Colors.HotPink))
                );
            }

            // Cell에 대한 동적 규칙 적용
            var cellRegex = new Regex($@"\b({string.Join("|", _modelCells.Values.SelectMany(x => x))})\b");
            foreach (Match match in cellRegex.Matches(text))
            {
                ChangeLinePart(
                    line.Offset + match.Index,
                    line.Offset + match.Index + match.Length,
                    element => element.TextRunProperties.SetForegroundBrush(new SolidColorBrush(Colors.Orange))
                );
            }

            // Context 변수 하이라이팅
            var contextVarRegex = new Regex($@"\b({string.Join("|", _contextVariables)})\b");
            foreach (Match match in contextVarRegex.Matches(text))
            {
                ChangeLinePart(
                    line.Offset + match.Index,
                    line.Offset + match.Index + match.Length,
                    element => element.TextRunProperties.SetForegroundBrush(new SolidColorBrush(Colors.LightBlue))
                );
            }

            // 가정 하이라이팅
        }

        private async void TextArea_TextEntered(object sender, TextCompositionEventArgs e)
        {
            await Task.Delay(200);

            if (!(sender is TextArea textArea)) return;

            string currentWord = GetCurrentWord(textArea);
            if (string.IsNullOrEmpty(currentWord))
            {
                _completionWindow?.Close();
                return;
            }

            var completionData = new List<ICompletionData>();

            // Check if we're inside Assum function
            var textBeforeCaret = textArea.Document.GetText(0, textArea.Caret.Offset);
            var assumMatch = Regex.Match(textBeforeCaret, @"Assum\([""']([^""']*)$");
            if (assumMatch.Success)
            {
                // We're inside Assum function, offer assumption completions
                completionData.AddRange(_assumptions
                    .Where(assumption => MatchesKoreanInitial(assumption, currentWord))
                    .Select(assumption => new CustomCompletionData(assumption, CompletionType.Assumption)));
            }
            else
            {
                // Check if we're after a model name and a dot
                var modelMatch = Regex.Match(textBeforeCaret, @"(\w+)\.\s*$");
                if (modelMatch.Success)
                {
                    string modelName = modelMatch.Groups[1].Value;
                    if (_modelCells.TryGetValue(modelName, out var cells))
                    {
                        completionData.AddRange(cells
                            .Select(cell => new CustomCompletionData(cell, CompletionType.Cell)));
                    }
                }
                else
                {
                    // Add cell completions
                    completionData.AddRange(_modelCells.Values.SelectMany(cells => cells)
                        .Where(cell => cell.StartsWith(currentWord, StringComparison.OrdinalIgnoreCase))
                        .Select(cell => new CustomCompletionData(cell, CompletionType.Cell)));

                    // Add function completions
                    completionData.AddRange(Functions
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

        private void ShowCompletionWindow(TextArea textArea, List<ICompletionData> completionData)
        {
            _completionWindow?.Close();

            _completionWindow = new CompletionWindow(textArea)
            {
                ResizeMode = ResizeMode.NoResize,
                OverridesDefaultStyle = false,
                CloseWhenCaretAtBeginning = true,
                CloseAutomatically = true,
            };

            var currentWord = GetCurrentWord(textArea);
            _completionWindow.StartOffset = textArea.Caret.Offset - currentWord.Length;

            var dataList = _completionWindow.CompletionList.CompletionData;
            foreach (var data in completionData)
            {
                dataList.Add(data);
            }

            _completionWindow.Show();
        }

        private static string GetCurrentWord(TextArea textArea)
        {
            var caretPosition = textArea.Caret.Offset;
            var lineStart = textArea.Document.GetLineByOffset(caretPosition).Offset;
            var text = textArea.Document.GetText(lineStart, caretPosition - lineStart);
            var match = Regex.Match(text, @"[\w.]+$");
            return match.Success ? match.Value : string.Empty;
        }

        public void UpdateModels(IEnumerable<string> models)
        {
            _models = new HashSet<string>(models);
            _textEditor.TextArea.TextView.Redraw();
        }

        public void UpdateModelCells(string modelName, IEnumerable<string> cells)
        {
            if (!_modelCells.ContainsKey(modelName))
            {
                _modelCells[modelName] = new HashSet<string>();
            }
            _modelCells[modelName] = new HashSet<string>(cells);
            _textEditor.TextArea.TextView.Redraw();
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

        // 한글 초성 매칭 메서드
        private bool MatchesKoreanInitial(string target, string input)
        {
            if (string.IsNullOrEmpty(input)) return true;

            var targetInitials = GetKoreanInitials(target);
            var inputInitials = GetKoreanInitials(input);

            return targetInitials.StartsWith(inputInitials, StringComparison.OrdinalIgnoreCase);
        }

        // 한글 초성 추출 메서드
        private string GetKoreanInitials(string text)
        {
            var result = new System.Text.StringBuilder();
            foreach (char c in text)
            {
                if (c >= '가' && c <= '힣')
                {
                    char initial = (char)((c - '가') / 588 + 'ㄱ');
                    result.Append(initial);
                }
                else
                {
                    result.Append(c);
                }
            }
            return result.ToString();
        }
    }

    public enum CompletionType
    {
        Cell,
        Function,
        Model,
        ContextVariable,
        Assumption
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
            return new TextBlock
            {
                Text = Text,
                Foreground = CompletionType switch
                {
                    CompletionType.Cell => Brushes.Orange,
                    CompletionType.Function => Brushes.DarkCyan,
                    CompletionType.Model => Brushes.HotPink,
                    CompletionType.ContextVariable => Brushes.LightBlue,
                    _ => Brushes.Black
                },
                FontSize = 11.5
            };
        }

        public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs)
        {
            string insertText;

            if (CompletionType == CompletionType.Cell)
            {
                // Get the text from the start of the line to the completion segment
                var textBeforeSegment = textArea.Document.GetText(completionSegment.Offset, completionSegment.Length);
                var modelMatch = Regex.Match(textBeforeSegment, @"(\w+)\.\s*$");

                if (modelMatch.Success)
                {
                    string modelName = modelMatch.Groups[1].Value;
                    insertText = $"{modelName}.{Text}[t]";
                }
                else
                {
                    insertText = $"{Text}[t]";
                }
            }
            else if (CompletionType == CompletionType.Function)
            {
                insertText = Text == "Assum" ? $"{Text}(\"\")" : $"{Text}()";
            }
            else if (CompletionType == CompletionType.Model)
            {
                insertText = $"{Text}.";
            }
            else if (CompletionType == CompletionType.Assumption)
            {
                insertText = Text;
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
                    textArea.Caret.Offset -= 1;
                }
                else
                {
                    textArea.Caret.Offset -= 1;
                }
            }
        }
    }
}