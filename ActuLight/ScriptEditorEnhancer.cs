using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Search;
using ICSharpCode.AvalonEdit.Document;
using ModernWpf.Controls;
using ModernWpf;
using Windows.ApplicationModel.Search;
using System.Windows.Media;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Controls.Primitives;
using ICSharpCode.AvalonEdit.Editing;
using System.Windows.Documents;
using System.Diagnostics;
using System.Windows.Threading;

public class ScriptEditorEnhancer
{
    private readonly TextEditor _editor;
    private CustomSearchPanel _searchPanel;
    private CustomReplacePanel _replacePanel;

    public ScriptEditorEnhancer(TextEditor editor)
    {
        _editor = editor ?? throw new ArgumentNullException(nameof(editor));
        InitializeEnhancements();
    }

    private void InitializeEnhancements()
    {
        // 키 바인딩 추가
        _editor.InputBindings.Add(new KeyBinding(FindCommand, new KeyGesture(Key.F, ModifierKeys.Control)));
        _editor.InputBindings.Add(new KeyBinding(ReplaceCommand, new KeyGesture(Key.H, ModifierKeys.Control)));

        _editor.InputBindings.Add(new KeyBinding(ToggleCommentCommand, new KeyGesture(Key.OemQuestion, ModifierKeys.Control)));
    }

    // 찾기 명령
    private ICommand FindCommand => new RelayCommand(OpenFindDialog);
    private void OpenFindDialog()
    {
        if (_searchPanel == null || !_searchPanel.IsVisible)
        {
            _searchPanel = new CustomSearchPanel(_editor);
            _searchPanel.Show();
        }
        else
        {
            _searchPanel.Activate();
        }
    }

    // 바꾸기 명령
    private ICommand ReplaceCommand => new RelayCommand(OpenReplaceDialog);
    private void OpenReplaceDialog()
    {
        if (_replacePanel == null || !_replacePanel.IsVisible)
        {
            _replacePanel = new CustomReplacePanel(_editor);
            _replacePanel.Show();
        }
        else
        {
            _replacePanel.Activate();
        }
    }

    // 주석 토글 명령
    private ICommand ToggleCommentCommand => new RelayCommand(ToggleComment);

    private void ToggleComment()
    {
        var textArea = _editor.TextArea;
        var document = textArea.Document;
        using (document.RunUpdate())
        {
            if (textArea.Selection.IsEmpty)
            {
                var currentLine = document.GetLineByOffset(textArea.Caret.Offset);
                ToggleCommentForLine(document, currentLine);
            }
            else
            {
                var startLine = document.GetLineByOffset(textArea.Selection.SurroundingSegment.Offset);
                var endLine = document.GetLineByOffset(textArea.Selection.SurroundingSegment.EndOffset);
                for (int lineNumber = startLine.LineNumber; lineNumber <= endLine.LineNumber; lineNumber++)
                {
                    var line = document.GetLineByNumber(lineNumber);
                    ToggleCommentForLine(document, line);
                }
            }
        }
    }

    private void ToggleCommentForLine(TextDocument document, DocumentLine line)
    {
        var lineText = document.GetText(line.Offset, line.Length);
        if (lineText.TrimStart().StartsWith("//"))
        {
            // 주석 제거
            int commentStart = lineText.IndexOf("//");
            document.Remove(line.Offset + commentStart, 2);
        }
        else
        {
            // 주석 추가
            document.Insert(line.Offset, "//");
        }
    }
}

public class RelayCommand : ICommand
{
    private readonly Action _execute;
    private readonly Func<bool> _canExecute;

    public RelayCommand(Action execute, Func<bool> canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public event EventHandler CanExecuteChanged
    {
        add { CommandManager.RequerySuggested += value; }
        remove { CommandManager.RequerySuggested -= value; }
    }

    public bool CanExecute(object parameter) => _canExecute == null || _canExecute();
    public void Execute(object parameter) => _execute();
}

public class CustomSearchPanel : Window
{
    private readonly TextEditor _editor;
    private TextBox _searchTextBox;
    private Button _findNextButton;
    private Button _findPreviousButton;
    private CheckBox _matchCaseCheckBox;
    private CheckBox _wholeWordCheckBox;
    private TextBlock _resultTextBlock;

    private int _currentSearchIndex = -1;
    private List<TextSegment> _searchResults = new List<TextSegment>();

    public CustomSearchPanel(TextEditor editor)
    {
        _editor = editor;
        InitializeComponent();

        this.Loaded += (s, e) =>
        {
            this.Topmost = true;
            _editor.Focus();
        };

        this.Closed += (s, e) =>
        {
            this.Topmost = false;
        };

        _searchTextBox.Focus();
    }

    private void InitializeComponent()
    {
        Title = "Find";
        Width = 300;
        Height = 200;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        // ModernWpf 스타일 적용
        Style = (Style)Application.Current.Resources[typeof(Window)];

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        _searchTextBox = new TextBox { Margin = new Thickness(5) };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;
        grid.Children.Add(_searchTextBox);

        var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(5) };
        Grid.SetRow(buttonPanel, 1);

        _findNextButton = new Button { Content = "Find Next", Margin = new Thickness(0, 0, 5, 0) };
        _findNextButton.Click += FindNextButton_Click;
        buttonPanel.Children.Add(_findNextButton);

        _findPreviousButton = new Button { Content = "Find Previous", Margin = new Thickness(0, 0, 5, 0) };
        _findPreviousButton.Click += FindPreviousButton_Click;
        buttonPanel.Children.Add(_findPreviousButton);

        grid.Children.Add(buttonPanel);

        var checkBoxPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(5) };
        Grid.SetRow(checkBoxPanel, 2);

        _matchCaseCheckBox = new CheckBox { Content = "Match case", Margin = new Thickness(0, 0, 5, 0) };
        _matchCaseCheckBox.Checked += UpdateSearch;
        _matchCaseCheckBox.Unchecked += UpdateSearch;
        checkBoxPanel.Children.Add(_matchCaseCheckBox);

        _wholeWordCheckBox = new CheckBox { Content = "Whole word", Margin = new Thickness(0, 0, 5, 0) };
        _wholeWordCheckBox.Checked += UpdateSearch;
        _wholeWordCheckBox.Unchecked += UpdateSearch;
        checkBoxPanel.Children.Add(_wholeWordCheckBox);

        grid.Children.Add(checkBoxPanel);

        _resultTextBlock = new TextBlock { Margin = new Thickness(5) };
        Grid.SetRow(_resultTextBlock, 3);
        grid.Children.Add(_resultTextBlock);

        Content = grid;

        //FindNext, Close 키다운 이벤트 설정
        _searchTextBox.KeyDown += SearchTextBox_KeyDown;
        PreviewKeyDown += CustomSearchPanel_PreviewKeyDown;
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateSearch(sender, e);
    }

    private void UpdateSearch(object sender, RoutedEventArgs e)
    {
        string searchText = _searchTextBox.Text;
        if (string.IsNullOrEmpty(searchText))
        {
            _searchResults.Clear();
            _currentSearchIndex = -1;
            UpdateResultText();
            return;
        }

        _searchResults = FindAll(searchText, _matchCaseCheckBox.IsChecked == true, _wholeWordCheckBox.IsChecked == true);
        _currentSearchIndex = -1;
        UpdateResultText();

        if (_searchResults.Count > 0)
        {
            HighlightNext();
        }
    }

    private List<TextSegment> FindAll(string searchText, bool matchCase, bool wholeWord)
    {
        var results = new List<TextSegment>();
        string text = _editor.Text;
        int start = 0;

        StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        while (start < text.Length)
        {
            int index = text.IndexOf(searchText, start, comparison);
            if (index == -1) break;

            if (wholeWord)
            {
                if ((index == 0 || !char.IsLetterOrDigit(text[index - 1])) &&
                    (index + searchText.Length == text.Length || !char.IsLetterOrDigit(text[index + searchText.Length])))
                {
                    results.Add(new TextSegment { StartOffset = index, Length = searchText.Length });
                }
            }
            else
            {
                results.Add(new TextSegment { StartOffset = index, Length = searchText.Length });
            }

            start = index + 1;
        }

        return results;
    }

    private void FindNextButton_Click(object sender, RoutedEventArgs e)
    {
        HighlightNext();
    }

    private void FindPreviousButton_Click(object sender, RoutedEventArgs e)
    {
        HighlightPrevious();
    }

    private void HighlightNext()
    {
        if (_searchResults.Count == 0) return;

        _currentSearchIndex = (_currentSearchIndex + 1) % _searchResults.Count;
        HighlightCurrentResult();
    }

    private void HighlightPrevious()
    {
        if (_searchResults.Count == 0) return;

        _currentSearchIndex = (_currentSearchIndex - 1 + _searchResults.Count) % _searchResults.Count;
        HighlightCurrentResult();
    }

    private void HighlightCurrentResult()
    {
        if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
        {
            var result = _searchResults[_currentSearchIndex];
            _editor.Select(result.StartOffset, result.Length);
            _editor.ScrollTo(_editor.TextArea.Caret.Line, _editor.TextArea.Caret.Column);
            UpdateResultText();
        }
    }

    private void UpdateResultText()
    {
        if (_searchResults.Count == 0)
        {
            _resultTextBlock.Text = "No results found";
        }
        else
        {
            _resultTextBlock.Text = $"{_currentSearchIndex + 1} of {_searchResults.Count} results";
        }
    }

    private void CustomSearchPanel_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            this.Close();
            e.Handled = true;
        }
    }

    private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            HighlightNext();
            e.Handled = true;
        }
    }
}

public class CustomReplacePanel : Window
{
    private readonly TextEditor _editor;
    private TextBox _searchTextBox;
    private TextBox _replaceTextBox;
    private Button _findNextButton;
    private Button _replaceButton;
    private Button _replaceAllButton;
    private CheckBox _matchCaseCheckBox;
    private CheckBox _wholeWordCheckBox;
    private TextBlock _resultTextBlock;

    private int _currentSearchIndex = -1;
    private List<TextSegment> _searchResults = new List<TextSegment>();

    public CustomReplacePanel(TextEditor editor)
    {
        _editor = editor;
        InitializeComponent();

        this.Loaded += (s, e) =>
        {
            this.Topmost = true;
            _editor.Focus();
        };

        this.Closed += (s, e) =>
        {
            this.Topmost = false;
        };

        _searchTextBox.Focus();
    }

    private void InitializeComponent()
    {
        Title = "Replace";
        Width = 400;
        Height = 250;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        Style = (Style)Application.Current.Resources[typeof(Window)];

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // Find label and textbox
        grid.Children.Add(new Label { Content = "Find:", VerticalAlignment = VerticalAlignment.Center });
        _searchTextBox = new TextBox { Margin = new Thickness(5) };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;
        Grid.SetColumn(_searchTextBox, 1);
        grid.Children.Add(_searchTextBox);

        // Replace label and textbox
        var replaceLabel = new Label { Content = "Replace:", VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(replaceLabel, 1);
        grid.Children.Add(replaceLabel);

        _replaceTextBox = new TextBox { Margin = new Thickness(5) };
        Grid.SetRow(_replaceTextBox, 1);
        Grid.SetColumn(_replaceTextBox, 1);
        grid.Children.Add(_replaceTextBox);

        var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(5) };
        Grid.SetRow(buttonPanel, 2);
        Grid.SetColumnSpan(buttonPanel, 2);

        _findNextButton = new Button { Content = "Find Next", Margin = new Thickness(0, 0, 5, 0) };
        _findNextButton.Click += FindNextButton_Click;
        buttonPanel.Children.Add(_findNextButton);

        _replaceButton = new Button { Content = "Replace", Margin = new Thickness(0, 0, 5, 0) };
        _replaceButton.Click += ReplaceButton_Click;
        buttonPanel.Children.Add(_replaceButton);

        _replaceAllButton = new Button { Content = "Replace All", Margin = new Thickness(0, 0, 5, 0) };
        _replaceAllButton.Click += ReplaceAllButton_Click;
        buttonPanel.Children.Add(_replaceAllButton);

        grid.Children.Add(buttonPanel);

        var checkBoxPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(5) };
        Grid.SetRow(checkBoxPanel, 3);
        Grid.SetColumnSpan(checkBoxPanel, 2);

        _matchCaseCheckBox = new CheckBox { Content = "Match case", Margin = new Thickness(0, 0, 5, 0) };
        _matchCaseCheckBox.Checked += UpdateSearch;
        _matchCaseCheckBox.Unchecked += UpdateSearch;
        checkBoxPanel.Children.Add(_matchCaseCheckBox);

        _wholeWordCheckBox = new CheckBox { Content = "Whole word", Margin = new Thickness(0, 0, 5, 0) };
        _wholeWordCheckBox.Checked += UpdateSearch;
        _wholeWordCheckBox.Unchecked += UpdateSearch;
        checkBoxPanel.Children.Add(_wholeWordCheckBox);

        grid.Children.Add(checkBoxPanel);

        _resultTextBlock = new TextBlock { Margin = new Thickness(5) };
        Grid.SetRow(_resultTextBlock, 4);
        Grid.SetColumnSpan(_resultTextBlock, 2);
        grid.Children.Add(_resultTextBlock);

        Content = grid;

        //FindNext, Close 키다운 이벤트 설정
        _searchTextBox.KeyDown += SearchTextBox_KeyDown;
        _replaceTextBox.KeyDown += ReplaceTextBox_KeyDown;
        PreviewKeyDown += CustomSearchPanel_PreviewKeyDown;
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateSearch(sender, e);
    }

    private void UpdateSearch(object sender, RoutedEventArgs e)
    {
        string searchText = _searchTextBox.Text;
        if (string.IsNullOrEmpty(searchText))
        {
            _searchResults.Clear();
            _currentSearchIndex = -1;
            UpdateResultText();
            return;
        }

        _searchResults = FindAll(searchText, _matchCaseCheckBox.IsChecked == true, _wholeWordCheckBox.IsChecked == true);
        _currentSearchIndex = -1;
        UpdateResultText();

        if (_searchResults.Count > 0)
        {
            HighlightNext();
        }
    }

    private List<TextSegment> FindAll(string searchText, bool matchCase, bool wholeWord)
    {
        var results = new List<TextSegment>();
        string text = _editor.Text;
        int start = 0;

        StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        while (start < text.Length)
        {
            int index = text.IndexOf(searchText, start, comparison);
            if (index == -1) break;

            if (wholeWord)
            {
                if ((index == 0 || !char.IsLetterOrDigit(text[index - 1])) &&
                    (index + searchText.Length == text.Length || !char.IsLetterOrDigit(text[index + searchText.Length])))
                {
                    results.Add(new TextSegment { StartOffset = index, Length = searchText.Length });
                }
            }
            else
            {
                results.Add(new TextSegment { StartOffset = index, Length = searchText.Length });
            }

            start = index + 1;
        }

        return results;
    }

    private void FindNextButton_Click(object sender, RoutedEventArgs e)
    {
        HighlightNext();
    }

    private void ReplaceButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
        {
            var segment = _searchResults[_currentSearchIndex];
            _editor.Document.Replace(segment.StartOffset, segment.Length, _replaceTextBox.Text);
            UpdateSearch(sender, e);
        }
    }

    private void ReplaceAllButton_Click(object sender, RoutedEventArgs e)
    {
        int replacements = 0;
        var document = _editor.Document;

        using (document.RunUpdate())
        {
            for (int i = _searchResults.Count - 1; i >= 0; i--)
            {
                var segment = _searchResults[i];
                document.Replace(segment.StartOffset, segment.Length, _replaceTextBox.Text);
                replacements++;
            }
        }

        UpdateSearch(sender, e);
        MessageBox.Show($"Replaced {replacements} occurrences.", "Replace All", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void HighlightNext()
    {
        if (_searchResults.Count == 0) return;

        _currentSearchIndex = (_currentSearchIndex + 1) % _searchResults.Count;
        HighlightCurrentResult();
    }

    private void HighlightCurrentResult()
    {
        if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
        {
            var result = _searchResults[_currentSearchIndex];
            _editor.Select(result.StartOffset, result.Length);
            _editor.ScrollTo(_editor.TextArea.Caret.Line, _editor.TextArea.Caret.Column);
            UpdateResultText();
        }
    }

    private void UpdateResultText()
    {
        if (_searchResults.Count == 0)
        {
            _resultTextBlock.Text = "No results found";
        }
        else
        {
            _resultTextBlock.Text = $"{_currentSearchIndex + 1} of {_searchResults.Count} results";
        }
    }

    private void CustomSearchPanel_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            this.Close();
            e.Handled = true;
        }
    }

    private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            HighlightNext();
            e.Handled = true;
        }
    }

    private void ReplaceTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            ReplaceButton_Click(null, null);
            e.Handled = true;
        }
    }
}

