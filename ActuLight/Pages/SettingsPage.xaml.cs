using System.Windows;
using System.Windows.Controls;
using ModernWpf;
using System.IO;
using Newtonsoft.Json;
using System.Linq;

namespace ActuLight.Pages
{
    public partial class SettingsPage : Page
    {
        public class AppSettings
        {
            public string Theme { get; set; }
            public int SignificantDigits { get; set; }
        }

        private AppSettings _settings;

        public SettingsPage()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            _settings = LoadSettingsFromFile() ?? new AppSettings { Theme = "Light", SignificantDigits = 8 };
            ThemeSelector.SelectedIndex = _settings.Theme == "Dark" ? 1 : 0;
            SignificantDigitsSelector.SelectedItem = SignificantDigitsSelector.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(item => int.Parse((string)item.Content) == _settings.SignificantDigits);
        }

        private void SaveSettings()
        {
            SaveSettingsToFile(_settings);
        }

        private AppSettings LoadSettingsFromFile()
        {
            if (File.Exists("settings.json"))
            {
                string json = File.ReadAllText("settings.json");
                return JsonConvert.DeserializeObject<AppSettings>(json);
            }
            return null;
        }

        private void SaveSettingsToFile(AppSettings settings)
        {
            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            File.WriteAllText("settings.json", json);
        }

        private void ThemeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ThemeSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                string selectedTheme = ((string)selectedItem.Content).Replace(" Theme", "");
                ChangeTheme(selectedTheme);
                _settings.Theme = selectedTheme;
                SaveSettings();
            }
        }

        private void SignificantDigitsSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SignificantDigitsSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                _settings.SignificantDigits = int.Parse((string)selectedItem.Content);
                SaveSettings();

                // SpreadSheetPage의 SignificantDigits 업데이트
                var mainWindow = (MainWindow)Application.Current.MainWindow;
                if (mainWindow.pageCache.TryGetValue("Pages/SpreadsheetPage.xaml", out var page) && page is SpreadSheetPage spreadSheetPage)
                {
                    spreadSheetPage.SignificantDigits = _settings.SignificantDigits;
                }
            }
        }

        private void ChangeTheme(string themeName)
        {
            if (themeName == "Dark")
            {
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
            }
            else
            {
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
            }
        }
    }
}