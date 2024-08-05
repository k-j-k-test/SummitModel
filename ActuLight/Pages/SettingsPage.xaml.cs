using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ModernWpf;

namespace ActuLight.Pages
{
    public partial class SettingsPage : Page
    {     
        public SettingsPage()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            var settings = App.SettingsManager.CurrentSettings;
            ThemeSelector.SelectedIndex = settings.Theme == "Dark" ? 1 : 0;
            SignificantDigitsSelector.SelectedItem = SignificantDigitsSelector.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(item => int.Parse((string)item.Content) == settings.SignificantDigits);
        }

        private void ThemeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ThemeSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                string selectedTheme = ((string)selectedItem.Content).Replace(" Theme", "");
                App.SettingsManager.CurrentSettings.Theme = selectedTheme;
                App.SettingsManager.SaveSettings();
                App.ApplyTheme();
            }
        }

        private void SignificantDigitsSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SignificantDigitsSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                App.SettingsManager.CurrentSettings.SignificantDigits = int.Parse((string)selectedItem.Content);
                App.SettingsManager.SaveSettings();

                var mainWindow = (MainWindow)Application.Current.MainWindow;
                if (mainWindow.pageCache.TryGetValue("Pages/SpreadsheetPage.xaml", out var page) && page is SpreadSheetPage spreadSheetPage)
                {
                    spreadSheetPage.SignificantDigits = App.SettingsManager.CurrentSettings.SignificantDigits;
                    spreadSheetPage.UpdateInvokes();
                }
            }
        }
    }
}