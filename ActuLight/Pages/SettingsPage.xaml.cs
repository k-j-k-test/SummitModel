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
            LoadCurrentTheme();
        }

        private void LoadCurrentTheme()
        {
            ThemeSelector.SelectedIndex = ThemeManager.Current.ActualApplicationTheme == ApplicationTheme.Dark ? 1 : 0;
        }

        private void ThemeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ThemeSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                string selectedTheme = ((string)selectedItem.Content).Replace(" Theme", "");
                ChangeTheme(selectedTheme);
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