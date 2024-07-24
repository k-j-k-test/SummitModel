using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ActuLight.Pages
{
    public partial class SettingsPage : Page
    {
        private AppSettingsManager settings;

        public SettingsPage()
        {
            InitializeComponent();
            settings = AppSettingsManager.Load();
            LoadCurrentTheme();
        }

        private void LoadCurrentTheme()
        {
            ThemeSelector.SelectedItem = ThemeSelector.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(item => ((string)item.Content).Replace(" ", "") == settings.CurrentTheme);
        }

        private void ThemeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ThemeSelector.SelectedItem is ComboBoxItem selectedItem)
            {
                string selectedTheme = ((string)selectedItem.Content).Replace(" ", "");
                ChangeTheme(selectedTheme);
            }
        }

        private void ChangeTheme(string themeName)
        {
            var app = (App)Application.Current;

            // ColourDictionaries 업데이트
            UpdateResourceDictionary(app, "ColourDictionaries", $"Themes/ColourDictionaries/{themeName}.xaml");

            // ControlColours 업데이트
            UpdateResourceDictionary(app, "ControlColours", "Themes/ControlColours.xaml");

            // Controls 업데이트
            UpdateResourceDictionary(app, "Controls", "Themes/Controls.xaml");

            settings.CurrentTheme = themeName;
            settings.Save();
        }

        private void UpdateResourceDictionary(Application app, string dictionaryName, string newSource)
        {
            var resourceDict = app.Resources.MergedDictionaries
                .FirstOrDefault(d => d.Source != null && d.Source.OriginalString.Contains(dictionaryName));

            if (resourceDict != null)
            {
                resourceDict.Source = new Uri(newSource, UriKind.Relative);
            }
            else
            {
                // 해당 ResourceDictionary가 없으면 새로 추가
                app.Resources.MergedDictionaries.Add(new ResourceDictionary
                {
                    Source = new Uri(newSource, UriKind.Relative)
                });
            }
        }
    }
}