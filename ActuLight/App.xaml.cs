using ActuLiteModel;
using System.Text.Json;
using System.Windows;
using System.IO;
using System.Text.Json;

namespace ActuLight
{
    public partial class App : Application
    {
        public static AppSettingsManager Settings { get; private set; }
        public static ModelEngine ModelEngine { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ModelEngine = new ModelEngine();
            Settings = AppSettingsManager.Load();
            LoadTheme();
        }

        private void LoadTheme()
        {
            string themeName = Settings.CurrentTheme;

            // ColourDictionaries 업데이트
            UpdateResourceDictionary("ColourDictionaries", $"Themes/ColourDictionaries/{themeName}.xaml");

            // ControlColours 업데이트
            UpdateResourceDictionary("ControlColours", "Themes/ControlColours.xaml");

            // Controls 업데이트
            UpdateResourceDictionary("Controls", "Themes/Controls.xaml");
        }

        private void UpdateResourceDictionary(string dictionaryName, string newSource)
        {
            var resourceDict = Resources.MergedDictionaries
                .FirstOrDefault(d => d.Source != null && d.Source.OriginalString.Contains(dictionaryName));

            if (resourceDict != null)
            {
                resourceDict.Source = new Uri(newSource, UriKind.Relative);
            }
            else
            {
                // 해당 ResourceDictionary가 없으면 새로 추가
                Resources.MergedDictionaries.Add(new ResourceDictionary
                {
                    Source = new Uri(newSource, UriKind.Relative)
                });
            }
        }
    }

    public class AppSettingsManager
    {
        private const string SettingsFileName = "appSettings.json";
        private static string SettingsFilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SettingsFileName);

        public string CurrentTheme { get; set; } = "DarkGreyTheme"; // 기본 테마

        public static AppSettingsManager Load()
        {
            if (File.Exists(SettingsFilePath))
            {
                string json = File.ReadAllText(SettingsFilePath);
                return JsonSerializer.Deserialize<AppSettingsManager>(json) ?? new AppSettingsManager();
            }
            return new AppSettingsManager();
        }

        public void Save()
        {
            string json = JsonSerializer.Serialize(this);
            File.WriteAllText(SettingsFilePath, json);
        }
    }
}