using ActuLiteModel;
using System.Windows;
using ModernWpf;
using Newtonsoft.Json;
using System.IO;

namespace ActuLight
{
    public partial class App : Application
    {
        public static ModelEngine ModelEngine { get; private set; }
        public static AppSettingsManager SettingsManager { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ModelEngine = new ModelEngine();
            SettingsManager = new AppSettingsManager();
            LoadAndApplySettings();
        }

        private void LoadAndApplySettings()
        {
            SettingsManager.LoadSettings();
        }

        public static void ApplyTheme()
        {
            if (SettingsManager.CurrentSettings.Theme == "Dark")
            {
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
            }
            else
            {
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
            }
        }

    }

    public class AppSettingsManager
    {
        private const string SettingsFileName = "settings.json";
        private static string SettingsFilePath => Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, SettingsFileName);

        public AppSettings CurrentSettings { get; private set; }

        public void LoadSettings()
        {
            if (File.Exists(SettingsFilePath))
            {
                string json = File.ReadAllText(SettingsFilePath);
                CurrentSettings = JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
            }
            else
            {
                CurrentSettings = new AppSettings();
            }
        }

        public void SaveSettings()
        {
            string json = JsonConvert.SerializeObject(CurrentSettings, Formatting.Indented);
            File.WriteAllText(SettingsFilePath, json);
        }
    }

    public class AppSettings
    {
        public string Theme { get; set; } = "Light";
        public int SignificantDigits { get; set; } = 8;
    }
}