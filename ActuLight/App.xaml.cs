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
            ApplyTheme();
        }

        public static void ApplyTheme()
        {
            ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;

            if (SettingsManager.CurrentSettings.Theme == "Dark")
            {               
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
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
                CurrentSettings = JsonConvert.DeserializeObject<AppSettings>(json) ?? CreateDefaultSettings();
            }
            else
            {
                CurrentSettings = CreateDefaultSettings();
                SaveSettings(); // 기본 설정을 파일로 저장
            }
        }

        private AppSettings CreateDefaultSettings()
        {
            return new AppSettings
            {
                Theme = "Dark",
                SignificantDigits = 8,
                DataGridSortOption = DataGridSortOption.Default
            };
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
        public DataGridSortOption DataGridSortOption { get; set; } = DataGridSortOption.Default;
    }

    public enum DataGridSortOption
    {
        Default,
        CellDefinitionOrder,
        Alphabetical
    }
}