using ActuLiteModel;
using System.Windows;
using ModernWpf;
using Newtonsoft.Json;
using System.IO;
using System;
using System.Windows.Interop;
using System.Windows.Media;

namespace ActuLight
{
    public partial class App : Application
    {
        public static string CurrentVersion { get; private set; }

        public static ModelEngine ModelEngine { get; private set; }
        public static AppSettingsManager SettingsManager { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;
            base.OnStartup(e);
            ModelEngine = new ModelEngine();
            SettingsManager = new AppSettingsManager();
            LoadAndApplySettings();
            LoadVersionInfo();
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

        private void LoadVersionInfo()
        {
            string versionFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "version.json");
            if (File.Exists(versionFilePath))
            {
                string json = File.ReadAllText(versionFilePath);
                var versionInfo = JsonConvert.DeserializeObject<VersionInfo>(json);
                CurrentVersion = versionInfo.Version;
            }
            else
            {
                CurrentVersion = "v0.0.1"; // 기본 버전
            }
        }
    }

    public class AppSettingsManager
    {
        private const string SettingsFileName = "settings.json";
        private static string SettingsFilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SettingsFileName);

        public AppSettings CurrentSettings { get; private set; }

        public void LoadSettings()
        {
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(SettingsFilePath);
                    var loadedSettings = JsonConvert.DeserializeObject<AppSettings>(json);

                    // 새로운 설정 항목이 추가되었을 경우를 대비해 기본값으로 초기화된 새 객체 생성
                    CurrentSettings = new AppSettings();

                    // 로드된 설정값을 현재 설정에 복사
                    if (loadedSettings != null)
                    {
                        foreach (var prop in typeof(AppSettings).GetProperties())
                        {
                            var loadedValue = prop.GetValue(loadedSettings);
                            if (loadedValue != null)
                            {
                                prop.SetValue(CurrentSettings, loadedValue);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 로드 실패 시 기본 설정 사용
                    CurrentSettings = new AppSettings();
                    Console.WriteLine($"Failed to load settings: {ex.Message}. Using default settings.");
                }
            }
            else
            {
                CurrentSettings = new AppSettings();
            }
        }

        public void SaveSettings()
        {
            try
            {
                string json = JsonConvert.SerializeObject(CurrentSettings, Formatting.Indented);
                File.WriteAllText(SettingsFilePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save settings: {ex.Message}");
            }
        }
    }

    public class AppSettings
    {
        public string Theme { get; set; } = "Light";
        public int SignificantDigits { get; set; } = 8;
        public DataGridSortOption DataGridSortOption { get; set; } = DataGridSortOption.Default;
    }

    public class VersionInfo
    {
        public string Version { get; set; }
    }

    public enum DataGridSortOption
    {
        Default,
        CellDefinitionOrder,
        Alphabetical
    }
}