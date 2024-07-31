using ActuLiteModel;
using Newtonsoft.Json;
using System.Windows;
using System.IO;
using ModernWpf;
using System;

namespace ActuLight
{
    public partial class App : Application
    {
        public static ModelEngine ModelEngine { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ModelEngine = new ModelEngine();
            LoadTheme();
        }

        private void LoadTheme()
        {
            // 기본 테마를 Light로 설정
            ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
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
                return JsonConvert.DeserializeObject<AppSettingsManager>(json) ?? new AppSettingsManager();
            }
            return new AppSettingsManager();
        }

        public void Save()
        {
            string json = JsonConvert.SerializeObject(this, Formatting.Indented);
            File.WriteAllText(SettingsFilePath, json);
        }
    }
}