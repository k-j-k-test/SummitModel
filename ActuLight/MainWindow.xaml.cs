using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;
using ActuLight.Pages;
using ActuLiteModel;
using Microsoft.Win32;
using ModernWpf;

namespace ActuLight
{
    public partial class MainWindow : Window
    {
        public Dictionary<string, Page> pageCache = new Dictionary<string, Page>();

        public MainWindow()
        {
            InitializeComponent();
            UpdateWindowTitle();

            MainFrame.Navigating += MainFrame_Navigating;

            // 초기 페이지를 FilePage로 설정
            NavigateTo("FilePage");

            Application.Current.MainWindow.InvalidateVisual();

            // 키다운 이벤트 핸들러 추가
            this.KeyDown += MainWindow_KeyDown;

            // 윈도우 종료 이벤트 핸들러 추가
            this.Closing += MainWindow_Closing;
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            App.ApplyTheme();
        }

        private void MainFrame_Navigating(object sender, NavigatingCancelEventArgs e)
        {
            if (e.NavigationMode == NavigationMode.New)
            {
                if (e.Uri != null)
                {
                    string pageName = e.Uri.OriginalString;
                    if (pageCache.ContainsKey(pageName))
                    {
                        e.Cancel = true;
                        MainFrame.Navigate(pageCache[pageName]);
                    }
                }
            }
        }

        private Page CreatePage(string pageName)
        {
            switch (pageName)
            {
                case "Pages/FilePage.xaml":
                    return new FilePage();
                case "Pages/ModelPointPage.xaml":
                    return new ModelPointPage();
                case "Pages/AssumptionPage.xaml":
                    return new AssumptionPage();
                case "Pages/SpreadsheetPage.xaml":
                    return new SpreadSheetPage();
                case "Pages/OutputPage.xaml":
                    return new OutputPage();
                case "Pages/SettingsPage.xaml":
                    return new SettingsPage();
                default:
                    return null;
            }
        }

        public void NavigateTo(string pageName)
        {
            string fullPagePath = $"Pages/{pageName}.xaml";
            if (!pageCache.ContainsKey(fullPagePath))
            {
                Page newPage = CreatePage(fullPagePath);
                if (newPage != null)
                {
                    pageCache[fullPagePath] = newPage;
                }
            }

            if (pageCache.TryGetValue(fullPagePath, out Page cachedPage))
            {
                MainFrame.Navigate(cachedPage);
            }
            else
            {
                MainFrame.Navigate(new Uri(fullPagePath, UriKind.Relative));
            }
        }

        private void UpdateWindowTitle()
        {
            this.Title = $"ActuLight {App.CurrentVersion}";
        }

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.S)
            {
                
            }
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                MessageBox.Show($"자동 저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}