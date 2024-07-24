using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Navigation;
using ActuLight.Pages;
using ActuLiteModel;

namespace ActuLight
{
    public partial class MainWindow : Window
    {
        public Dictionary<string, Page> pageCache = new Dictionary<string, Page>();

        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigating += MainFrame_Navigating;
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
    }
}