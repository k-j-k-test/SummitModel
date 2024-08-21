﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        private string lastSavedFilePath;

        public MainWindow()
        {
            InitializeComponent();
            UpdateWindowTitle();

            MainFrame.Navigating += MainFrame_Navigating;

            // Ctrl+S 키 이벤트 핸들러 추가
            this.KeyDown += MainWindow_KeyDown;

            // 초기 페이지를 FilePage로 설정
            NavigateTo("FilePage");

            Application.Current.MainWindow.InvalidateVisual();
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

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                SaveExcelFile();
            }
        }

        public void SaveExcelFile()
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true
            };

            if (!string.IsNullOrEmpty(lastSavedFilePath))
            {
                saveFileDialog.FileName = Path.GetFileName(lastSavedFilePath);
                saveFileDialog.InitialDirectory = Path.GetDirectoryName(lastSavedFilePath);
            }

            if (saveFileDialog.ShowDialog() == true)
            {
                var filePath = saveFileDialog.FileName;
                var sheets = new Dictionary<string, object>();

                // mp 및 assum 데이터 가져오기
                var filePage = pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage != null && filePage.excelData != null)
                {
                    if (filePage.excelData.ContainsKey("mp"))
                    {
                        sheets["mp"] = filePage.excelData["mp"];
                    }
                    if (filePage.excelData.ContainsKey("assum"))
                    {
                        sheets["assum"] = filePage.excelData["assum"];
                    }
                }

                // Scripts 내용을 각 시트로 저장
                var spreadSheetPage = pageCache["Pages/SpreadsheetPage.xaml"] as SpreadSheetPage;
                if (spreadSheetPage != null && spreadSheetPage.Scripts != null)
                {
                    foreach (var script in spreadSheetPage.Scripts)
                    {
                        var scriptData = new List<List<object>>();
                        var lines = script.Value.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                        foreach (var line in lines)
                        {
                            scriptData.Add(new List<object> { line });
                        }
                        sheets[script.Key] = scriptData;
                    }
                }

                try
                {
                    ExcelExporter.ExportMultipleSheets(filePath, sheets);
                    lastSavedFilePath = filePath;
                    MessageBox.Show("Excel file saved successfully.", "Save Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving Excel file: {ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateWindowTitle()
        {
            this.Title = $"ActuLight {App.CurrentVersion}";
        }
    }
}