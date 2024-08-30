﻿using System;
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

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                SaveExcelFile();
            }
        }

        public void SaveExcelFile()
        {
            string initialDirectory = null;
            string fileName = null;

            if (string.IsNullOrEmpty(lastSavedFilePath))
            {
                var filePage = pageCache["Pages/FilePage.xaml"] as FilePage;
                if (filePage != null && !string.IsNullOrEmpty(filePage.currentFilePath))
                {
                    lastSavedFilePath = filePage.currentFilePath;
                }
            }

            if (!string.IsNullOrEmpty(lastSavedFilePath))
            {
                initialDirectory = Path.GetDirectoryName(lastSavedFilePath);
                fileName = Path.GetFileName(lastSavedFilePath);
            }
            else
            {
                initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                fileName = "NewFile.xlsx";
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true,
                InitialDirectory = initialDirectory,
                FileName = fileName
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
                    if (filePage.excelData.ContainsKey("out"))
                    {
                        sheets["out"] = filePage.excelData["out"];
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

        private void AutoSaveExcelFile()
        {
            if (string.IsNullOrEmpty(lastSavedFilePath))
            {
                // 마지막으로 저장된 파일 경로가 없는 경우, 기본 경로 설정
                lastSavedFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ActuLight_AutoSave.xlsx");
            }

            string directory = Path.GetDirectoryName(lastSavedFilePath);
            string fileName = Path.GetFileNameWithoutExtension(lastSavedFilePath);
            string extension = Path.GetExtension(lastSavedFilePath);

            // "_AutoSave" 접미사를 추가한 새 파일명 생성
            string autoSaveFileName = $"{fileName}_AutoSave{extension}";
            string autoSaveFilePath = Path.Combine(directory, autoSaveFileName);

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
                ExcelExporter.ExportMultipleSheets(autoSaveFilePath, sheets);
                // 자동 저장 성공 메시지를 상태 바에 표시하거나 로그에 기록할 수 있습니다.
                // 예: StatusBar.Text = $"자동 저장 완료: {autoSaveFileName}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"자동 저장 중 오류 발생: {ex.Message}", "자동 저장 오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateWindowTitle()
        {
            this.Title = $"ActuLight {App.CurrentVersion}";
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                SaveExcelFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"자동 저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}