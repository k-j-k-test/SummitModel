using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ActuLight
{
    /// <summary>
    /// NavigationBar.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class NavigationBar : UserControl
    {
        private static NavigationBar _instance;

        public static NavigationBar Instance => _instance ??= new NavigationBar();

        public NavigationBar()
        {
            InitializeComponent();
        }

        private void FileButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("FilePage");
        }

        private void ModelPointButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("ModelPointPage");
        }

        private void AssumptionButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("AssumptionPage");
        }

        private void SpreadsheetButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("SpreadSheetPage");
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("SettingsPage");
        }

        private void OutputButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("OutputPage");
        }

        private void DataProcessButton_Click(object sender, RoutedEventArgs e)
        {
            ((MainWindow)Window.GetWindow(this)).NavigateTo("DataProcessingPage");
        }

        public void SetButtonsEnabled(bool enabled)
        {
            FileButton.IsEnabled = enabled;
            ModelPointButton.IsEnabled = enabled;
            AssumptionButton.IsEnabled = enabled;
            SpreadsheetButton.IsEnabled = enabled;
            OutputButton.IsEnabled = enabled;
            DataProcessButton.IsEnabled = enabled;
            SettingsButton.IsEnabled = enabled;
        }

    }
}
