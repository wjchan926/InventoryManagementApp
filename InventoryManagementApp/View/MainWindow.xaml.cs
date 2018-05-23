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
using InventoryManagementApp.Model;
using InventoryManagementApp.ViewModel;
using System.Deployment.Application;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace InventoryManagementApp.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelDocViewModel excelDocViewModel;
        ConsoleWriter consoleWriter;

        private string version;
        
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch { }
            Title = "Inventory Mangement Tool V" + version;
            Topmost = true;

            consoleWriter = new ConsoleWriter();

            outputTb.SetBinding(TextBox.TextProperty, new Binding("Output")
            {
                Source = consoleWriter,
                Mode = BindingMode.OneWay
            });
        }
        

        private void openBtn_Click(object sender, RoutedEventArgs e)
        {
            excelDocViewModel = new ExcelDocViewModel();
            excelDocViewModel.Open();
            consoleWriter.WriteLine("Min-Max Document Opened.");
        }

        private void analyzeBtn_Click(object sender, RoutedEventArgs e)
        {
            consoleWriter.WriteLine("...Analyzing Part Numbers...");
            openBtn.IsEnabled = false;
            analyzeBtn.IsEnabled = false;
            saveCloseBtn.IsEnabled = false;
            exitBtn.IsEnabled = false;

            SOTableViewModel soTableViewModel = new SOTableViewModel();
            ItemTableViewModel itemTableViewModel = new ItemTableViewModel();
            excelDocViewModel.Analyze(itemTableViewModel.itemDataTable, soTableViewModel.soDataTable);

            consoleWriter.WriteLine("Analysis Complete.\nPlease Refer to Min-Max Document.");

            openBtn.IsEnabled = true;
            analyzeBtn.IsEnabled = true;
            saveCloseBtn.IsEnabled = true;
            exitBtn.IsEnabled = true;
        }

        private void saveCloseBtn_Click(object sender, RoutedEventArgs e)
        {
            using (excelDocViewModel.excelDoc)
            {
                excelDocViewModel.excelDoc.Close();
                consoleWriter.WriteLine("Min-Max Document Saved and Closed.");
            }

        }

        private void exitBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                excelDocViewModel.excelDoc.Dispose();
            }
            catch { }

            this.Close();
        }

        private void outputTb_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void clearBtn_Click(object sender, RoutedEventArgs e)
        {
            consoleWriter.Clear();   
        }
    }
}
