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
using System.Data;

namespace InventoryManagementApp.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelDocViewModel excelDocViewModel;
        LogViewModel logViewModel;
        SOReqViewModel soReqViewModel;

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

            logViewModel = new LogViewModel();

            outputTb.SetBinding(TextBox.TextProperty, new Binding("Output")
            {
                Source = logViewModel,
                Mode = BindingMode.OneWay
            });
        }        

        private void openBtn_Click(object sender, RoutedEventArgs e)
        {
            excelDocViewModel = new ExcelDocViewModel();
            excelDocViewModel.Open();

            logViewModel.UpdateStatus();
            outputTb.ScrollToEnd();
        }

        private void analyzeBtn_Click(object sender, RoutedEventArgs e)
        {
            Log.WriteLine("...Analyzing Part Numbers...");

            logViewModel.UpdateStatus();
            outputTb.ScrollToEnd();

            openBtn.IsEnabled = false;
            analyzeBtn.IsEnabled = false;
            saveCloseBtn.IsEnabled = false;
            exitBtn.IsEnabled = false;

            SOTableViewModel soTableViewModel = new SOTableViewModel();
            ItemTableViewModel itemTableViewModel = new ItemTableViewModel();
            DataTable minMaxDt = excelDocViewModel.Analyze(itemTableViewModel.itemDataTable, soTableViewModel.soDataTable);

            soReqViewModel = new SOReqViewModel(minMaxDt);

            soReqDataGrid.SetBinding(DataGrid.ItemsSourceProperty, new Binding("SOReqDataTable")
            {
                Source = soReqViewModel,
                Mode = BindingMode.TwoWay
            });

            logViewModel.UpdateStatus();
            outputTb.ScrollToEnd();

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

                logViewModel.UpdateStatus();
                outputTb.ScrollToEnd();
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

        private void clearBtn_Click(object sender, RoutedEventArgs e)
        {
            Log.Clear();

            logViewModel.UpdateStatus();
            outputTb.ScrollToEnd();
        }
    }
}
