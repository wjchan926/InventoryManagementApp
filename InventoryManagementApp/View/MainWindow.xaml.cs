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

namespace InventoryManagementApp.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelDocViewModel excelDocViewModel;

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
        }

        private void openBtn_Click(object sender, RoutedEventArgs e)
        {
            excelDocViewModel = new ExcelDocViewModel();
            excelDocViewModel.Open();
        }

        private void analyzeBtn_Click(object sender, RoutedEventArgs e)
        {
            SOTableViewModel soTableViewModel = new SOTableViewModel();
            ItemTableViewModel itemTableViewModel = new ItemTableViewModel();
            excelDocViewModel.Analyze(itemTableViewModel.itemDataTable, soTableViewModel.soDataTable);
        }

        private void saveCloseBtn_Click(object sender, RoutedEventArgs e)
        {
            using (excelDocViewModel.excelDoc)
            {
                excelDocViewModel.excelDoc.Close();
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
    }
}
