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
using System.Diagnostics;
using System.Threading;

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
        PendingTableViewModel pendingTableViewModel;
        DataTable minMaxDt;
        SOTableViewModel soTableViewModel;
        ItemTableViewModel itemTableViewModel;

        private string version;
        
        public MainWindow()
        {
            InitializeComponent();

            // Give program a second to spool up
            Thread.Sleep(1000);

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
            excelDocViewModel.SetExcelObjects();

            logViewModel.Update();
            outputTb.ScrollToEnd();
        }

        private void analyzeBtn_Click(object sender, RoutedEventArgs e)
        {   
            if(excelDocViewModel == null)
            {
                excelDocViewModel = new ExcelDocViewModel();
                try
                {
                    excelDocViewModel.SetExcelObjects();
                }
                catch { }
            }
            else
            {
                try
                {
                    excelDocViewModel.SetExcelObjects();
                }
                catch { }
            }

            if (excelDocViewModel != null && excelDocViewModel.excelObjSet)
            {
                Log.WriteLine("...Analyzing Part Numbers...");

                logViewModel.Update();
                outputTb.ScrollToEnd();

                openBtn.IsEnabled = false;
                analyzeBtn.IsEnabled = false;
                saveCloseBtn.IsEnabled = false;

                soTableViewModel = new SOTableViewModel();
                itemTableViewModel = new ItemTableViewModel();

                minMaxDt = excelDocViewModel.Analyze(itemTableViewModel.itemDataTable, soTableViewModel.soDataTable);

                soReqViewModel = new SOReqViewModel(minMaxDt);
                pendingTableViewModel = new PendingTableViewModel(minMaxDt);

                soReqDataGrid.SetBinding(DataGrid.ItemsSourceProperty, new Binding("SOReqDataTable")
                {
                    Source = soReqViewModel,
                    Mode = BindingMode.TwoWay
                });

                pendingDataGrid.SetBinding(DataGrid.ItemsSourceProperty, new Binding("PendingDataTable")
                {
                    Source = pendingTableViewModel,
                    Mode = BindingMode.OneWay
                });

                openBtn.IsEnabled = true;
                analyzeBtn.IsEnabled = true;
                saveCloseBtn.IsEnabled = true;
                updateExcelBtn.IsEnabled = true;
            }
            else
            {
                Log.WriteLine("Min-Max Document Not Found.");

            }

            logViewModel.Update();
            outputTb.ScrollToEnd();
        }

        private void saveCloseBtn_Click(object sender, RoutedEventArgs e)
        {
            if (excelDocViewModel == null)
            {
                excelDocViewModel = new ExcelDocViewModel();
                try
                {
                    excelDocViewModel.SetExcelObjects();
                }
                catch { }
            }
            else
            {
                try
                {
                    excelDocViewModel.SetExcelObjects();
                }
                catch { }
            }

            if (excelDocViewModel != null && excelDocViewModel.excelObjSet)
            {
                using (excelDocViewModel.excelDoc)
                {
                    excelDocViewModel.excelDoc.Close();
                }
                updateExcelBtn.IsEnabled = false;
            }
            else
            {
                Log.WriteLine("Min-Max Document Not Found.");
            }
            logViewModel.Update();
            outputTb.ScrollToEnd();
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

            logViewModel.Update();
            outputTb.ScrollToEnd();
        }

        private void updateExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                excelDocViewModel.UpdateSO(soReqViewModel.SOReqDataTable);
            }
            catch (Exception)
            {
                Log.WriteLine("Min-Max Update Failed.");
            }
            
            logViewModel.Update();
            outputTb.ScrollToEnd();
        }
    }
}
