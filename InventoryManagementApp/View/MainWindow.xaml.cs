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
using System.Drawing.Printing;

namespace InventoryManagementApp.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IExcelViewModel excelDocViewModel = new ExcelDocViewModel();
        LogViewModel logViewModel;
        SOReqViewModel soReqViewModel;
        PendingTableViewModel pendingTableViewModel;
        DataTable minMaxDt;

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
     //       Topmost = true;
            
            logViewModel = new LogViewModel();

            outputTb.SetBinding(TextBox.TextProperty, new Binding("Output")
            {
                Source = logViewModel,
                Mode = BindingMode.OneWay
            });
        }        

        private void openBtn_Click(object sender, RoutedEventArgs e)
        {    
            excelDocViewModel.Open();

            logViewModel.Update();
            outputTb.ScrollToEnd();
        }

        private void analyzeBtn_Click(object sender, RoutedEventArgs e)
        {
            // Disable Buttons
            openBtn.IsEnabled = false;
            analyzeBtn.IsEnabled = false;
            saveCloseBtn.IsEnabled = false;

            // Analyze
            try 
            {

                minMaxDt = excelDocViewModel.Analyze();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
          
            // Update Status Viewer
            logViewModel.Update();
            outputTb.ScrollToEnd();

            // Create Other DataTables
            soReqViewModel = new SOReqViewModel(minMaxDt);
            pendingTableViewModel = new PendingTableViewModel(minMaxDt);

            // Bind other DataTables
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

            // Reenable Buttons
            openBtn.IsEnabled = true;
            analyzeBtn.IsEnabled = true;
            saveCloseBtn.IsEnabled = true;             
        }

        private void saveCloseBtn_Click(object sender, RoutedEventArgs e)
        {
            excelDocViewModel.Close();         
       
            logViewModel.Update();
            outputTb.ScrollToEnd();
        }

        private void exitBtn_Click(object sender, RoutedEventArgs e)
        {
            excelDocViewModel.Dispose();

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
                Log.WriteLine("Nothing to Update.");
            }
            
            logViewModel.Update();
            outputTb.ScrollToEnd();
        }

        private void updateStatistics()
        {

        }

        private void printBtn_Click(object sender, RoutedEventArgs e)
        {
            //PrintDialog printDlg = new PrintDialog();

            //bool? result = printDlg.ShowDialog();

            //if (result == true)
            //{
            //    Size pageSize = new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
            //    DataGrid tempGrid = new DataGrid();

            //    tempGrid.SetBinding(DataGrid.ItemsSourceProperty, new Binding("SOReqDataTable")
            //    {
            //        Source = soReqViewModel,
            //        Mode = BindingMode.TwoWay
            //    });

            //    tempGrid.Measure(pageSize);
            //    tempGrid.Arrange(new Rect(5, 5, pageSize.Width, pageSize.Height));
            //    printDlg.PrintVisual(tempGrid, "SO Required Data Grid Printing.");
            //}       

        }

        private void printPendingBtn_Click(object sender, RoutedEventArgs e)
        {
            //PrintDialog printDlg = new PrintDialog();

            //bool? result = printDlg.ShowDialog();

            //if (result == true)
            //{
            //    Size pageSize = new Size(printDlg.PrintableAreaWidth, printDlg.PrintableAreaHeight);
            //    DataGrid tempGrid = new DataGrid();

            //    tempGrid.SetBinding(DataGrid.ItemsSourceProperty, new Binding("PendingDataTable")
            //    {
            //        Source = pendingTableViewModel,
            //        Mode = BindingMode.OneWay
            //    });

            //    tempGrid.Measure(pageSize);
            //    tempGrid.Arrange(new Rect(5, 5, pageSize.Width, pageSize.Height));
            //    printDlg.PrintVisual(tempGrid, "Pending Build Printing.");
            //}
        }
    }
}
