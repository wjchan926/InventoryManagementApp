using InventoryManagementApp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InventoryManagementApp.Model;

namespace InventoryManagementApp.ViewModel
{
    sealed class LogViewModel : INotifyPropertyChanged
    {
        // private static string output;

        public LogViewModel()
        {
            Log.WriteLine("Inventory Management Tool Started.\nPlease ensure QuickBooks is open and logged in.");
            OnPropertyChanged("Output");
        }

        public string Output
        {
            get
            {
                return Log.logSB.ToString();
            }
            private set
            {
    //            Log.WriteLine(value);
                OnPropertyChanged("Output");
            }
        }

        #region INotifyProperyChanged       
        public event PropertyChangedEventHandler PropertyChanged;           

        private void OnPropertyChanged(string propertyName)
        {
            //PropertyChangedEventHandler handler = PropertyChanged;
            //if (handler != null)
            //{
            //    handler(this, new PropertyChangedEventArgs(propertyName));
            //}

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        public void Update()
        {            
            OnPropertyChanged("Output");
        }
    }
}
