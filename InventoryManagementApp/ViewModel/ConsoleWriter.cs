using InventoryManagementApp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagementApp.ViewModel
{
    public sealed class ConsoleWriter : INotifyPropertyChanged
    {
        private static string output;
        
        public string Output
        {
            get
            {
                return output;
            }
            set
            {
                output = value;
                OnPropertyChanged("Output");
            }
        }
        
        public event PropertyChangedEventHandler PropertyChanged;

        public ConsoleWriter()
        {
            output = "Inventory Management Tool Started.\nPlease ensure QuickBooks is open and logged in.\n";
            OnPropertyChanged("Output");
        }

        private void OnPropertyChanged(string propertyName)
        {
            // PropertyChanged?.Invoke(null, new PropertyChangedEventArgs(propertyName));

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public void WriteLine(string value)
        {
            output = output + value + "\n";
            OnPropertyChanged("Output");
        }

        public void Clear()
        {
            output = string.Empty;
            OnPropertyChanged("Output");
        }        
    }
}
