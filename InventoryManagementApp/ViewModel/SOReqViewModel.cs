using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Collections.ObjectModel;
using InventoryManagementApp.Model;

namespace InventoryManagementApp.ViewModel
{
    /// <summary>
    /// Data binds the soReqDataTable to the soReqDataTableView on GUI
    /// </summary>
    sealed class SOReqViewModel : INotifyPropertyChanged
    {
        private DataTable soReqDataTable;

        public DataTable SOReqDataTable
        {
            get
            {
                return soReqDataTable;
            }
            set
            {
                soReqDataTable = value;
                OnPropertyChanged("SOReqDataTable");
            }
        }

        SOReqViewModel()
        {

        }

        public SOReqViewModel(DataTable minMaxDt)
        {
            soReqDataTable = new DataTable().BuildSOReqTable(minMaxDt);
            OnPropertyChanged("SOReqDataTable");
        }

        #region INotifyProperyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        public void Update()
        {
            OnPropertyChanged("SOReqDataTable");
        }
    }
}
