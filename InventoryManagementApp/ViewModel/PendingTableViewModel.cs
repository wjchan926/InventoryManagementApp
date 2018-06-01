using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.ComponentModel;
using InventoryManagementApp.Model;

namespace InventoryManagementApp.ViewModel
{
    /// <summary>
    /// Data binds the pendingDataTable to the pendingDataTableView on GUI.
    /// </summary>
    class PendingTableViewModel : INotifyPropertyChanged
    {
        private DataTable pendingDataTable;

        public DataTable PendingDataTable
        {
            get
            {
                return pendingDataTable;
            }
            set
            {
                pendingDataTable = value;
                OnPropertyChanged("PendingDataTable");
            }
        }

        PendingTableViewModel()
        {

        }

        public PendingTableViewModel(DataTable minMaxDt)
        {
            pendingDataTable = new DataTable().BuildPending(minMaxDt);
            OnPropertyChanged("PendingDataTable");
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
            OnPropertyChanged("PendingDataTable");
        }
    }
}
