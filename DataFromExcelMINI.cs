using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoParser
{
    public class DataFromExcelMINI //: INotifyPropertyChanged
    {
        public DataFromExcelMINI() { }
        public DataFromExcelMINI(string Id, string NameUBI)
        {
            this.Id = Id;
            this.NameUBI = NameUBI;
        }
        private string nameUBI;
        public string Id { get; set; }
        public string NameUBI
        {
            get { return nameUBI; }
            set 
            { 
                if(nameUBI == value) { return; }
                nameUBI = value;
               // OnPropertyChanged();
            }
        }



        //public event PropertyChangedEventHandler PropertyChanged;

        //protected virtual void OnPropertyChanged(string propertyName = "")
       // {
       //     PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName)); // проверка на null   
        //}
    }
}
