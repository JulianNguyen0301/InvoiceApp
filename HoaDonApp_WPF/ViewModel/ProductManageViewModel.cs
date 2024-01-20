using HoaDonApp_WPF.View;
using QuanLyKhoWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Navigation;

namespace HoaDonApp_WPF.ViewModel
{
    public class ProductManageViewModel : BaseViewModel
    {
        public ICommand AddProduct { get; set; }
        public ProductManageViewModel() 
        {
            AddProduct = new RelayCommand<object>((p) => { return true; }, (p) => { AddProductView wd = new AddProductView(); wd.ShowDialog(); });
        }
    }
}
