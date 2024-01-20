using QuanLyKhoWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows;
using HoaDonApp_WPF.View;

namespace HoaDonApp_WPF.ViewModel
{
    public class LoginViewModel : BaseViewModel
    {
        public bool IsLogin { get; set; }
        public ICommand LoginCommand { get; set; }
        public string _UserName;
        public string UserName { get => _UserName; set { _UserName = value; OnPropertyChanged(); } }
        public string _Password;
        public string Password { get => _Password; set { _Password = value; OnPropertyChanged(); } }
        public ICommand PasswordChangedCommand { get; set; }
        public ICommand CloseCommand { get; set; }

        public LoginViewModel()
        {
            IsLogin = false;
            LoginCommand = new RelayCommand<Window>((p) => { return true; }, (p) => {Login(p); });
            CloseCommand = new RelayCommand<Window>((p) => { return true; }, (p) => {p.Close(); });
            PasswordChangedCommand = new RelayCommand<PasswordBox>((p) => { return true; }, (p) => { Password = p.Password; });
        }
        void Login(Window p)
        {
            if (p == null)
            {
                return;
            }
            //var accCount = DataProvider.Ins.DB.Users.Where(x => x.UserName == UserName && x.Password == Password).Count();
            //if (accCount > 0)
            //{
            //    IsLogin = true;
            //    p.Close();
            //}
            if (UserName == "admin" && Password == "123")
            {
                IsLogin = true;
                p.Close();
            }
            else
            {
                IsLogin = false;
                MessageBox.Show("Tên đăng nhập hoặc mật khẩu không chính xác!","Lỗi!");
            }
        }
    }
}
