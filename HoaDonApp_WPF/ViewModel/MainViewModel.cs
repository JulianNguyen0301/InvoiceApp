using HoaDonApp_WPF.View;
using Microsoft.Win32;
using QuanLyKhoWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace HoaDonApp_WPF.ViewModel
{
    internal class MainViewModel : BaseViewModel
    {
        public ICommand LoadedWindowCommand { get; set; }
        public MainViewModel()
        {
            LoadedWindowCommand = new RelayCommand<Window>((p) => { return true; }, (p) => {
                if (p == null)
                    return;

                p.Hide();

                LoginView loginView = new LoginView();
                loginView.ShowDialog();

                if (loginView.DataContext == null)
                    return;

                var loginVM = loginView.DataContext as LoginViewModel;

                if (loginVM.IsLogin == true)
                {
                    p.Show();
                }
                else p.Close();
            });
        }
    }
}
