using SecureApp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace HoaDonApp_WPF
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            MainWindow window = new MainWindow();
            window.Show();
            base.OnStartup(e);
            //string abc = @"Software\Tanmay\Protect123";
            //Secure scr = new Secure();
            //List<string> listKey = new List<string>()
            //{
            //  "4ZcyUNjEKT",
            //  "fDwvGb6V1o",
            //  "ieh858hBgG",
            //  "l4yqnrQhP9",
            //  "oScSdd2udr",
            //  "wKnoJIaf74",
            //  "x1HyAZXVHt",
            //  "EHEvepym7w",
            //  "7NO0PwCC5X",
            //  "lUVAyb90is"
            //};
            //bool logic = scr.Algorithm(listKey, abc, 5, 1);
            //if (logic == true)
            //{
            //    MainWindow window = new MainWindow();
            //    window.Show();
            //    base.OnStartup(e);
            //}
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
        }
    }
}
