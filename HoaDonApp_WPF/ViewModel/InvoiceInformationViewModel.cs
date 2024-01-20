using QuanLyKhoWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;
using System.Collections.ObjectModel;

namespace HoaDonApp_WPF.ViewModel
{
    public class InvoiceInformationViewModel : BaseViewModel
    {
        public ObservableCollection<string> _InvoiceExport;
        public ObservableCollection<string> InvoiceExport {  get; set; } 
        public bool _IsEnable;
        public bool IsEnable { get => _IsEnable; set { _IsEnable = value; OnPropertyChanged(); } }

        public bool _IsSelectedLogo;
        public bool IsSelectedLogo { get => _IsSelectedLogo; set { _IsSelectedLogo = value; OnPropertyChanged(); } }

        public bool _IsSelectedInvoice;
        public bool IsSelectedInvoice { get => _IsSelectedInvoice; set { _IsSelectedInvoice = value; OnPropertyChanged(); } }
        public ICommand DataButton { get; set; }
        public ICommand LoadingButton { get; set; }
        public string _TxtBox;
        public string TxtBox { get => _TxtBox; set { _TxtBox = value; OnPropertyChanged(); } }
        public InvoiceInformationViewModel()
        {
            IsEnable = false;
            IsSelectedLogo = false;
            IsSelectedInvoice = true;
            DataButton = new RelayCommand<Window>((p) => { return true; },

            (p) =>
            {
                // Create OpenFileDialog 
                OpenFileDialog dlg = new OpenFileDialog();

                dlg.Multiselect = true;

                // Set filter for file extension and default file extension 
                dlg.Filter = "JPEG Files (*.pdf)|*.pdf";

                // Display OpenFileDialog by calling ShowDialog method 

                Nullable<bool> result = dlg.ShowDialog();

                // Get the selected file name and display in a TextBox 
                if (result == true)
                {
                    int i = 0;
                    // Open document 
                    foreach (string filename in dlg.FileNames)
                    {
                        i++;
                        //string dest = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + @"\VIAM_INVOICE_PROJECT\PDF for E-Invoice\" + Path.GetFileName(filename);
                        string dest = Directory.GetCurrentDirectory() + @"\VIAM_INVOICE_PROJECT\PDF for E-Invoice\" + Path.GetFileName(filename);
                        File.Copy(filename, dest, true);
                    }
                    IsEnable = true;
                    TxtBox = $"Bạn đã chọn {i} file.";
                }
            });

            LoadingButton = new RelayCommand<MainWindow>((p) =>
            {
                if (TxtBox != null)
                    return true;
                else
                    return false;
            },


            (p) => {

                string fileName = Directory.GetCurrentDirectory() + @"\Invoice_Online.py";
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();

                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.FileName = Directory.GetCurrentDirectory() + @"\VIAM_INVOICE_PROJECT\Python\python.exe";
                startInfo.Arguments = fileName;

                process.StartInfo = startInfo;
                
                LoadingScreen loadingScreen = new LoadingScreen();
                loadingScreen.Show();

                process.Start();

                process.WaitForExit();

                loadingScreen.Close();
                var output = process.StandardOutput.ReadToEnd();

                char[] chars = output.ToCharArray();

                foreach (char c in chars)
                {
                    switch (c)
                    {
                        case '!':
                            MessageBox.Show("Hóa đơn không hợp lệ.\nCông ty Linh Khoa đã thông báo về việc tạm ngừng hoạt động có thời hạn và được cơ quan có thẩm quyền chấp thuận.\nVui lòng kiểm tra lại.");
                            goto end;
                        case '@':
                            MessageBox.Show("Hóa đơn không hợp lệ.\nDoanh nghiệp Linh Khoa không hoàn thành các nghĩa vụ thuế.\nVui lòng kiểm tra lại.");
                            goto end;
                        case '#':
                            MessageBox.Show("Hóa đơn không hợp lệ.\nCông ty Linh Khoa đang tra cứu đã bị cơ quan thuế quản lý khóa mã số thuế do doanh nghiệp không hoạt động tại địa điểm như đã đăng ký trên Giấy chứng nhận đăng ký kinh doanh.\nVui lòng kiểm tra lại.");
                            goto end;
                        case '$':
                            MessageBox.Show("Hóa đơn không hợp lệ.\nCông ty Cocacola đã thông báo về việc tạm ngừng hoạt động có thời hạn và được cơ quan có thẩm quyền chấp thuận.\nVui lòng kiểm tra lại.");
                            goto end;
                        case '%':
                            MessageBox.Show("Hóa đơn không hợp lệ.\nDoanh nghiệp Cocacola không hoàn thành các nghĩa vụ thuế.\nVui lòng kiểm tra lại.");
                            goto end;
                        case '&':
                            MessageBox.Show("Hóa đơn không hợp lệ.\nCông ty Cocacola đang tra cứu đã bị cơ quan thuế quản lý khóa mã số thuế do doanh nghiệp không hoạt động tại địa điểm như đã đăng ký trên Giấy chứng nhận đăng ký kinh doanh.\nVui lòng kiểm tra lại.");
                            goto end;
                    }
                }

                DirectoryInfo directoryInfo = new DirectoryInfo(Directory.GetCurrentDirectory() + @"\VIAM_INVOICE_PROJECT\PDF for E-Invoice");
                var fileCount = directoryInfo.GetFiles().Length;

                if (fileCount == 0)
                    MessageBox.Show("Quá trình trích xuất đã hoàn tất.\nToàn bộ file đã được trích xuất thành công.");
                else
                    MessageBox.Show($"Quá trình trích xuất đã hoàn tất.\nCó {fileCount} hóa đơn không hợp lệ không được trích xuất.\nXui lòng kiểm tra lại:\n" + Directory.GetCurrentDirectory() + @"\VIAM_INVOICE_PROJECT\PDF for E-Invoice");
                end:

                TxtBox = null;
            });
        }
    }
}

