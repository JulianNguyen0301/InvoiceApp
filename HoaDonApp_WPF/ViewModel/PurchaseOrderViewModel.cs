using Aspose.Words;
using Aspose.Words.Tables;
using HoaDonApp_WPF.Model;
using HoaDonApp_WPF.View;
using QuanLyKhoWPF.Model;
using QuanLyKhoWPF.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ThuVienWinform.Report.AsposeWordExtension;

namespace HoaDonApp_WPF.ViewModel
{
    public class PurchaseOrderViewModel : BaseViewModel
    {
        public ObservableCollection<Supplier> _ListSupplier;
        public ObservableCollection<Supplier> ListSupplier { get => _ListSupplier; set { _ListSupplier = value; OnPropertyChanged(); } }
        public Supplier _SelectedSupplier;
        public Supplier SelectedSupplier { get => _SelectedSupplier;
            set { 
                    _SelectedSupplier = value;
                    ListProduct = null;
                    if (_SelectedSupplier != null)
                    {
                        Address1 = SelectedSupplier.AddressSupplier;
                        ListProduct = new ObservableCollection<Product>(DataProvider.Ins.DB.Products.Where(x => x.IdSupplier == SelectedSupplier.Id));
                        OnPropertyChanged();
                    }
            } 
        }
        public ObservableCollection<Product> _ListProduct;
        public ObservableCollection<Product> ListProduct { get => _ListProduct; set { _ListProduct = value; OnPropertyChanged(); }}
        public Product _SelectedProduct;
        public Product SelectedProduct { get => _SelectedProduct;
            set 
            {
                _SelectedProduct = value;
                ListUoM = null;
                PriceProduct = null;
                if(_SelectedProduct != null)
                {
                    ListUoM = new ObservableCollection<UoM>(DataProvider.Ins.DB.UoMs.Where(x => x.NameUoM == SelectedProduct.UoM));
                    OnPropertyChanged();
                }
               
            }
        }
        
        public ObservableCollection<UoM> _ListUoM;
        public ObservableCollection<UoM> ListUoM { get => _ListUoM; set { _ListUoM = value; OnPropertyChanged(); } }

        public UoM _SelectedUoM;
        public UoM SelectedUoM { get => _SelectedUoM; 
            set 
            {
                _SelectedUoM = value; 

                if(_SelectedUoM != null)
                {
                    PriceProduct = SelectedProduct.Price;
                    OnPropertyChanged();
                }
            }
        }
        public Nullable<decimal> _PriceProduct;
        public Nullable<decimal> PriceProduct { get => _PriceProduct; set { _PriceProduct = value; OnPropertyChanged(); } }
        public PODataGrid _POSelectedListView;
        public PODataGrid POSelectedListView
        {
            get => _POSelectedListView;
            set
            {
                _POSelectedListView = value;
                if (_POSelectedListView != null)
                {
                    ListProduct = new ObservableCollection<Product>(DataProvider.Ins.DB.Products.Where(x => x.Id == POSelectedListView.POItemCodeDataGrid));
                    SelectedProduct = ListProduct.First();
                    ListUoM = new ObservableCollection<UoM>(DataProvider.Ins.DB.UoMs.Where(x => x.NameUoM == POSelectedListView.POUoMDataGrid));
                    SelectedUoM = ListUoM.First();
                    PriceProduct = Decimal.Parse(POSelectedListView.POItemPriceDataGrid);
                    Quantity = POSelectedListView.POItemQuantityDataGrid;
                    OnPropertyChanged();
                }
            }
        }
        
        public string _Address1;
        public string Address1 { get => _Address1; set { _Address1 = value; OnPropertyChanged(); } }
        public string _Address2;
        public string Address2 { get => _Address2; set { _Address2 = value; OnPropertyChanged(); } }
        public string _PIC;
        public string PIC { get => _PIC; set { _PIC = value; OnPropertyChanged(); } }
        public string _NameStock;
        public string NameStock { get => _NameStock; set { _NameStock = value; OnPropertyChanged(); } }
        public string _NO;
        public string NO { get => _NO; set { _NO = value; OnPropertyChanged(); } }
        public string _Tax;
        public string Tax { get => _Tax; set { _Tax = value; OnPropertyChanged(); } }
        public string _Note;
        public string Note { get=> _Note; set { _Note = value; OnPropertyChanged(); } }
        public string _Quantity;
        public string Quantity { get => _Quantity; set {  _Quantity = value; OnPropertyChanged(); } }
        public string _Giatritruocthue = null;
        public string Giatritruocthue { get => _Giatritruocthue; set { _Giatritruocthue = value; OnPropertyChanged(); } }
        public string _Thue= null;
        public string Thue { get => _Thue; set { _Thue = value; OnPropertyChanged(); } }
        public string _Giatrisauthue = null;
        public string Giatrisauthue { get => _Giatrisauthue; set { _Giatrisauthue = value; OnPropertyChanged(); } }
        public string _TaxString;
        public string TaxString { get => _TaxString; set { _TaxString = value; OnPropertyChanged(); } }
        public string _UntaxedString;
        public string UntaxedString { get => _UntaxedString; set { _UntaxedString = value; OnPropertyChanged(); } }
        public string _TotalString;
        public string TotalString { get => _TotalString; set { _TotalString = value; OnPropertyChanged(); } }
        public DateTime _OrderDate = DateTime.Today.Date;
        public DateTime OrderDate { get => _OrderDate; set { _OrderDate = value; OnPropertyChanged(); } } 
        public DateTime _ShipDate = DateTime.Today.AddDays(+7).Date;
        public DateTime ShipDate { get => _ShipDate; set { _ShipDate = value; OnPropertyChanged(); } }
        public ICommand AddItem {  get; set; }
        public ICommand EditItem { get; set; }
        public ICommand DeleteItem { get; set; }
        public ICommand ExportPO { get; set; }
        public ICommand NewPO { get; set; }
        public ICommand ProductManagementPO { get; set; }
        public ICommand TotalPO { get; set; }
        public ObservableCollection<PODataGrid> _POExportDatagrid;
        public ObservableCollection<PODataGrid> POExportDatagrid { get => _POExportDatagrid;  set { _POExportDatagrid = value; OnPropertyChanged(); } }
        
        public PurchaseOrderViewModel()
        {
            ListSupplier = new ObservableCollection<Supplier>(DataProvider.Ins.DB.Suppliers);
            POExportDatagrid = new ObservableCollection<PODataGrid>();

            Double Sum = 0;
            Random random = new Random();
            NO = random.NextString(9);

            ProductManagementPO = new RelayCommand<object>((p) => { return true; }, (p) => { ProductManagementView wd = new ProductManagementView(); wd.ShowDialog(); });

            AddItem = new RelayCommand<object>(
                (p) => { 
                if(Quantity != null && SelectedUoM != null && SelectedProduct != null && PriceProduct != null)
                        return true;
                else 
                        return false;
                } ,
                (p) => 
                {
                    POExportDatagrid.Add(new PODataGrid(POExportDatagrid.Count + 1, SelectedProduct.Id, SelectedProduct.NameProduct, Quantity, SelectedUoM.NameUoM, (Double.Parse(PriceProduct.ToString())).ToString("0,0,0"), (Double.Parse(Quantity) * Double.Parse(PriceProduct.ToString())).ToString("0,0,0")));

                    ListProduct = new ObservableCollection<Product>();
                    ListProduct = new ObservableCollection<Product>(DataProvider.Ins.DB.Products.Where(x => x.IdSupplier == SelectedSupplier.Id));
                    ListUoM = null;
                    PriceProduct = null;
                    Quantity = null;
                    
                });
            EditItem = new RelayCommand<object>(
                (p) => 
                {
                    if (POSelectedListView != null)
                        return true;
                    else 
                        return false;
                },
                (p) =>
                {
                    POSelectedListView.POItemQuantityDataGrid = Quantity;
                    ListProduct = new ObservableCollection<Product>();
                    ListProduct = new ObservableCollection<Product>(DataProvider.Ins.DB.Products.Where(x => x.IdSupplier == SelectedSupplier.Id));
                    Quantity = null;
                }
                );
            DeleteItem = new RelayCommand<object>(
                p =>
                {
                    if (POSelectedListView != null)
                        return true;
                    else
                        return false;
                },
                (p) =>
                {
                    int k = POSelectedListView.POIdDataGrid;
                    POExportDatagrid.Remove(POSelectedListView);
                    for(int i = k - 1; i < POExportDatagrid.Count; i++)
                    {
                        POExportDatagrid[i].POIdDataGrid = k;
                        k++;
                    }
                    ListProduct = new ObservableCollection<Product>();
                    ListProduct = new ObservableCollection<Product>(DataProvider.Ins.DB.Products.Where(x => x.IdSupplier == SelectedSupplier.Id));
                    Quantity = null;
                });

            ExportPO = new RelayCommand<object>(
                p =>
                {
                    return true;
                },
                (p) =>
                {
                    if (POExportDatagrid == null || SelectedSupplier == null || Address1 == null || Address2 == null || PIC == null || ShipDate == null || OrderDate == null || Tax == null || NameStock == null)
                    {
                        MessageBox.Show("Vui lòng điền đầy đủ thông tin!", "Lỗi!");
                    }
                    else
                    {
                        Giatritruocthue = "Giá trị trước thuế/ Untaxed Amount:";
                        Thue = "Giá trị thuế/ Tax:";
                        Giatrisauthue = "Tổng tiền/ Total:";
                        for (int i = 0; i < POExportDatagrid.Count; i++)
                        {
                            UntaxedString = (Sum += Double.Parse(POExportDatagrid[i].POItemSumDataGrid)).ToString("0,0,0");
                        }
                        TaxString = (Math.Round(Double.Parse(UntaxedString) * (Double.Parse(Tax) / 100))).ToString("0,0,0");
                        TotalString = (Math.Round(Double.Parse(UntaxedString) + Double.Parse(TaxString))).ToString("0,0,0");
                        //thay thế chữ
                        Document baoCao = new Document("po_template.docx");
                        baoCao.MailMerge.Execute(new[] { "no" }, new[] { NO });
                        baoCao.MailMerge.Execute(new[] { "printdate" }, new[] { DateTime.Now.ToString() });
                        baoCao.MailMerge.Execute(new[] { "vendor" }, new[] { SelectedSupplier.NameSupplier });
                        baoCao.MailMerge.Execute(new[] { "address1" }, new[] { Address1 });
                        baoCao.MailMerge.Execute(new[] { "address2" }, new[] { Address2 });
                        baoCao.MailMerge.Execute(new[] { "pic" }, new[] { PIC });
                        baoCao.MailMerge.Execute(new[] { "note" }, new[] { Note });
                        baoCao.MailMerge.Execute(new[] { "saletax" }, new[] { Tax });
                        baoCao.MailMerge.Execute(new[] { "warehouse" }, new[] { NameStock });
                        baoCao.MailMerge.Execute(new[] { "ngaydathang" }, new[] { OrderDate.ToString("dd/MM/yyyy") });
                        baoCao.MailMerge.Execute(new[] { "ngaygiaohang" }, new[] { ShipDate.ToString("dd/MM/yyyy") });
                        baoCao.MailMerge.Execute(new[] { "subtotal" }, new[] { UntaxedString });
                        baoCao.MailMerge.Execute(new[] { "tax" }, new[] { TaxString });
                        baoCao.MailMerge.Execute(new[] { "total" }, new[] { TotalString });

                        //điền thông tin vào bảng
                        Table bangSanPhamMua = baoCao.GetChild(NodeType.Table, 0, true) as Table;
                        int hangHienTai = 1;
                        bangSanPhamMua.InsertRows(hangHienTai, hangHienTai, POExportDatagrid.Count - 1);
                        for (int i = 0; i < POExportDatagrid.Count; i++)
                        {
                            bangSanPhamMua.PutValue(hangHienTai, 0, POExportDatagrid[i].POIdDataGrid.ToString());
                            bangSanPhamMua.PutValue(hangHienTai, 1, POExportDatagrid[i].POItemCodeDataGrid);
                            bangSanPhamMua.PutValue(hangHienTai, 2, POExportDatagrid[i].POItemNameDataGrid);
                            bangSanPhamMua.PutValue(hangHienTai, 3, POExportDatagrid[i].POUoMDataGrid);
                            bangSanPhamMua.PutValue(hangHienTai, 4, POExportDatagrid[i].POItemQuantityDataGrid);
                            bangSanPhamMua.PutValue(hangHienTai, 5, POExportDatagrid[i].POItemPriceDataGrid);
                            bangSanPhamMua.PutValue(hangHienTai, 6, POExportDatagrid[i].POItemSumDataGrid);
                            hangHienTai++;
                        }
                        baoCao.SaveAndOpenFile("PO" + NO + ".docx");
                    }
                });

            NewPO = new RelayCommand<object>(
                p =>
                {
                    return true;
                },
                (p) =>
                {
                    Address1 = null;
                    Address2 = null;
                    PIC = null;
                    NameStock = null;
                    TaxString = null;
                    UntaxedString = null;
                    TotalString = null;
                    POExportDatagrid = new ObservableCollection<PODataGrid>();
                    Tax = null;
                    OrderDate = DateTime.Today;
                    ShipDate = DateTime.Today.AddDays(7);
                    Note = null;
                    NO = random.NextString(9);
                    Giatrisauthue = null;
                    Thue = null;
                    Giatritruocthue = null;
                    ListProduct = new ObservableCollection<Product>();
                    ListSupplier = new ObservableCollection<Supplier>();
                    ListSupplier = new ObservableCollection<Supplier>(DataProvider.Ins.DB.Suppliers);
                });

            TotalPO = new RelayCommand<object>(
                p =>
                {
                    if (POExportDatagrid != null && Tax != null)
                        return true;
                    else
                        return false;
                },
                (p) =>
                {
                    Giatritruocthue = "Giá trị trước thuế/ Untaxed Amount:";
                    Thue = "Giá trị thuế/ Tax:";
                    Giatrisauthue = "Tổng tiền/ Total:";
                    for (int i = 0; i < POExportDatagrid.Count; i++)
                    {
                        UntaxedString = (Sum += Double.Parse(POExportDatagrid[i].POItemSumDataGrid)).ToString("0,0,0");
                    }
                    TaxString = (Double.Parse(UntaxedString) * (Double.Parse(Tax) / 100)).ToString("0,0,0");
                    TotalString = (Double.Parse(UntaxedString) + Double.Parse(TaxString)).ToString("0,0,0");
                    Sum = 0;
                });
        }
    }
    public class PODataGrid : BaseViewModel
    {
        public int _POIdDataGrid;
        public int POIdDataGrid { get => _POIdDataGrid; set { _POIdDataGrid = value; OnPropertyChanged(); } }
        public string _POItemCodeDataGrid;
        public string POItemCodeDataGrid { get => _POItemCodeDataGrid; set { _POItemCodeDataGrid = value; OnPropertyChanged(); } }

        public string _POItemNameDataGrid;
        public string POItemNameDataGrid { get => _POItemNameDataGrid; set { _POItemNameDataGrid = value; OnPropertyChanged(); } }
        public string _POItemQuantityDataGrid;
        public string POItemQuantityDataGrid { get => _POItemQuantityDataGrid; set { _POItemQuantityDataGrid = value; OnPropertyChanged(); } }
        public string _POUoMDataGrid;
        public string POUoMDataGrid { get => _POUoMDataGrid; set { _POUoMDataGrid = value; OnPropertyChanged(); } }
        public string _POItemPriceDataGrid;
        public string POItemPriceDataGrid { get => _POItemPriceDataGrid; set { _POItemPriceDataGrid = value; OnPropertyChanged(); } }
        public string _POItemSumDataGrid;
        public string POItemSumDataGrid { get => _POItemSumDataGrid; set { _POItemSumDataGrid = value; OnPropertyChanged(); } }
        public PODataGrid(int pOIdDataGrid, string pOItemCodeDataGrid, string pOItemNameDataGrid, string pOItemQuantityDataGrid, string pOUoMDataGrid, string pOItemPriceDataGrid, string pOItemSumDataGrid)
        {
            POIdDataGrid = pOIdDataGrid;
            POItemCodeDataGrid = pOItemCodeDataGrid;
            POItemNameDataGrid = pOItemNameDataGrid;
            POItemQuantityDataGrid = pOItemQuantityDataGrid;
            POUoMDataGrid = pOUoMDataGrid;
            POItemPriceDataGrid = pOItemPriceDataGrid;
            POItemSumDataGrid = pOItemSumDataGrid;
        }
    }
}