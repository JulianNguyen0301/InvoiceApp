import tkinter 
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox,FLAT
from PIL import Image, ImageTk
from tkcalendar import Calendar
from datetime import datetime
from datetime import date
import os

class PurchaseOrderApp:
    def __init__(self,window):
        self.window = window
        self.window.geometry("1366x635")
        self.window.iconbitmap("H:\logobkra_cCj_1.ico")
        self.window.title("Purchase Order Generator Form")
        self.window.resizable(0,0)
        # Initialize product data

        self.products_data = {
            "Sprite Plastic 300ml - Sprite chai nhua 300ml": {"code": "Coca 1", "price": 2702.42, "uom": "Bottle"},
            "Dasani Water": {"code": "Coca 2", "price": 3165.71, "uom": "Bottle"},
            "Fanta Plastic 300ml - Fanta chai nhua 300ml": {"code": "Coca 3", "price": 2702.42, "uom": "Bottle"},
            "Coke Plastic 300ml - Coca chai nhua 300ml": {"code": "Coca 4", "price": 2702.42, "uom": "Bottle"}
        }
        self.label_font = ("times new roman",13,'bold')
        self.entry_font = ("times new roman",10,'bold')
        self.desc_selected = False
        self.dem = 1
        self.po_list = []
        self.po_list_gui = []
        self.selected_item_id = None

        self.center_window(1366, 635) 
        self.main_frame()
        self.information1_frame()
        self.information2_frame()
        self.create_button()
        self.create_table()
        self.selection_products_frame()

    def main_frame(self):
        self.window.configure(bg='white')
        self.frame = tkinter.Frame(self.window, bg='#00BFFF', bd=10)
        self.frame.pack(fill='both', expand=True)
        self.style = ttk.Style()
        self.style.configure('Custom.TLabel', background='white', foreground='#00BFFF', borderwidth=5, relief='flat')
        self.label = ttk.Label(self.frame, style='Custom.TLabel')
        self.label.pack(fill='both', expand=True)

    def center_window(self, width, height):
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2 - 40
        self.window.geometry(f'{width}x{height}+{x}+{y}')
#Information1
    def information1_frame(self):
        infor_frame = tkinter.Frame(self.window, bg="#00BFFF", width=880, height=200)
        infor_frame.place(x=15, y=15)

        infor_lable = tkinter.Label(infor_frame, text = "Information Details",font=self.label_font)
        infor_lable.place(x = 380, y = 0)
        infor_lable.configure(foreground="#FF3333", background="#00BFFF")


        self.vendor_label = tkinter.Label(infor_frame, text = "Nhà cung cấp/ Vendor:",font=self.label_font)
        self.vendor_label.place(x = 0, y = 21)
        self.vendor_label.configure(foreground="black", background="#00BFFF")

        self.vendor_entry = tkinter.Entry(infor_frame)
        self.vendor_entry.place(x = 220, y = 23,width= 655, height= 23)
        self.vendor_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

        self.address1_label = tkinter.Label(infor_frame, text = "Địa chỉ/ Address:",font=self.label_font)
        self.address1_label.place(x = 0, y = 51)
        self.address1_label.configure(foreground="black", background="#00BFFF")

        self.address1_entry = tkinter.Entry(infor_frame)
        self.address1_entry.place(x = 220, y = 53,width= 655, height= 23)
        self.address1_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


        self.pic_label = tkinter.Label(infor_frame, text = "Người phụ trách/ PIC:",font=self.label_font)
        self.pic_label.place(x = 0, y = 81)
        self.pic_label.configure(foreground="black", background="#00BFFF")

        self.pic_entry = tkinter.Entry(infor_frame)
        self.pic_entry.place(x = 220, y = 83,width= 655, height= 23)
        self.pic_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


        self.warehouse_label = tkinter.Label(infor_frame, text = "Kho hàng tại/ Warehouse at:",font=self.label_font)
        self.warehouse_label.place(x = 0, y = 111)
        self.warehouse_label.configure(foreground="black", background="#00BFFF")

        self.warehouse_entry = tkinter.Entry(infor_frame)
        self.warehouse_entry.place(x = 220, y = 113,width= 655, height= 23)
        self.warehouse_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


        self.address2_label = tkinter.Label(infor_frame, text = "Địa chỉ/ Address:",font=self.label_font)
        self.address2_label.place(x = 0, y = 141)
        self.address2_label.configure(foreground="#000000",background="#00BFFF")

        self.address2_entry = tkinter.Entry(infor_frame)
        self.address2_entry.place(x = 220, y = 143,width= 655, height= 23)
        self.address2_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


        self.notes_label = tkinter.Label(infor_frame, text = "Ghi chú/ Notes:",font=self.label_font)
        self.notes_label.place(x = 0, y = 171)
        self.notes_label.configure(foreground="#000000",background="#00BFFF")

        self.notes_entry = tkinter.Entry(infor_frame)
        self.notes_entry.place(x = 220, y = 173,width= 655, height= 23)
        self.notes_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")
#-----------------------------
#Information2
    def show_date1_window(self):
        self.date1.deiconify()

    def dathang_date(self,event=None):
        selected_date = self.cal1.get_date()
        self.ngaydathang_entry.delete(0, tkinter.END)
        self.ngaydathang_entry.insert(0, selected_date)
        self.date1.withdraw() 

    def show_date2_window(self):
        self.date2.deiconify()

    def giaohang_date(self,event=None):
        selected_date = self.cal2.get_date()
        self.ngaygiaohang_entry.delete(0, tkinter.END)
        self.ngaygiaohang_entry.insert(0, selected_date)
        self.date2.withdraw() 

    def information2_frame(self):
        infor_frame2 = tkinter.Frame(self.window, bg="#00BFFF", width=250, height=120)
        infor_frame2.place(x=910, y=15)

        self.NO_label = tkinter.Label(infor_frame2, text = "Số phiếu / NO:",font=self.label_font)
        self.NO_label.place(x = 3, y = 3)
        self.NO_label.configure(foreground="#000000",background="#00BFFF")

        self.NO_entry = tkinter.Entry(infor_frame2,justify="center")
        self.NO_entry.place(x = 127, y = 3,width= 90, height= 23)
        self.NO_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")
        self.NO_entry.insert(0, "PO")

        self.tax_label = tkinter.Label(infor_frame2, text = "Thuế / TAX:",font=self.label_font)
        self.tax_label.place(x = 3, y = 33)
        self.tax_label.configure(foreground="#000000",background="#00BFFF")
        self.percent_label = tkinter.Label(infor_frame2, text = "(%)",font=self.label_font)
        self.percent_label.place(x = 216, y = 31)
        self.percent_label.configure(foreground="#000000",background="#00BFFF")

        self.tax_entry = tkinter.Entry(infor_frame2,justify="center")
        self.tax_entry.place(x = 127, y = 33,width= 90, height= 23)
        self.tax_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")
        self.tax_entry.insert(0, "8")

        self.calendar_icon = Image.open("H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Customize PDF\date.png")
        self.calendar_icon = self.calendar_icon.resize((19, 19))
        self.calendar_icon = ImageTk.PhotoImage(self.calendar_icon)
        current_date = date.today()

        self.ngaydathang_label = tkinter.Label(infor_frame2, text="Ngày đặt hàng:", font=self.label_font)
        self.ngaydathang_label.place(x = 3, y = 63)
        self.ngaydathang_label.configure(foreground="#000000",background="#00BFFF")

        self.ngaydathang_entry = tkinter.Entry(infor_frame2,justify="center")
        self.ngaydathang_entry.place(x = 127, y = 63,width= 90, height= 23)
        self.ngaydathang_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

        self.ngaydathang_button = tkinter.Button(infor_frame2, image=self.calendar_icon, font=("Arial", 12), command = self.show_date1_window, cursor="hand2")
        self.ngaydathang_button.place(x = 222, y = 63)
        self.ngaydathang_button.configure(foreground="#00BFFF",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        self.date1 = tkinter.Tk()
        self.date1.title("Date - đặt hàng")
        self.date1.withdraw()
        self.date1.overrideredirect(1)

        self.date1_var = tkinter.StringVar()
        self.date1_var.set(current_date.strftime("%d/%m/%Y"))
        self.cal1 = Calendar(self.date1, selectmode="day", date_var=self.date1_var,date_pattern='dd/mm/yyyy')
        self.cal1.pack(pady=20, fill='both', expand=True)
        self.cal1.bind("<<CalendarSelected>>", self.dathang_date)

        ngaygiaohang_label = tkinter.Label(infor_frame2, text = "Ngày giao hàng:",font=self.label_font)
        ngaygiaohang_label.place(x = 3, y = 93)
        ngaygiaohang_label.configure(foreground="#000000",background="#00BFFF")

        self.ngaygiaohang_entry = tkinter.Entry(infor_frame2,justify="center")
        self.ngaygiaohang_entry.place(x = 127, y = 93,width= 90, height= 23)
        self.ngaygiaohang_entry.configure(font=self.entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

        ngaygiaohang_button = tkinter.Button(infor_frame2, image = self.calendar_icon, font=("Arial", 12), command= self.show_date2_window, cursor="hand2")
        ngaygiaohang_button.place(x = 222, y = 93)
        ngaygiaohang_button.configure(foreground="#00BFFF",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        self.date2 = tkinter.Tk()
        self.date2.title("Date - giao hàng")
        self.date2.withdraw()
        self.date2.overrideredirect(1)

        date2_var = tkinter.StringVar()
        date2_var.set(current_date.strftime("%d/%m/%Y"))
        self.cal2 = Calendar(self.date2, selectmode="day", date_var=date2_var,date_pattern='dd/mm/yyyy')
        self.cal2.pack(pady=20, fill='both', expand=True)
        self.cal2.bind("<<CalendarSelected>>", self.giaohang_date)
#----------------------------------
#Selection_Frame
    def on_product_name_select1(self,event):
        selected_product_name = self.products_name_combobox.get()
        corresponding_product_code = self.products_code_dict.get(selected_product_name, "")
        self.products_code_var.set(corresponding_product_code)
        
        # Cập nhật đơn giá dựa trên tên sản phẩm được chọn
        selected_product_price = self.product_prices.get(selected_product_name, "0.00")
        self.dg_var.set(selected_product_price)
        selected_product_price = self.product_dvt.get(selected_product_name,"")
        self.dvt_var.set(selected_product_price)

    def on_desc_selected(self,event):
        global desc_selected
        if self.products_name_combobox.get() != "":
            desc_selected = True
        else:
            desc_selected = False

    def selection_products_frame(self):
        seletion_frame = tkinter.Frame(self.window, bg="#00BFFF", width=1335, height=60)
        seletion_frame.place(x=15, y=220)

        products_name_label = tkinter.Label(seletion_frame, text = "Tên sản phẩm/ Product Name",font=self.label_font)
        products_name_label.place(x = 105, y = 3)
        products_name_label.configure(foreground="#000000",background="#00BFFF")

        self.products_name = ["Sprite Plastic 300ml - Sprite chai nhua 300ml", 
                        "Dasani Water", 
                        "Fanta Plastic 300ml - Fanta chai nhua 300ml",
                        "Coke Plastic 300ml - Coca chai nhua 300ml"]
        self.products_name_combobox = ttk.Combobox(seletion_frame, values=self.products_name)
        self.products_name_combobox.place(x = 3, y = 27 ,width= 450, height= 26)
        self.products_name_combobox.configure(font=self.label_font)
        self.products_name_combobox.option_add("*TCombobox*Listbox.font", self.label_font)
        # products_name_combobox.bind("<<ComboboxSelected>>", on_product_name_select1)
        # products_name_combobox.bind("<<ComboboxSelected>>", on_desc_selected)
        self.products_name_combobox.bind("<<ComboboxSelected>>", lambda event: (self.on_product_name_select1(event), self.on_desc_selected(event)))
        # # Xử lý lựa chọn mã hàng ở đây
        product_code_label = tkinter.Label(seletion_frame, text = "Mã hàng/ Product Code",font=self.label_font)
        product_code_label.place(x = 505, y = 3)
        product_code_label.configure(foreground="#000000",background="#00BFFF")

        self.products_code_dict = { "Sprite Plastic 300ml - Sprite chai nhua 300ml": "Coca 1",
                                "Dasani Water": "Coca 2",
                                "Fanta Plastic 300ml - Fanta chai nhua 300ml": "Coca 3",
                                "Coke Plastic 300ml - Coca chai nhua 300ml": "Coca 4"}
        self.products_code_var = tkinter.StringVar()

        self.products_code_combobox = ttk.Combobox(seletion_frame, textvariable=self.products_code_var, state="normal",justify="center")
        self.products_code_combobox.place(x = 510, y = 27,width= 170, height= 26)
        self.products_code_combobox.configure(font=self.label_font)

        dg_label = tkinter.Label(seletion_frame, text = "Đơn giá/ Price Unit",font=self.label_font)
        dg_label.place(x = 765, y = 3)
        dg_label.configure(foreground="#000000",background="#00BFFF")

        self.product_prices = {
            "Sprite Plastic 300ml - Sprite chai nhua 300ml": "2,702.42",
            "Dasani Water": "3,165.71",
            "Fanta Plastic 300ml - Fanta chai nhua 300ml": "2,702.42",  
            "Coke Plastic 300ml - Coca chai nhua 300ml": "2,702.42" }
        self.dg_var = tkinter.StringVar()
        self.dg_combobox = ttk.Combobox(seletion_frame, textvariable=self.dg_var, state="normal",justify="center")
        self.dg_combobox.configure(font=self.label_font)
        self.dg_combobox.place(x = 755, y = 28,width= 170, height= 26)

        dvt_label = tkinter.Label(seletion_frame, text = "Đơn vị tính/ UoM",font=self.label_font)
        dvt_label.place(x = 994, y = 3)
        dvt_label.configure(foreground="#000000",background="#00BFFF")
        self.product_dvt = {
            "Sprite Plastic 300ml - Sprite chai nhua 300ml": "Bottle",
            "Dasani Water": "Bottle",
            "Fanta Plastic 300ml - Fanta chai nhua 300ml": "Bottle",  
            "Coke Plastic 300ml - Coca chai nhua 300ml": "Bottle" }
        self.dvt_var = tkinter.StringVar()
        self.dvt_combobox = ttk.Combobox(seletion_frame,textvariable=self.dvt_var,state="normal",justify="center")
        self.dvt_combobox.place(x = 1000, y = 28,width= 120, height= 26)
        self.dvt_combobox.configure(font=self.label_font)

        quantity_label = tkinter.Label(seletion_frame, text="Số lượng/ Quantity",font=self.label_font)
        quantity_label.place(x = 1178, y = 3)
        quantity_label.configure(foreground="#000000",background="#00BFFF")
        self.quantity_spinbox = tkinter.Spinbox(seletion_frame,from_= 1, to = 1000,justify="center")
        self.quantity_spinbox.configure(font=self.label_font)
        self.quantity_spinbox.place(x = 1190, y = 28,width= 120, height= 26)
#----------------------------------------
#Table
    def select_tree_item(self,event):
        # Xóa dữ liệu cũ trong các widget
        self.products_name_combobox.set("")
        self.dvt_combobox.set("")
        self.quantity_spinbox.delete(0, tkinter.END)
        self.products_code_combobox.delete(0, tkinter.END)
        self.dg_combobox.delete(0, tkinter.END)
        # Lấy thông tin của hàng được chọn
        selected_item = self.tree.selection()
        if selected_item:
            self.selected_item_id = selected_item[0]
            values = self.tree.item(selected_item, "values")
            self.products_name_combobox.set(values[2])
            self.dvt_combobox.set(values[4])
            self.quantity_spinbox.insert(0, values[3])
            self.products_code_combobox.insert(0, values[1])
            self.dg_combobox.insert(0, values[5])
            self.update_button["state"] = "normal"
            self.delete_button["state"] = "normal"
        else:
            self.selected_item_id = None
            self.delete_button["state"] = "disabled"

    def create_table(self):
        columns = ( 'stt','code','desc','qty', 'UoM' ,'price', 'total',)

        self.tree = ttk.Treeview(self.window, columns=columns, show="headings")
        self.tree.column('stt', width=50,anchor='center')
        self.tree.column('code', width=130,anchor='center')
        self.tree.column('qty', width=90,anchor='center')
        self.tree.column('desc', width=500)
        self.tree.column('UoM', width=95,anchor='center')
        self.tree.column('price', width=105,anchor='center')
        self.tree.column('total', width=100,anchor='e')

        self.tree.heading('stt', text='STT')
        self.tree.heading('code', text='Product Code')
        self.tree.heading('qty', text='Quantity')
        self.tree.heading('desc', text='Description')
        self.tree.heading('UoM', text='UoM')
        self.tree.heading('price', text='Unit Price')
        self.tree.heading('total', text="Total")


        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Times new roman", 13,'bold'), foreground = "black", background = "black")

        self.tree.place(x=265,y=300)
        self.tree.bind("<Button-1>", self.select_tree_item)
        self.tree.bind("<<TreeviewSelect>>", self.select_tree_item)
#-----------------------------------------
#Button
    def clear_item(self):
        self.quantity_spinbox.delete(0, tkinter.END)
        self.quantity_spinbox.insert(0, "1")
        self.products_name_combobox.delete(0, tkinter.END)
        self.dg_combobox.delete(0, tkinter.END)
        self.dg_combobox.insert(0, "")
        self.dvt_combobox.delete(0, tkinter.END)
        self.products_code_combobox.delete(0, tkinter.END)
        global desc_selected
        desc_selected = False

    def add_item(self):
        global desc_selected
        if not desc_selected:
            # Hiển thị thông báo hoặc thông báo lỗi khi "desc" chưa được chọn
            messagebox.showerror("Lỗi", "Vui lòng chọn sản phẩm")
            return  # Dừng hàm và không thêm mục nếu "desc" chưa được chọn
        UoM = self.dvt_combobox.get()
        qty = int(self.quantity_spinbox.get())
        desc = self.products_name_combobox.get()
        code_pro = self.products_code_combobox.get()
        price_str = self.dg_combobox.get()
        price_str1 = price_str.replace(',', '')
        price = float(price_str1)
        line_total = int(qty*price)
        formatted_line_total = '{:,}'.format(line_total)
        po_item_gui = [self.dem,code_pro,desc,qty,UoM, price_str, formatted_line_total]
        po_item = [self.dem,code_pro,desc,qty,UoM, price_str, line_total]
        self.tree.insert('',"end", values=po_item_gui)
        self.clear_item()
        self.dem += 1
        self.po_list_gui.append(po_item_gui)
        self.po_list.append(po_item)
        print(self.po_list_gui)

    def modify_item(self):
        selected_item = self.tree.selection()
        if self.selected_item_id:
            # Lấy dữ liệu mới từ các điều khiển
            self.selected_item_id = selected_item[0]
            stt = self.selected_item_id[-1]
            index = int(stt) - 1

            new_products_name = self.products_name_combobox.get()
            new_dvt = self.dvt_combobox.get()
            new_quantity = int(self.quantity_spinbox.get())
            new_products_code = self.products_code_combobox.get()
            new_dg = self.dg_combobox.get()
            new_dg1 = new_dg.replace(',', '')
            new_price = float(new_dg1)
            new_line_total = int(new_price*new_quantity)
            new_formatted_line_total = '{:,}'.format(new_line_total)
            # Cập nhật dữ liệu trên hàng đã chọn trong ttk.Treeview
            self.tree.item(self.selected_item_id, values=(stt, new_products_code, new_products_name, new_quantity, new_dvt, new_dg, new_formatted_line_total))
            if 0 <= index <= len(self.po_list_gui):
                self.po_list_gui[index] = [stt, new_products_code, new_products_name, new_quantity, new_dvt, new_dg, new_formatted_line_total]
                self.po_list[index] = [stt, new_products_code, new_products_name, new_quantity, new_dvt, new_dg, new_line_total]
            print(self.po_list_gui)
            # Đặt lại trạng thái của nút "Modify"
            self.update_button["state"] = "disabled"
            self.quantity_spinbox.delete(0, tkinter.END)
            self.quantity_spinbox.insert(0, "1")
            self.products_name_combobox.delete(0, tkinter.END)
            self.dg_combobox.delete(0, tkinter.END)
            self.dg_combobox.insert(0, "")
            self.dvt_combobox.delete(0, tkinter.END)
            self.products_code_combobox.delete(0, tkinter.END)    

    def update_stt(self):
        children = self.tree.get_children()
        for index, child in enumerate(children, start=1):
            self.tree.item(child, values=(index,) + self.tree.item(child, "values")[1:])
            stt = self.tree.item(child, "values")[0]
            item_index = int(stt) - 1
            if 0 <= item_index < len(self.po_list):
                self.po_item = self.po_list[item_index]
                self.po_item_gui = self.po_list_gui[item_index]
                # Update other values in the same row as needed
                # For example, if you want to update the 'desc' value:
                self.tree.item(child, values=(index, self.po_item_gui[1], self.po_item_gui[2], self.po_item_gui[3], self.po_item_gui[4], self.po_item_gui[5], self.po_item_gui[6]))

    def delete_item(self):
        selected_item = self.tree.selection()
        if self.selected_item_id:
            # Get the selected item's stt value
            self.selected_item_id = selected_item[0]
            stt = self.selected_item_id[-1]
            # Find the index of the item in the list
            index = int(stt) - 1  # The list index is one less than the stt
            # Remove the item from the lists
            if 0 <= index < len(self.po_list):
                del self.po_list[index]
                del self.po_list_gui[index]

            for i, item in enumerate(self.po_list_gui):
                item[0] = i + 1 
            print(self.po_list_gui)
            # Delete the item from the tree view
            self.tree.delete(self.selected_item_id)
            self.selected_item_id = None
            self.update_stt()
            self.quantity_spinbox.delete(0, tkinter.END)
            self.quantity_spinbox.insert(0, "1")
            self.products_name_combobox.delete(0, tkinter.END)
            self.dg_combobox.delete(0, tkinter.END)
            self.dg_combobox.insert(0, "")
            self.dvt_combobox.delete(0, tkinter.END)
            self.products_code_combobox.delete(0, tkinter.END)

    def new_po(self):
        self.vendor_entry.delete(0, tkinter.END)
        self.address1_entry.delete(0, tkinter.END)
        self.pic_entry.delete(0, tkinter.END)
        self.warehouse_entry.delete(0, tkinter.END)
        self.address2_entry.delete(0, tkinter.END)
        self.notes_entry.delete(0, tkinter.END)
        self.NO_entry.delete(0, tkinter.END)
        self.NO_entry.insert(0,"PO")
        self.ngaydathang_entry.delete(0, tkinter.END)
        self.ngaygiaohang_entry.delete(0, tkinter.END)
        self.clear_item()
        self.tree.delete(*self.tree.get_children())
        # global dem 
        # dem = 1
        self.total_frame.place_forget() 
        self.po_list.clear()
        self.po_list_gui.clear()

    def generate_po(self):
        doc = DocxTemplate("H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Customize PDF\po_template.docx")
        vendor = self.vendor_entry.get()
        address1 = self.address1_entry.get()
        pic = self.pic_entry.get()
        warehouse = self.warehouse_entry.get()
        address2 = self.address2_entry.get()
        notes = self.notes_entry.get()
        no = self.NO_entry.get()
        current_date = datetime.today()
        current_date1 = current_date.strftime("%d/%m/%Y %H:%M:%S")
        #columns = ( 'stt','code','desc','qty', 'UoM' ,'price', 'total',)
        subtotal = sum(item[6] for item in self.po_list) 
        formatted_subtotal = '{:,}'.format(subtotal)
        ngaydathang = self.ngaydathang_entry.get()
        ngaygiaohang = self.ngaygiaohang_entry.get()
        #Thêm salestax1 vào screen
        salestax1 = float(self.tax_entry.get())/100
        salestax2 = int(salestax1*subtotal)
        formatted_salestax2 = '{:,}'.format(salestax2)
        total = int(subtotal + subtotal*salestax1)
        formatted_total = '{:,}'.format(total)
        if no != "":
            doc.render({"vendor":vendor, 
                    "address1":address1,
                    "pic":pic,
                    "warehouse":warehouse,
                    "address2":address2,
                    "notes":notes,
                    "no":no,
                    "po_list": self.po_list_gui,
                    "subtotal":formatted_subtotal,
                    "salestax1": int(salestax1*100),
                    "salestax2":formatted_salestax2,
                    "total":formatted_total,
                    "ngaydathang": ngaydathang,
                    "ngaygiaohang": ngaygiaohang,
                    "current_date": current_date1})
            output_directory = r"H:/APP UNIVERSITY/CODE PYTHON/OpenCV-Master-Computer-Vision-in-Python/SourcecodeOCR/Customize PDF/word_po"
            doc_name = os.path.join(output_directory, no + ".docx")        
            doc.save(doc_name)
    
            messagebox.showinfo("Purchase Order Complete", "Purchase Order Complete")
        else:
            messagebox.showinfo("Purchase Order Complete", "Cần phải nhập Số phiếu / NO")

    def total_po(self):
        if not self.tree.get_children():  # Kiểm tra xem có bất kỳ item nào trong Treeview không
            messagebox.showwarning("Thông báo", "Hãy nhập ít nhất một sản phẩm.")
        elif self.tax_entry.get() == "":
            messagebox.showwarning("Thông báo", "Hãy nhập mục thuế GTGT")
        else:
            self.total_frame.place(x=928, y=530)
            untax_lable = tkinter.Label(self.total_frame, text = "Giá trị trước thuế/ Untaxed Amount:",font=self.label_font)
            untax_lable.place(x = 0, y = 0)
            untax_lable.configure(foreground="black", background="#FFFF33")

            tax_lable = tkinter.Label(self.total_frame, text = "Giá trị thuế/ Tax:",font=self.label_font)
            tax_lable.place(x = 140, y = 30)
            tax_lable.configure(foreground="black", background="#FFFF33")

            total_lable = tkinter.Label(self.total_frame, text = "Tổng tiền/ Total:",font=self.label_font)
            total_lable.place(x = 145, y = 60)
            total_lable.configure(foreground="black", background="#FFFF33")

            subtotal = sum(item[6] for item in self.po_list) 
            formatted_subtotal = '{:,}'.format(subtotal)
            subtotal_value_label = tkinter.Label(self.total_frame, text=formatted_subtotal, font=self.label_font)
            subtotal_value_label.place(x=305, y=0)
            subtotal_value_label.configure(foreground="black", background="#FFFF33")

            salestax1 = float(self.tax_entry.get())/100
            salestax2 = int(salestax1*subtotal)
            formatted_salestax2 = '{:,}'.format(salestax2)
            tax_value_label = tkinter.Label(self.total_frame, text=formatted_salestax2, font=self.label_font)
            tax_value_label.place(x=305, y=30)
            tax_value_label.configure(foreground="black", background="#FFFF33")

            total = int(subtotal + subtotal*salestax1)
            formatted_total = '{:,}'.format(total)
            total_value_label = tkinter.Label(self.total_frame, text=formatted_total, font=self.label_font)
            total_value_label.place(x=305, y=60)
            total_value_label.configure(foreground="black", background="#FFFF33")

    def close_window(self):
            self.window.quit() 

    def on_closing(self):
        messagebox.showinfo("Warning","Nếu bạn muốn thoát chương trình hãy nhấn (Turn Off Application)")
    
    def create_button(self):
        add_item_button = tkinter.Button(self.window, text = "Add item", command= self.add_item,font=self.label_font)
        add_item_button.place(x = 910, y = 160, width=80, height=40)
        add_item_button.configure(foreground="black",background="#4ED70F",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        self.update_button = tkinter.Button(self.window, text="Modify", command= self.modify_item, font=self.label_font)
        self.update_button.place(x = 1030, y = 160, width=80, height=40)
        self.update_button.configure(foreground="black",background="#FFFF33",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        self.delete_button = tkinter.Button(self.window, text="Delete", command=self.delete_item, font=self.label_font)
        self.delete_button.place(x = 1150, y = 160, width=80, height=40)
        self.delete_button.configure(foreground="black",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        save_po_button = tkinter.Button(self.window, text="Generate Purchase Order",command=self.generate_po,font=self.label_font)
        save_po_button.place(x = 23, y = 300, width=230, height=40) 
        save_po_button.configure(foreground="black",background="#00BFFF",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        new_po_button = tkinter.Button(self.window, text="New Purchase Order",command=self.new_po,font=self.label_font)
        new_po_button.place(x = 23, y = 360, width=230, height=40) 
        new_po_button.configure(foreground="black",background="#4ED70F",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")
        
        exit_button = tkinter.Button(self.window, text="Turn Off Application", command = self.close_window,font=self.label_font)
        exit_button.place(x = 23, y = 420, width=230, height=40)
        exit_button.configure(foreground="#ffffff",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

        self.total_button = tkinter.Button(self.window, text="Total", command= self.total_po,font=self.label_font)
        self.total_button.place(x = 23, y = 480, width=230, height=40)
        self.total_button.configure(foreground="black",background="#FFFF33",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")
        self.total_frame = tkinter.Frame(self.window, bg="#FFFF33", width=400, height=90)
#-----------------------
def main():
    window = tkinter.Tk()
    app = PurchaseOrderApp(window)
    window.protocol("WM_DELETE_WINDOW", app.on_closing)
    window.mainloop()
if __name__ == "__main__":
    main()