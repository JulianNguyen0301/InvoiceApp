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

def clear_item():
    quantity_spinbox.delete(0, tkinter.END)
    quantity_spinbox.insert(0, "1")
    products_name_combobox.delete(0, tkinter.END)
    dg_combobox.delete(0, tkinter.END)
    dg_combobox.insert(0, "")
    dvt_combobox.delete(0, tkinter.END)
    products_code_combobox.delete(0, tkinter.END)
    global desc_selected
    desc_selected = False

desc_selected = False
def on_desc_selected(event):
    global desc_selected
    if products_name_combobox.get() != "":
        desc_selected = True
    else:
        desc_selected = False

po_list = []
po_list_gui = []
dem = 1
def add_item():
    global dem, desc_selected
    if not desc_selected:
        # Hiển thị thông báo hoặc thông báo lỗi khi "desc" chưa được chọn
        messagebox.showerror("Lỗi", "Vui lòng chọn sản phẩm")
        return  # Dừng hàm và không thêm mục nếu "desc" chưa được chọn

    UoM = dvt_combobox.get()
    qty = int(quantity_spinbox.get())
    desc = products_name_combobox.get()
    code_pro = products_code_combobox.get()
    price_str = dg_combobox.get()
    price_str1 = price_str.replace(',', '')
    price = float(price_str1)
    line_total = int(qty*price)
    formatted_line_total = '{:,}'.format(line_total)
    po_item_gui = [dem,code_pro,desc,qty,UoM, price_str, formatted_line_total]
    po_item = [dem,code_pro,desc,qty,UoM, price_str, line_total]
    tree.insert('',"end", values=po_item_gui)
    clear_item()
    dem += 1
    po_list_gui.append(po_item_gui)
    po_list.append(po_item)
    print(po_list_gui)
#dem = 1

def new_po():
    vendor_entry.delete(0, tkinter.END)
    address1_entry.delete(0, tkinter.END)
    pic_entry.delete(0, tkinter.END)
    warehouse_entry.delete(0, tkinter.END)
    address2_entry.delete(0, tkinter.END)
    notes_entry.delete(0, tkinter.END)
    NO_entry.delete(0, tkinter.END)
    NO_entry.insert(0,"PO000")
    ngaydathang_entry.delete(0, tkinter.END)
    ngaygiaohang_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    global dem 
    dem = 1
    total_frame.place_forget() 
    po_list.clear()
    po_list_gui.clear()
    
def generate_po():
    doc = DocxTemplate("H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Customize PDF\po_template.docx")
    vendor = vendor_entry.get()
    address1 = address1_entry.get()
    pic = pic_entry.get()
    warehouse = warehouse_entry.get()
    address2 = address2_entry.get()
    notes = notes_entry.get()
    no = NO_entry.get()
    current_date = datetime.today()
    current_date1 = current_date.strftime("%d/%m/%Y %H:%M:%S")
    #columns = ( 'stt','code','desc','qty', 'UoM' ,'price', 'total',)
    subtotal = sum(item[6] for item in po_list) 
    formatted_subtotal = '{:,}'.format(subtotal)
    ngaydathang = ngaydathang_entry.get()
    ngaygiaohang = ngaygiaohang_entry.get()
    #Thêm salestax1 vào screen
    salestax1 = float(tax_entry.get())/100
    tax_temp = float(salestax1*subtotal)
    salestax2 = int(salestax1*subtotal)
    print(salestax1)
    print(tax_temp)
    print(salestax2)
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
                "po_list": po_list_gui,
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

def total_po():
    if not tree.get_children():  # Kiểm tra xem có bất kỳ item nào trong Treeview không
        messagebox.showwarning("Thông báo", "Hãy nhập ít nhất một sản phẩm.")
    elif tax_entry.get() == "":
        messagebox.showwarning("Thông báo", "Hãy nhập mục thuế GTGT")
    else:
        total_frame.place(x=928, y=530)
        untax_lable = tkinter.Label(total_frame, text = "Giá trị trước thuế/ Untaxed Amount:",font=label_font)
        untax_lable.place(x = 0, y = 0)
        untax_lable.configure(foreground="black", background="#FFFF33")

        tax_lable = tkinter.Label(total_frame, text = "Giá trị thuế/ Tax:",font=label_font)
        tax_lable.place(x = 140, y = 30)
        tax_lable.configure(foreground="black", background="#FFFF33")

        total_lable = tkinter.Label(total_frame, text = "Tổng tiền/ Total:",font=label_font)
        total_lable.place(x = 145, y = 60)
        total_lable.configure(foreground="black", background="#FFFF33")

        subtotal = sum(item[6] for item in po_list) 
        formatted_subtotal = '{:,}'.format(subtotal)
        subtotal_value_label = tkinter.Label(total_frame, text=formatted_subtotal, font=label_font)
        subtotal_value_label.place(x=305, y=0)
        subtotal_value_label.configure(foreground="black", background="#FFFF33")

        salestax1 = float(tax_entry.get())/100
        salestax2 = int(salestax1*subtotal)
        formatted_salestax2 = '{:,}'.format(salestax2)
        tax_value_label = tkinter.Label(total_frame, text=formatted_salestax2, font=label_font)
        tax_value_label.place(x=305, y=30)
        tax_value_label.configure(foreground="black", background="#FFFF33")

        total = int(subtotal + subtotal*salestax1)
        formatted_total = '{:,}'.format(total)
        total_value_label = tkinter.Label(total_frame, text=formatted_total, font=label_font)
        total_value_label.place(x=305, y=60)
        total_value_label.configure(foreground="black", background="#FFFF33")

def on_product_name_select1(event):
    selected_product_name = products_name_combobox.get()
    corresponding_product_code = products_code_dict.get(selected_product_name, "")
    products_code_var.set(corresponding_product_code)
    
    # Cập nhật đơn giá dựa trên tên sản phẩm được chọn
    selected_product_price = product_prices.get(selected_product_name, "0.00")
    dg_var.set(selected_product_price)
    selected_product_price = product_dvt.get(selected_product_name,"")
    dvt_var.set(selected_product_price)

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2 - 40

    window.geometry(f'{width}x{height}+{x}+{y}')

def increase_font_size(font, size_increase):
    font_family, font_size, font_weight = font.split()
    new_font_size = int(font_size) + size_increase
    return f"{font_family} {new_font_size} {font_weight}"

def dathang_date(event=None):
    selected_date = cal1.get_date()
    ngaydathang_entry.delete(0, tkinter.END)
    ngaydathang_entry.insert(0, selected_date)
    date1.withdraw() 

def show_date1_window():
    date1.deiconify()

def giaohang_date(event=None):
    selected_date = cal2.get_date()
    ngaygiaohang_entry.delete(0, tkinter.END)
    ngaygiaohang_entry.insert(0, selected_date)
    date2.withdraw() 

def show_date2_window():
    date2.deiconify()

def close_window():
    window.quit() 

def on_closing():
    messagebox.showinfo("Warning","Nếu bạn muốn thoát chương trình hãy nhấn (Turn Off Application)")

selected_item_id = None
def select_tree_item(event):
    global selected_item_id
    # Xóa dữ liệu cũ trong các widget
    products_name_combobox.set("")
    dvt_combobox.set("")
    quantity_spinbox.delete(1, tkinter.END) #0
    products_code_combobox.delete(0, tkinter.END)
    dg_combobox.delete(0, tkinter.END)
    # Lấy thông tin của hàng được chọn
    selected_item = tree.selection()
    if selected_item:
        selected_item_id = selected_item[0]
        values = tree.item(selected_item, "values")
        products_name_combobox.set(values[2])
        dvt_combobox.set(values[4])
        quantity_spinbox.insert(0, values[3])
        products_code_combobox.insert(0, values[1])
        dg_combobox.insert(0, values[5])
        update_button["state"] = "normal"
        delete_button["state"] = "normal"
    else:
        selected_item_id = None
        delete_button["state"] = "disabled"

def modify_item():
    global selected_item_id  # Sử dụng biến toàn cục để lấy ID (hoặc index) của hàng đã chọn
    selected_item = tree.selection()
    if selected_item_id:
        # Lấy dữ liệu mới từ các điều khiển
        selected_item_id = selected_item[0]
        stt = selected_item_id[-1]
        index = int(stt) - 1

        new_products_name = products_name_combobox.get()
        new_dvt = dvt_combobox.get()
        new_quantity = int(quantity_spinbox.get())
        new_products_code = products_code_combobox.get()
        new_dg = dg_combobox.get()
        new_dg1 = new_dg.replace(',', '')
        new_price = float(new_dg1)
        new_line_total = int(new_price*new_quantity)
        new_formatted_line_total = '{:,}'.format(new_line_total)
        # Cập nhật dữ liệu trên hàng đã chọn trong ttk.Treeview
        tree.item(selected_item_id, values=(stt, new_products_code, new_products_name, new_quantity, new_dvt, new_dg, new_formatted_line_total))
        if 0 <= index <= len(po_list_gui):
            po_list_gui[index] = [stt, new_products_code, new_products_name, new_quantity, new_dvt, new_dg, new_formatted_line_total]
            po_list[index] = [stt, new_products_code, new_products_name, new_quantity, new_dvt, new_dg, new_line_total]
        print(po_list_gui)
        # Đặt lại trạng thái của nút "Modify"
        update_button["state"] = "disabled"
        quantity_spinbox.delete(0, tkinter.END)
        quantity_spinbox.insert(0, "1")
        products_name_combobox.delete(0, tkinter.END)
        dg_combobox.delete(0, tkinter.END)
        dg_combobox.insert(0, "")
        dvt_combobox.delete(0, tkinter.END)
        products_code_combobox.delete(0, tkinter.END)

def update_stt():
    children = tree.get_children()
    for index, child in enumerate(children, start=1):
        tree.item(child, values=(index,) + tree.item(child, "values")[1:])
        stt = tree.item(child, "values")[0]
        # Update other values in the same row as the item number
        # You can access the data from po_list or po_list_gui based on the stt
        item_index = int(stt) - 1
        if 0 <= item_index < len(po_list):
            po_item = po_list[item_index]
            po_item_gui = po_list_gui[item_index]
            # Update other values in the same row as needed
            # For example, if you want to update the 'desc' value:
            tree.item(child, values=(index, po_item_gui[1], po_item_gui[2], po_item_gui[3], po_item_gui[4], po_item_gui[5], po_item_gui[6]))
    # children = tree.get_children()
    # for index, child in enumerate(children, start=1):
    #     tree.item(child, values=(index,) + tree.item(child, "values")[1:])

def delete_item():
    global selected_item_id
    selected_item = tree.selection()
    if selected_item_id:
        # Get the selected item's stt value
        selected_item_id = selected_item[0]
        stt = selected_item_id[-1]
        
        # Find the index of the item in the list
        index = int(stt) - 1  # The list index is one less than the stt

        # Remove the item from the lists
        if 0 <= index < len(po_list):
            del po_list[index]
            del po_list_gui[index]

        for i, item in enumerate(po_list_gui):
            item[0] = i + 1 
        print(po_list_gui)
        # Delete the item from the tree view
        tree.delete(selected_item_id)
        selected_item_id = None
        update_stt()
        quantity_spinbox.delete(0, tkinter.END)
        quantity_spinbox.insert(0, "1")
        products_name_combobox.delete(0, tkinter.END)
        dg_combobox.delete(0, tkinter.END)
        dg_combobox.insert(0, "")
        dvt_combobox.delete(0, tkinter.END)
        products_code_combobox.delete(0, tkinter.END)
    

        
#Main
window = tkinter.Tk()
window.geometry("1366x635")
window.iconbitmap("H:\logobkra_cCj_1.ico")
window.configure(bg='white')
frame = tkinter.Frame(window, bg='#00BFFF', bd=10)
frame.pack(fill='both', expand=True)
style = ttk.Style()
style.configure('Custom.TLabel', background='white', foreground='#00BFFF', borderwidth=5, relief='flat')
label = ttk.Label(frame, style='Custom.TLabel')
label.pack(fill='both', expand=True)
window.title("Purchase Order Generator Form")
window.resizable(0,0)
center_window(window, 1366, 635)
label_font = ("times new roman",13,'bold')
entry_font = ("times new roman",10,'bold')

#Information

infor_frame = tkinter.Frame(window, bg="#00BFFF", width=880, height=200)
infor_frame.place(x=15, y=15)

infor_lable = tkinter.Label(infor_frame, text = "Information Details",font=label_font)
infor_lable.place(x = 380, y = 0)
infor_lable.configure(foreground="#FF3333", background="#00BFFF")


vendor_label = tkinter.Label(infor_frame, text = "Nhà cung cấp/ Vendor:",font=label_font)
vendor_label.place(x = 0, y = 21)
vendor_label.configure(foreground="black", background="#00BFFF")

vendor_entry = tkinter.Entry(infor_frame)
vendor_entry.place(x = 220, y = 23,width= 655, height= 23)
vendor_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

address1_label = tkinter.Label(infor_frame, text = "Địa chỉ/ Address:",font=label_font)
address1_label.place(x = 0, y = 51)
address1_label.configure(foreground="black", background="#00BFFF")

address1_entry = tkinter.Entry(infor_frame)
address1_entry.place(x = 220, y = 53,width= 655, height= 23)
address1_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


pic_label = tkinter.Label(infor_frame, text = "Người phụ trách/ PIC:",font=label_font)
pic_label.place(x = 0, y = 81)
pic_label.configure(foreground="black", background="#00BFFF")

pic_entry = tkinter.Entry(infor_frame)
pic_entry.place(x = 220, y = 83,width= 655, height= 23)
pic_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


warehouse_label = tkinter.Label(infor_frame, text = "Kho hàng tại/ Warehouse at:",font=label_font)
warehouse_label.place(x = 0, y = 111)
warehouse_label.configure(foreground="black", background="#00BFFF")

warehouse_entry = tkinter.Entry(infor_frame)
warehouse_entry.place(x = 220, y = 113,width= 655, height= 23)
warehouse_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


address2_label = tkinter.Label(infor_frame, text = "Địa chỉ/ Address:",font=label_font)
address2_label.place(x = 0, y = 141)
address2_label.configure(foreground="#000000",background="#00BFFF")

address2_entry = tkinter.Entry(infor_frame)
address2_entry.place(x = 220, y = 143,width= 655, height= 23)
address2_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")


notes_label = tkinter.Label(infor_frame, text = "Ghi chú/ Notes:",font=label_font)
notes_label.place(x = 0, y = 171)
notes_label.configure(foreground="#000000",background="#00BFFF")

notes_entry = tkinter.Entry(infor_frame)
notes_entry.place(x = 220, y = 173,width= 655, height= 23)
notes_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

#frame2
infor_frame2 = tkinter.Frame(window, bg="#00BFFF", width=250, height=120)
infor_frame2.place(x=910, y=15)

NO_label = tkinter.Label(infor_frame2, text = "Số phiếu / NO:",font=label_font)
NO_label.place(x = 3, y = 3)
NO_label.configure(foreground="#000000",background="#00BFFF")

NO_entry = tkinter.Entry(infor_frame2,justify="center")
NO_entry.place(x = 127, y = 3,width= 90, height= 23)
NO_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")
NO_entry.insert(0, "PO000")

tax_label = tkinter.Label(infor_frame2, text = "Thuế / TAX:",font=label_font)
tax_label.place(x = 3, y = 33)
tax_label.configure(foreground="#000000",background="#00BFFF")
percent_label = tkinter.Label(infor_frame2, text = "(%)",font=label_font)
percent_label.place(x = 216, y = 31)
percent_label.configure(foreground="#000000",background="#00BFFF")

tax_entry = tkinter.Entry(infor_frame2,justify="center")
tax_entry.place(x = 127, y = 33,width= 90, height= 23)
tax_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")
tax_entry.insert(0, "8")

calendar_icon = Image.open("H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Customize PDF\date.png")
calendar_icon = calendar_icon.resize((19, 19))
calendar_icon = ImageTk.PhotoImage(calendar_icon)
current_date = date.today()

ngaydathang_label = tkinter.Label(infor_frame2, text="Ngày đặt hàng:", font=label_font)
ngaydathang_label.place(x = 3, y = 63)
ngaydathang_label.configure(foreground="#000000",background="#00BFFF")

ngaydathang_entry = tkinter.Entry(infor_frame2,justify="center")
ngaydathang_entry.place(x = 127, y = 63,width= 90, height= 23)
ngaydathang_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

ngaydathang_button = tkinter.Button(infor_frame2, image=calendar_icon, font=("Arial", 12), command=show_date1_window, cursor="hand2")
ngaydathang_button.place(x = 222, y = 63)
ngaydathang_button.configure(foreground="#00BFFF",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

date1 = tkinter.Tk()
date1.title("Date - đặt hàng")
date1.withdraw()
date1.overrideredirect(1)

date1_var = tkinter.StringVar()
date1_var.set(current_date.strftime("%d/%m/%Y"))
cal1 = Calendar(date1, selectmode="day", date_var=date1_var,date_pattern='dd/mm/yyyy')
cal1.pack(pady=20, fill='both', expand=True)
cal1.bind("<<CalendarSelected>>", dathang_date)

ngaygiaohang_label = tkinter.Label(infor_frame2, text = "Ngày giao hàng:",font=label_font)
ngaygiaohang_label.place(x = 3, y = 93)
ngaygiaohang_label.configure(foreground="#000000",background="#00BFFF")

ngaygiaohang_entry = tkinter.Entry(infor_frame2,justify="center")
ngaygiaohang_entry.place(x = 127, y = 93,width= 90, height= 23)
ngaygiaohang_entry.configure(font=entry_font,relief="flat",foreground="#000000",background="#ffffff",highlightthickness= 0.5,highlightbackground="black")

ngaygiaohang_button = tkinter.Button(infor_frame2, image=calendar_icon, font=("Arial", 12), command=show_date2_window, cursor="hand2")
ngaygiaohang_button.place(x = 222, y = 93)
ngaygiaohang_button.configure(foreground="#00BFFF",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

date2 = tkinter.Tk()
date2.title("Date - giao hàng")
date2.withdraw()
date2.overrideredirect(1)

date2_var = tkinter.StringVar()
date2_var.set(current_date.strftime("%d/%m/%Y"))
cal2 = Calendar(date2, selectmode="day", date_var=date2_var,date_pattern='dd/mm/yyyy')
cal2.pack(pady=20, fill='both', expand=True)
cal2.bind("<<CalendarSelected>>", giaohang_date)

#Seletion
seletion_frame = tkinter.Frame(window, bg="#00BFFF", width=1335, height=60)
seletion_frame.place(x=15, y=220)

products_name_label = tkinter.Label(seletion_frame, text = "Tên sản phẩm/ Product Name",font=label_font)
products_name_label.place(x = 105, y = 3)
products_name_label.configure(foreground="#000000",background="#00BFFF")

products_name = ["Sprite Plastic 300ml - Sprite chai nhua 300ml", 
                 "Dasani Water", 
                 "Fanta Plastic 300ml - Fanta chai nhua 300ml",
                 "Coke Plastic 300ml - Coca chai nhua 300ml"]
products_name_combobox = ttk.Combobox(seletion_frame, values=products_name)
products_name_combobox.place(x = 3, y = 27 ,width= 450, height= 26)
products_name_combobox.configure(font=label_font)
products_name_combobox.option_add("*TCombobox*Listbox.font", label_font)
# products_name_combobox.bind("<<ComboboxSelected>>", on_product_name_select1)
# products_name_combobox.bind("<<ComboboxSelected>>", on_desc_selected)
products_name_combobox.bind("<<ComboboxSelected>>", lambda event: (on_product_name_select1(event), on_desc_selected(event)))
# # Xử lý lựa chọn mã hàng ở đây
product_code_label = tkinter.Label(seletion_frame, text = "Mã hàng/ Product Code",font=label_font)
product_code_label.place(x = 505, y = 3)
product_code_label.configure(foreground="#000000",background="#00BFFF")

products_code_dict = { "Sprite Plastic 300ml - Sprite chai nhua 300ml": "Coca 1",
                        "Dasani Water": "Coca 2",
                        "Fanta Plastic 300ml - Fanta chai nhua 300ml": "Coca 3",
                        "Coke Plastic 300ml - Coca chai nhua 300ml": "Coca 4"}
products_code_var = tkinter.StringVar()

products_code_combobox = ttk.Combobox(seletion_frame, textvariable=products_code_var, state="normal",justify="center")
products_code_combobox.place(x = 510, y = 27,width= 170, height= 26)
products_code_combobox.configure(font=label_font)

dg_label = tkinter.Label(seletion_frame, text = "Đơn giá/ Price Unit",font=label_font)
dg_label.place(x = 765, y = 3)
dg_label.configure(foreground="#000000",background="#00BFFF")

product_prices = {
    "Sprite Plastic 300ml - Sprite chai nhua 300ml": "2,702.42",
    "Dasani Water": "3,165.71",
    "Fanta Plastic 300ml - Fanta chai nhua 300ml": "2,702.42",  
    "Coke Plastic 300ml - Coca chai nhua 300ml": "2,702.42" }
dg_var = tkinter.StringVar()
dg_combobox = ttk.Combobox(seletion_frame, textvariable=dg_var, state="normal",justify="center")
dg_combobox.configure(font=label_font)
dg_combobox.place(x = 755, y = 28,width= 170, height= 26)

dvt_label = tkinter.Label(seletion_frame, text = "Đơn vị tính/ UoM",font=label_font)
dvt_label.place(x = 994, y = 3)
dvt_label.configure(foreground="#000000",background="#00BFFF")
product_dvt = {
    "Sprite Plastic 300ml - Sprite chai nhua 300ml": "Bottle",
    "Dasani Water": "Bottle",
    "Fanta Plastic 300ml - Fanta chai nhua 300ml": "Bottle",  
    "Coke Plastic 300ml - Coca chai nhua 300ml": "Bottle" }
dvt_var = tkinter.StringVar()
dvt_combobox = ttk.Combobox(seletion_frame,textvariable=dvt_var,state="normal",justify="center")
dvt_combobox.place(x = 1000, y = 28,width= 120, height= 26)
dvt_combobox.configure(font=label_font)

quantity_label = tkinter.Label(seletion_frame, text="Số lượng/ Quantity",font=label_font)
quantity_label.place(x = 1178, y = 3)
quantity_label.configure(foreground="#000000",background="#00BFFF")
quantity_spinbox = tkinter.Spinbox(seletion_frame,from_= 1, to = 1000,justify="center")
quantity_spinbox.configure(font=label_font)
quantity_spinbox.place(x = 1190, y = 28,width= 120, height= 26)

#Table
columns = ( 'stt','code','desc','qty', 'UoM' ,'price', 'total',)

tree = ttk.Treeview(window, columns=columns, show="headings")
tree.column('stt', width=60,anchor='center')
tree.column('code', width=130,anchor='center')
tree.column('qty', width=80,anchor='center')
tree.column('desc', width=500)
tree.column('UoM', width=90,anchor='center')
tree.column('price', width=100,anchor='center')
tree.column('total', width=100,anchor='e')

tree.heading('stt', text='STT')
tree.heading('code', text='Product Code')
tree.heading('qty', text='Quantity')
tree.heading('desc', text='Description')
tree.heading('UoM', text='UoM')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")


style = ttk.Style()
style.configure("Treeview.Heading", font=("Times new roman", 13,'bold'), foreground = "black", background = "black")

tree.place(x=265,y=300)
tree.bind("<Button-1>", select_tree_item)
tree.bind("<<TreeviewSelect>>", select_tree_item)

#Button
add_item_button = tkinter.Button(window, text = "Add item", command= add_item,font=label_font)
add_item_button.place(x = 910, y = 160, width=80, height=40)
add_item_button.configure(foreground="black",background="#4ED70F",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

update_button = tkinter.Button(window, text="Modify", command=modify_item, font=label_font)
update_button.place(x = 1030, y = 160, width=80, height=40)
update_button.configure(foreground="black",background="#FFFF33",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")


delete_button = tkinter.Button(window, text="Delete", command=delete_item, font=label_font)
delete_button.place(x = 1150, y = 160, width=80, height=40)
delete_button.configure(foreground="black",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

new_po_button = tkinter.Button(window, text="New Purchase Order",command=new_po,font=label_font)
new_po_button.place(x = 20, y = 360, width=230, height=40) 
new_po_button.configure(foreground="black",background="#4ED70F",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

exit_button = tkinter.Button(window, text="Turn Off Application", command = close_window,font=label_font)
exit_button.place(x = 20, y = 420, width=230, height=40)
exit_button.configure(foreground="#ffffff",background="#CF1E14",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

total_button = tkinter.Button(window, text="Total", command= total_po,font=label_font)
total_button.place(x = 20, y = 480, width=230, height=40)
total_button.configure(foreground="black",background="#FFFF33",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")
total_frame = tkinter.Frame(window, bg="#FFFF33", width=400, height=90)

save_po_button = tkinter.Button(window, text="Generate Purchase Order",command=generate_po,font=label_font)
save_po_button.place(x = 20, y = 300, width=230, height=40) 
save_po_button.configure(foreground="black",background="#00BFFF",relief="flat",overrelief="flat",cursor="hand2",borderwidth="0")

window.protocol("WM_DELETE_WINDOW",on_closing)
window.mainloop()







