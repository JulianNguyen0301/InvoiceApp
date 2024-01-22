import pdfplumber 
from time import time
def feature1(document):
    PRODUCT_data = []
    MH_data = []
    DVT_data = []
    SL_data = []
    DG_data = []
    TT_data = []
    NOVAT_data = [] 
    VAT_data = []
    TOTAL_data = []
    STT_temp = []
    with pdfplumber.open(document) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            tables = page.extract_tables()
            for table_index, table in enumerate(tables, start=1):
                for row_index, row in enumerate(table, start=1):
                    for cell_index, cell_text in enumerate(row, start=1):
                        #any(keyword in cell_text for keyword in [""])
                        if cell_text and any(keyword in cell_text for keyword in ["STT\n(No)","Số\nTT","STT"]):
                            if row_index - 1  < len(table): #So sánh với số hàng hiện có trong bảng
                                next_row_value = table[row_index +1][cell_index - 1]
                                STT_temp.append(next_row_value.split())
                                count_STT = len(STT_temp[0])
                                print(STT_temp[0])
                        if cell_text and any(keyword in cell_text for keyword in ["Tên hàng hóa, dịch vụ","Products Description","Goods, services"]):
                            if row_index + 1 < len(table):
                                PRODUCT_temp = table[row_index + 1][cell_index - 1]
                                PRODUCT_data.append(PRODUCT_temp.splitlines())
                                print(PRODUCT_data[0])

                        if cell_text and any(keyword in cell_text for keyword in ["Article No","Mã hàng"]):
                            if row_index + 1 < len(table):
                                MH_temp = table[row_index + 1][cell_index - 1]
                                MH_data.append(MH_temp.splitlines())
                                print(MH_data[0])

                        if cell_text and any(keyword in cell_text for keyword in ["Đơn vị tính" ,"Đơn vị"]):
                            if row_index + 1 < len(table):
                                DVT_temp = table[row_index + 1][cell_index - 1]
                                DVT_data.append(DVT_temp.splitlines())
                                print(DVT_data[0])
                        if cell_text and any(keyword in cell_text for keyword in ["Quantity" ,"Số lượng"]):
                            if row_index + 1 < len(table):
                                SL_temp = table[row_index + 1][cell_index - 1]
                                SL_data.append(SL_temp.splitlines())
                                print(SL_data[0])
                        if cell_text and any(keyword in cell_text for keyword in ["Unit price" ,"Đơn giá"]):     
                            if row_index + 1 < len(table):
                                DG_temp = table[row_index + 1][cell_index - 1]
                                DG_data.append(DG_temp.splitlines())
                                print(DG_data[0])
                        if cell_text and any(keyword in cell_text for keyword in ["Amount" ,"Thành tiền"]):     
                            if row_index + 1 < len(table):
                                TT_temp = table[row_index + 1][cell_index - 1]
                                TT_data.append(TT_temp.splitlines())
                                print(TT_data[0])

                        if cell_text and any(keyword in cell_text for keyword in ["Cộng tiền hàng"]):  
                            colon_index = cell_text.index(':')
                            NOVAT_data = cell_text[colon_index + 1:].strip()
                            if ":" in NOVAT_data:
                                colon_index = NOVAT_data.index(':')
                                NOVAT_data = NOVAT_data[colon_index + 1:].strip()
                            print(NOVAT_data)

                        if cell_text and any(keyword in cell_text for keyword in ["Tiền thuế GTGT"]): 
                            colon_index = cell_text.index(':')
                            VAT_data = cell_text[colon_index + 1:].strip()
                            if ":" in VAT_data:
                                colon_index = VAT_data.index(':')
                                VAT_data = VAT_data[colon_index + 1:].strip()
                                print(VAT_data)
                        if cell_text and any(keyword in cell_text for keyword in ["thanh toán"]): 
                            colon_index = cell_text.index(':')
                            TOTAL_data = cell_text[colon_index + 1:].strip()
                            if ":" in TOTAL_data:
                                colon_index = TOTAL_data.index(':')
                                TOTAL_data = TOTAL_data[colon_index + 1:].strip()
                            print(TOTAL_data)
def feature2(document):
    PRODUCT_data = []
    MH_data = []
    DVT_data = []
    SL_data = []
    DG_data = []
    TT_data = []
    NOVAT_data = [] 
    VAT_data = []
    TOTAL_data = []
    STT_data = []
    Data_temp = []
    Data_table_filtered  = []
    count = []
    Data_table_processed1 = []
    with pdfplumber.open(document) as pdf:
        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            tables = page.extract_tables()
            for table_index, table in enumerate(tables, start=1):
                for row_index, row in enumerate(table, start=1):
                    Data_table_processed = []
                    Data_table_filtered = list(filter(None, row))
                    Data_table_processed1.append(Data_table_filtered)             
                    if any(keyword in Data_table_filtered for keyword in ["Sản phẩm (Products)", "Tên hàng hóa, dịch vụ", "Tên hàng hóa, dịch vụ\n(Products Description)", "Sản phẩm\n(Products)","Tên hàng hóa, dịch vụ\n(Description)"]):
                        row_index_temp = row_index
                        for cell in Data_table_filtered:
                            if cell:
                                cell = cell.replace("\n", " ")
                                Data_table_processed.append(cell)
                                
                        if any(keyword in Data_table_processed for keyword in ["STT (Seq)", "Số TT", "STT", "STT (No.)","STT (No)"]):
                            row_index_temp = row_index
                            all_rows = []      
                            while row_index_temp < len(table) - 1:
                                Current_row = []
                                for item in filter(None, table[row_index_temp]):
                                    Current_row.append(item.replace("\n", " "))       
                                if Current_row and Current_row[0].isdigit() :  # Check if the list is not empty
                                        all_rows.append(Current_row)
                                row_index_temp += 1
                            for index, row in enumerate(all_rows):
                                Data_temp.append(row)
                    newline_count = 0
                    for element in Data_table_filtered:
                        newline_count += element.count("\n")
                        count.append(newline_count)
        max_value = count[0] 
        for data in count:
            if data > max_value:
                max_value = data
        
        if max_value >= 30:
            data_temp1 = []
            for i in range(len(Data_table_processed1)-1):
                if Data_table_processed1[i] and any(keyword in Data_table_processed1[i][0] for keyword in ["Cộng tiền hàng","Tổng cộng tiền trước thuế "]):
                    data_temp1.append(Data_table_processed1[i][0].splitlines())
            for i in range(len(data_temp1[0])-1):
                if data_temp1[0][i] and any(keyword in data_temp1[0][i] for keyword in ["Cộng tiền hàng","Tổng cộng tiền trước thuế "]):
                    NOVAT_data.append(data_temp1[0][i])
                    VAT_data.append(data_temp1[0][i+1])
                    TOTAL_data.append(data_temp1[0][i+2])
        else:
            for i in range(len(Data_table_processed1)-1):
                if Data_table_processed1[i] and any(keyword in Data_table_processed1[i][0] for keyword in ["Cộng tiền hàng","Tổng cộng tiền trước thuế "]):
                    NOVAT_data.append(Data_table_processed1[i][-1])
                    VAT_data.append(Data_table_processed1[i+1][-1])
                    TOTAL_data.append(Data_table_processed1[i+2][-1])
        for index_data, i in enumerate(Data_temp):
            if len(Data_temp[index_data][1]) < 3:
                del Data_temp[index_data]
        index = 0
        max = int(Data_temp[0][0])
        while index < len(Data_temp) - 1:
            if max >= int(Data_temp[index+1][0]):
                del Data_temp[index+1]
                index -=1
            else: 
                max = int(Data_temp[index+1][0])
            index += 1
        count_data = len(Data_temp[0])
        #print(count_data)
        if count_data == 7:
            for element in Data_temp:
                STT_data.append(element[0])
                MH_data.append(element[1])
                PRODUCT_data.append(element[2])
                DVT_data.append(element[3])
                SL_data.append(element[4])
                DG_data.append(element[5])
                TT_data.append(element[6])
            
        elif count_data == 6:
            for element in Data_temp:
                STT_data.append(element[0])
                PRODUCT_data.append(element[1])
                DVT_data.append(element[2])
                SL_data.append(element[3])
                DG_data.append(element[4])
                TT_data.append(element[5])
            
        print(STT_data)
        print(PRODUCT_data)
        print(MH_data)
        print(DVT_data)
        print(SL_data)
        print(DG_data)
        print(TT_data)
        for element in NOVAT_data:
            if ":" in element:
                for element in NOVAT_data:
                    colon_index = element.index(':')
                    NOVAT_data = element[colon_index + 1:].strip()
                    if ":" in NOVAT_data:
                        colon_index = NOVAT_data.index(':')
                        NOVAT_data = NOVAT_data[colon_index + 1:].strip()
                    print(NOVAT_data)
                for element in VAT_data:
                    colon_index = element.index(':')
                    VAT_data = element[colon_index + 1:].strip()
                    if ":" in VAT_data:
                        colon_index = VAT_data.index(':')
                        VAT_data = VAT_data[colon_index + 1:].strip()
                    print(VAT_data)
                for element in TOTAL_data:
                    colon_index = element.index(':')
                    TOTAL_data = element[colon_index + 1:].strip()
                    if ":" in TOTAL_data:
                        colon_index = TOTAL_data.index(':')
                        TOTAL_data = TOTAL_data[colon_index + 1:].strip()
                    print(TOTAL_data)
            else:
                print(NOVAT_data)
                print(VAT_data)
                print(TOTAL_data)

#Dạng 2
#"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\TEST_INVOICE_PDF\pepsico.pdf"
#"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\TEST_INVOICE_PDF\highland.pdf"
#"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\TEST_INVOICE_PDF\vietthang.pdf"
#Dạng 1: Tất cả nội dung cần biết chứa trong 1 hàng
#"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\172819_Coca.pdf"
#"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\PDF_Invoice\Invoice_LinhKhoa_2.pdf"
#Dạng 3 
#Không có STT nhưng có mã hàng
#"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\TEST_INVOICE_PDF\PWC.pdf"
#Phân biệt dạng 1 và 2:
#Chung: Tìm hàng chứa Số
def main():
    document = r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\TEST_INVOICE_PDF\pepsico_00229788.pdf"
    print(document)
    start = time()
    with pdfplumber.open(document) as pdf:
        feature1_called = False
        feature2_called = False

        for page_number in range(len(pdf.pages)):
            page = pdf.pages[page_number]
            tables = page.extract_tables()

            for table_index, table in enumerate(tables, start=1):
                for row_index, row in enumerate(table, start=1):
                    Data_table_processed = []
                    Data_table_filtered = list(filter(None, row))
                    
                    if any(keyword in Data_table_filtered for keyword in ["Sản phẩm (Products)", "Tên hàng hóa, dịch vụ", "Tên hàng hóa, dịch vụ\n(Products Description)", "Sản phẩm\n(Products)","Tên hàng hóa, dịch vụ\n(Description)","Tên hàng hóa, dịch vụ\n(Goods, services)"]):
                        row_index_temp = row_index
                        current_row = []
                        
                        while row_index_temp < len(table) - 1:
                            for item in filter(None, table[row_index_temp]):
                                current_row.append(item)
                            row_index_temp += 1

                        count_max = current_row[0].count("\n")
                        for element in current_row:
                            count_temp = element.count("\n")
                            if count_temp >= count_max:
                                count_max = count_temp
                        print(count_max)
                        if count_max >= 3 and not feature1_called:
                            print("A")
                            feature1(document)
                            feature1_called = True
                            
                        elif not feature2_called:
                            print("B")
                            feature2(document)
                            feature2_called = True
    print(time()-start)                        
if __name__ == "__main__":
    main()                         
                            


                    
                        

           
# img = cv2.imread(r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Images_Invoice\Sample_6.png")
# #Nhập tùy chỉnh theo yêu cầu
# x1 = 75      #Điểm bắt đầu theo chiều ngang
# x2 = 2000    #Điểm kết thúc theo chiều ngang
# y1 = 800      #Điểm bắt đầu theo chiều dọc
# y2 = 2339    #Điểm kết thúc theo chiều dọc
# img_crop = img[y1:y2,x1:x2]
# cv2.imshow("img_crop",img_crop)
# print(img.shape)
# cv2.waitKey(0)
# num_regions = 10



# for _ in range(num_regions):
#     x, y, w, h = cv2.selectROI("Select ROI", img, fromCenter=False, showCrosshair=False)
#     cv2.destroyAllWindows()  # Close the window after selecting ROI

#     # Extract the selected region
#     roi = img[y:y + h, x:x + w]

#     data_stt = pytesseract.image_to_string(roi, lang = 'vie',config= '--oem 3 --psm 6')
#     print("Text in selected region:")
#     print(data_stt)

#     # Draw rectangle on the selected region
#     cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
#     cv2.imshow("Selected Region", img)
#     cv2.waitKey(0)

# cv2.destroyAllWindows()



