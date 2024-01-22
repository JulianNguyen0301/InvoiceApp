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
def main():
    document = r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\PDF_Invoice\159196_Coca.pdf"
    feature1(document)
if __name__ == "__main__":
    main() 