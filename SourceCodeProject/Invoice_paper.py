import pytesseract 
import numpy as np
import cv2
from pytesseract import Output
import os
import xlwings as xw
from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.styles import Font
pytesseract.pytesseract.tesseract_cmd =r'H:\\APP UNIVERSITY\\CODE PYTHON\\Tesseract-ocr\\tesseract.exe'

def make_rows_bold(excel_file_path, start_column, end_column, row):
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets.active
    for col in range(start_column, end_column + 1):
        sheet.range(row, col).api.Font.Bold = True
    wb.save()

def make_columns_bold(excel_file_path, start_row,end_row,col):
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets.active
    for row in range(start_row, end_row + 1):
        sheet.range(row, col).api.Font.Bold = True
    wb.save()

def change_font_size_and_save(wb, font_size):
    for sheet in wb.sheets:
        for row in sheet.used_range.rows:
            for cell in row:
                cell.api.Font.Size = font_size
    wb.save()

def fit_column_widths_for_one_sheet(sheet):
    cell_widths = {}
    used_range = sheet.used_range
    for col in range(1, used_range.columns.count + 1):
        for row in range(1, used_range.rows.count + 1):
            cell_value = used_range(row, col).value
            if cell_value:
                if str(cell_value)[0] == "=":
                    continue
                cell_widths[col] = max(
                    (cell_widths.get(col, 0), len(str(cell_value))+2)
                )
    for col, column_width in cell_widths.items():
        column_width = str(column_width)
        sheet.range((1, col), (used_range.rows.count, col)).column_width = column_width

def process_all_worksheets_cocacola(excel_file_path):
    wb = xw.Book(excel_file_path)
    for sheet in wb.sheets:
        fit_column_widths_for_one_sheet(sheet)
        sheet.range("A10:G10").api.HorizontalAlignment = -4108  # -4108 tương ứng với giá trị xlCenter trong Excel
        sheet.range("A11:A30").api.HorizontalAlignment = -4108
        sheet.range("A1:B8").api.HorizontalAlignment = -4131
        sheet.range("C11:F30").api.HorizontalAlignment = -4108
        sheet.range("G5:G8").api.HorizontalAlignment = -4152
        sheet.range("G11:G30").api.HorizontalAlignment = -4152
    wb.save()
    wb.close()

def process_all_worksheets_linhkhoa(excel_file_path):
    wb = xw.Book(excel_file_path)
    for sheet in wb.sheets:
        fit_column_widths_for_one_sheet(sheet)
        sheet.range("A11:F11").api.HorizontalAlignment = -4108  # -4108 tương ứng với giá trị xlCenter trong Excel
        sheet.range("A12:A30").api.HorizontalAlignment = -4108
        sheet.range("A1:B9").api.HorizontalAlignment = -4131       
        sheet.range("C12:D30").api.HorizontalAlignment = -4108
        sheet.range("E12:E30").api.HorizontalAlignment = -4108
        sheet.range("E6:E9").api.HorizontalAlignment = -4152
        sheet.range("F6:F30").api.HorizontalAlignment = -4152
    wb.save()
    wb.close()

def split_text_function(text):
    split_text = text.split("\n")
    split_text = list(filter(lambda item: item.strip(), split_text))  
    result_string = "\n".join(split_text)
    return result_string

def pre_image(img,x1,y1,x2,y2):
    region_of_interest = img[y1:y2, x1:x2].copy()
    _, thresh = cv2.threshold(region_of_interest, 128, 255, cv2.THRESH_BINARY)
    region = cv2.GaussianBlur(thresh, (5, 5), 0)
    img[y1:y2, x1:x2] = region
    return img

def del_space(result):
    indexes_to_remove = []
    for i in range(len(result["conf"])):
        if result["conf"][i] == -1 or result["text"][i] == '':
            indexes_to_remove.append(i)

    for key, value in result.items():
        result[key] = [value[i] for i in range(len(value)) if i not in indexes_to_remove]

    data = result["text"]
    row_size = 1  # Kích thước của từng dòng
    result1 = [data[i:i + row_size] for i in range(0, len(data), row_size)]
    return result1



def invoice_cocacola(img,dem):
    
    rgb = cv2.cvtColor(img,cv2.COLOR_BGR2RGB)

    start_row, end_row = 270,985
    start_col, end_col = 95, 1560
    half_1 = rgb[start_row:end_row, start_col:end_col]
    half_1 = cv2.convertScaleAbs(half_1, alpha=1.22, beta=10)

    start_row, end_row = 985, 1600
    start_col, end_col = 95, 1560
    image_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    grey = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    half_2 = grey[start_row:end_row, start_col:end_col]
    del_half_2 = half_2.copy()
    del_half_2 = pre_image(del_half_2, 12, 158,63, 343)
    del_half_2 = pre_image(del_half_2, 1127, 155,1245, 348)
    del_half_2 = pre_image(del_half_2, 1175, 354, 1463, 398)
    del_half_2 = pre_image(del_half_2, 1186, 407,1463, 452)
    del_half_2 = pre_image(del_half_2, 1181, 464, 1462, 510)
    del_half_2 = pre_image(del_half_2, 323, 518, 1249, 565)
    region_of_interest = del_half_2[156:346,870:995].copy()
    _, region_of_interest = cv2.threshold(region_of_interest, 128, 255, cv2.THRESH_BINARY)
    region = cv2.GaussianBlur(region_of_interest, (5, 5), 0)
    del_half_2[156:346,870:995] = region

    roi1 = [[(1013, 65), (1142, 98), 'text', 'Ký hiệu'],
            [(1242, 48), (1402, 94), 'text', 'HĐ số'], 
            [(1142, 112), (1312, 160), 'text', 'Ngày'],
            [(10, 235), (230, 294), 'text', 'Mã số khách hàng'], 
            [(402, 234), (1118, 345), 'text', 'Tên khách hàng'], 
            [(275, 354), (1118, 493), 'text', 'Địa chỉ khách hàng'], 
            [(1193, 427), (1409, 485), 'text', 'Mã số thuế khách hàng'], 
            [(8, 553), (1118, 701), 'text', 'Những thông tin giao hàng khác']]
    
    roi2 = [[(12, 158), (63, 343), 'text', 'STT'], 
            [(68, 155), (738, 346), 'text', 'Tên hàng hóa, dịch vụ'], 
            [(742, 155), (866, 344), 'text', 'Mã hàng'], 
            [(870, 156), (995, 346), 'text', 'Đơn vị tính'], 
            [(1000, 157), (1121, 346), 'text', 'Số lượng'], 
            [(1127, 155), (1245, 348), 'text', 'Đơn giá'], 
            [(1250, 154), (1457, 345), 'text', 'Thành tiền'], 
            [(1175, 354), (1463, 398), 'text', 'Cộng tiền hàng chưa có thuế GTGT'], 
            [(1186, 407), (1463, 452), 'text', 'Tiền thuế GTGT'], 
            [(1181, 464), (1462, 510), 'text', 'Tổng cộng tiền thanh toán'],
            [(323, 518), (1249, 565), 'text', 'Số tiền viết bằng chữ']]
        
    wb = xw.Book()
    sht = wb.sheets.active
    imgShow = half_1.copy()
    imgMask = np.zeros_like(imgShow)
    for x,r in enumerate(roi1):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,0,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = half_1[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        #cv2.imshow(str(x),imgCrop)
        if r[2] == 'text':
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("A" + str(x + 1)).value = r[3]
            sht.range("B"+ str(x + 1) ).number_format = "@"
            sht.range("B" + str(x + 1)).value = full_text
    imgShow = del_half_2.copy()
    imgMask = np.zeros_like(imgShow)
    for x,r in enumerate(roi2):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,0,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = del_half_2[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        if r[3] == 'STT':
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("A10").value = r[3]
            for index, line in enumerate(lines, start=11):
                sht.range("A" + str(index)).number_format = "@" 
                sht.range("A" + str(index)).value = line
        if r[3] == 'Tên hàng hóa, dịch vụ':
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("B10").value = r[3]
            for index, line in enumerate(lines, start=11):
                sht.range("B" + str(index)).number_format = "@" 
                sht.range("B" + str(index)).value = line
        if r[3] == 'Mã hàng':
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("C10").value = r[3]
            for index, row in enumerate(rows, start=11):
                sht.range("C" + str(index)).value = row[0]
        if r[3] == 'Đơn vị tính':
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            sht.range("D10").value = r[3]
            for index, line in enumerate(lines, start=11):
                sht.range("D" + str(index)).value = line
        if r[3] == 'Số lượng':
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("E10").value = r[3]
            for index, line in enumerate(lines, start=11):
                sht.range("E" + str(index)).number_format = "@"
                sht.range("E" + str(index)).value = line
        if r[3] == 'Đơn giá':
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("F10").value = r[3]
            for index, row in enumerate(rows, start=11):
                sht.range("F" + str(index)).number_format = "@"
                sht.range("F" + str(index)).value = row[0]
        if r[3] == 'Thành tiền':
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("G10").value = r[3]
            for index, row in enumerate(rows, start=11):
                sht.range("G" + str(index)).number_format = "@"
                sht.range("G" + str(index)).value = row[0]
        if r[3] == 'Cộng tiền hàng chưa có thuế GTGT':
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("F5").value = r[3]
            sht.range("G5").number_format = "@"
            sht.range("G5").value = rows
        if r[3] == 'Tiền thuế GTGT':
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("F6").value = r[3]
            sht.range("G6").number_format = "@"
            sht.range("G6").value = rows
        if r[3] == 'Tổng cộng tiền thanh toán':
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("F7").value = r[3]
            sht.range("G7").number_format = "@"
            sht.range("G7").value = rows
        if r[3] == 'Số tiền viết bằng chữ':
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("F8").value = r[3]
            sht.range("G8").value = full_text    

    file_name = f"Invoice_Cocacola_P_{dem}.xlsx"
    wb.save(fr'SourcecodeOCR\Data_excel_Cocacola\{file_name}')
    excel_file_path = fr"SourcecodeOCR\Data_excel_Cocacola\{file_name}"  
    make_columns_bold(excel_file_path, 1,8,1)
    make_columns_bold(excel_file_path, 5,8,6)
    make_rows_bold(excel_file_path, 1, 7, 10)
    wb = xw.Book(excel_file_path)
    change_font_size_and_save(wb, 12)
    process_all_worksheets_cocacola(excel_file_path)  

def invoice_linhkhoa(img,dem):

    grey = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    start_row, end_row = 65 , 803
    start_col, end_col = 75, 1530
    image1 = grey[start_row:end_row, start_col:end_col]
    image1 = cv2.GaussianBlur(image1, (5, 5), 0)
    
    start_row, end_row = 801 , 1474
    start_col, end_col = 78, 1573
    image_2_1 = grey[start_row:end_row, start_col:end_col]

    start_row, end_row = 1470 , 1665
    start_col, end_col = 78, 1573
    image3 = img[start_row:end_row, start_col:end_col]
    image3 = cv2.cvtColor(image3,cv2.COLOR_BGR2GRAY)  
   
    roi1 = [[(1323, 43), (1446, 77), 'text', 'Ký hiệu'], 
            [(583, 87), (914, 128), 'text', 'Ngày'], 
            [(1312, 88), (1449, 127), 'text', 'Số'], 
            [(275, 502), (809, 547), 'text', 'Họ tên người mua hàng'], 
            [(137, 547), (821, 587), 'text', 'Tên đơn vị'], 
            [(143, 599), (320, 633), 'text', 'Mã số thuế'], 
            [(97, 645), (1130, 682), 'text', 'Địa chỉ'], 
            [(248, 694), (391, 732), 'text', 'Hình thức thanh toán'], 
            [(790, 696), (1132, 732), 'text', 'Số tài khoản']]
    
    roi2 = [[(15, 99), (80, 142), 'STT', '1'], 
            [(15, 146), (80, 190), 'STT', '2'], 
            [(15, 194), (80, 236), 'STT', '3'], 
            [(15, 243), (80, 286), 'STT', '4'], 
            [(15, 292), (80, 332), 'STT', '5'], 
            [(15, 339), (80, 380), 'STT', '6'], 
            [(15, 389), (80, 428), 'STT', '7'], 
            [(15, 435), (80, 474), 'STT', '8'], 
            [(15, 484), (80, 523), 'STT', '9'], 
            [(15, 532), (80, 573), 'STT', '10'], 
            [(15, 578), (80, 621), 'STT', '11'], 
            [(15, 627), (80, 666), 'STT', '12'], 
            [(101, 100), (557, 143), 'HH', '1'],
            [(101, 147), (557, 192), 'HH', '2'], 
            [(101, 194), (557, 237), 'HH', '3'], 
            [(101, 243), (557, 285), 'HH', '4'], 
            [(101, 292), (557, 333), 'HH', '5'], 
            [(101, 340), (557, 380), 'HH', '6'], 
            [(101, 387), (557, 428), 'HH', '7'], 
            [(101, 435), (557, 477), 'HH', '8'], 
            [(101, 483), (557, 521), 'HH', '9'], 
            [(101, 532), (557, 571), 'HH', '10'], 
            [(101, 581), (557, 617), 'HH', '11'], 
            [(101, 627), (557, 666), 'HH', '12'],
            [(701, 105), (774, 664), 'text', 'Đơn vị tính'], 
            [(941, 104), (1032, 665), 'text', 'Số lượng'], 
            [(1054, 101), (1245, 667), 'text', 'Đơn giá'], 
            [(1303, 101), (1486, 666), 'text', 'Thành tiền']]
    
    roi3 = [[(1118, 7), (1487, 45), 'text', 'Cộng tiền hàng'],
        [(1116, 55), (1486, 93), 'text', 'Tiền thuế GTGT'],
        [(1121, 100), (1485, 140), 'text', 'Tổng tiền thanh toán'],
        [(253, 146), (1210, 190), 'text', 'Số tiền viết bằng chữ']]

    wb = xw.Book()
    sht = wb.sheets.active
    imgShow = image1.copy()
    imgMask = np.zeros_like(imgShow)
    for x,r in enumerate(roi1):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,0,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = image1[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        if x >= 0 and x <=2:
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("A" + str(x + 1)).value = r[3]
            sht.range("B"+ str(x + 1) ).number_format = "@"
            sht.range("B" + str(x + 1)).value = full_text


        if r[3] == "Họ tên người mua hàng":
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='vie+eng', config='tessdata', output_type=Output.DICT) #, output_type=Output.DICT
            rows = del_space(result_2_1)
            sht.range("A" + str(x + 1)).value = r[3]
            sht.range("B"+ str(x + 1) ).number_format = "@"
            sht.range("B" + str(x + 1)).value = rows

        if x >= 4 and x <=8:
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("A" + str(x + 1)).value = r[3]
            sht.range("B"+ str(x + 1) ).number_format = "@"
            sht.range("B" + str(x + 1)).value = full_text
       
    
    imgShow = image_2_1.copy()
    imgMask = np.zeros_like(imgShow)
    table = []
    line_count  = 0
    for x,r in enumerate(roi2):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,0,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = image_2_1[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        #cv2.imshow(str(x),imgCrop)
        if r[2] == "STT":
            _, imgCrop = cv2.threshold(imgCrop, 126, 255, cv2.THRESH_TRUNC)
            kernel = np.array([[-1, -1, -1],
                    [-1,  10, -1],
                    [-1, -1, -1]])
            imgCrop = cv2.filter2D(imgCrop, -1, kernel)
            #cv2.imshow(str(x),imgCrop)
            text = pytesseract.image_to_string(imgCrop, lang='eng', config='--oem 3 --psm 6')
            result_string = split_text_function(text)
            full_text = " ".join(lines)
            lines = result_string.split("\n")
            sht.range("A11").value = r[2]
            sht.range("A" + str(x + 11)).value = full_text
            for line in (lines):
                if line.strip():
                    line_count += 1
                    #print(line)
            

        elif r[3] == "Đơn vị tính":
            #_, imgCrop = cv2.threshold(imgCrop, 120, 255, cv2.THRESH_TRUNC)
            imgCrop = cv2.convertScaleAbs(imgCrop, alpha=1.4, beta=0)
            kernel = np.array([[-1, -1, -1],
                    [-1,  10, -1],
                    [-1, -1, -1]])
            imgCrop = cv2.filter2D(imgCrop, -1, kernel)
            #cv2.imshow(str(x),imgCrop)
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = "\n".join(lines)
            sht.range("C11").value = r[3]
            for index, line in enumerate(lines, start=12):
                sht.range("C" + str(index)).number_format = "@" 
                sht.range("C" + str(index)).value = line
        elif r[3] == "Số lượng":
            #cv2.imshow(str(x),imgCrop)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("D11").value = r[3]
            for index, row in enumerate(rows, start=12):
                sht.range("D" + str(index)).number_format = "@"
                sht.range("D" + str(index)).value = row[0]
        elif r[3] == "Đơn giá":
            #cv2.imshow(str(x),imgCrop)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("E11").value = r[3]
            for index, row in enumerate(rows, start=12):
                sht.range("E" + str(index)).number_format = "@"
                sht.range("E" + str(index)).value = row[0]
        elif r[3] == "Thành tiền":
            #cv2.imshow(str(x),imgCrop)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("F11").value = r[3]
            for index, row in enumerate(rows, start=12):
                sht.range("F" + str(index)).number_format = "@"
                sht.range("F" + str(index)).value = row[0]
        elif r[2] == "HH":
            #cv2.imshow(str(x),imgCrop)
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6 -c tessedit_char_whitelist=')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            sht.range("B11").value = "Tên hàng hóa, dịch vụ"
            for line in lines:
                table.append(line)
    start_row = 12
    for i in range(len(table)):
        if i == line_count:
            break
        else:
            cell_address = f"B{start_row + i}"
            sht.range(cell_address).value = table[i]
               

    imgShow = image3.copy()
    imgMask = np.zeros_like(imgShow)
    for z,r in enumerate(roi3):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,255,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = image3[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        if r[3] == "Cộng tiền hàng":
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("D6").value = r[3]
            sht.range("E6").number_format = "@"
            sht.range("E6").value = rows
        elif r[3] == "Tiền thuế GTGT":
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("D7").value = r[3]
            sht.range("E7").number_format = "@"
            sht.range("E7").value = rows   
        elif r[3] == "Tổng tiền thanh toán":
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("D8").value = r[3]
            sht.range("E8").number_format = "@"
            sht.range("E8").value = rows 
        elif r[3] == "Số tiền viết bằng chữ":
            imgCrop = cv2.convertScaleAbs(imgCrop, alpha=0.7, beta=0)
            imgCrop = cv2.resize(imgCrop,None,fx=1.1,fy=1.1,interpolation=cv2.INTER_LINEAR)
            #cv2.imshow("...",imgCrop)
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6 -c tessedit_char_whitelist=')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("D9").value = r[3]
            sht.range("E9").value = full_text 
         
    file_name = f"Invoice_LinhKhoa_P_{dem}.xlsx"
    wb.save(fr'SourcecodeOCR\Data_excel_LinhKhoa\{file_name}')  
    excel_file_path = fr"SourcecodeOCR\Data_excel_LinhKhoa\{file_name}" 
    make_columns_bold(excel_file_path, 1,9,1)
    make_columns_bold(excel_file_path, 6,9,4)
    make_rows_bold(excel_file_path, 1, 6, 11)
    wb = xw.Book(excel_file_path)
    change_font_size_and_save(wb, 12)
    process_all_worksheets_linhkhoa(excel_file_path)

def main():
    dem_0 = 0
    dem_1 = 0
    path_compare = 'SourcecodeOCR/Source_Compare_paper'
    path_source = 'SourcecodeOCR/Source_Images'
    orb = cv2.ORB_create(nfeatures = 1000)
    images_compare = []
    images_source = []
    myList_compare = os.listdir(path_compare)
    mylist_source = os.listdir(path_source)
    #print(myList_compare)
    print('Total Invoices Detected',len(myList_compare))
    for img1 in myList_compare:
        images_compare = cv2.imread(f'{path_compare}/{img1}')
        images_compare = cv2.cvtColor(images_compare,cv2.COLOR_BGR2GRAY)
        kp1, des1 = orb.detectAndCompute(images_compare,None)
        matchList = []
        for img2 in mylist_source:
            index_1 = myList_compare.index(img1)
            images_source = cv2.imread(f'{path_source}/{img2}')
            kp2, des2 = orb.detectAndCompute(images_source,None)
            finalVal = -1
            bf = cv2.BFMatcher()
            matches = bf.knnMatch(des1,des2,k=2)
            good = []
            for m,n in matches:
                if m.distance < 0.75*n.distance:
                    good.append([m])  
            print(len(good))
            matchList.append(len(good))
            if len(matchList) != 0:
                if max(matchList) >= 40:
                    finalVal = matchList.index(max(matchList))
                else: 
                    finalVal = -1
        if finalVal == 0: 
            img_path = os.path.join(path_compare, myList_compare[index_1])
            img = cv2.imread(img_path)
            dem_0 += 1
            invoice_cocacola(img,dem_0)
        elif finalVal == 1: 
            img_path = os.path.join(path_compare, myList_compare[index_1])
            img = cv2.imread(img_path)
            dem_1 += 1
            invoice_linhkhoa(img,dem_1)
        else:
            None
    print("Quá trình trích xuất hóa đơn hoàn tất")
if __name__ == "__main__":
    main()
