import pytesseract 
import numpy as np
import cv2
from pytesseract import Output
import os
import xlwings as xw
from pdf2image import convert_from_path
import re
from datetime import datetime
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from captcha_ocr_library.captcha_ocr import predict_captcha
import sys
import img2pdf
from Excel_Justified.Excel_Justify import Excel_Jusified
pytesseract.pytesseract.tesseract_cmd =r'H:\\APP UNIVERSITY\\CODE PYTHON\\Tesseract-ocr\\tesseract.exe'

def scan(img_source,img_path):
    h,w,c = img_source.shape
    orb = cv2.ORB_create(3000)
    kp1, des1 = orb.detectAndCompute(img_source,None)
    img = cv2.imread(img_path)
    kp2, des2 = orb.detectAndCompute(img,None)
    bf = cv2.BFMatcher(cv2.NORM_HAMMING)
    matches = bf.match(des2,des1)
    matches = list(matches)
    matches.sort(key =lambda x: x.distance)
    good = matches[:int(len(matches)*0.25)]
    imgMatch = cv2.drawMatches(img,kp2,img_source,kp1,good[:750],None,flags=2)
    srcPoints = np.float32([kp2[m.queryIdx].pt for m in good]).reshape(-1,1,2)
    dstPoints = np.float32([kp1[m.trainIdx].pt for m in good]).reshape(-1,1,2)  
    M, _ = cv2.findHomography(srcPoints,dstPoints,cv2.RANSAC, 6.0)
    imgScan = cv2.warpPerspective(img,M,(w,h))
    return imgScan

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

def pdf_to_png():
    pdf_folder = r"H:/APP UNIVERSITY/CODE PYTHON/OpenCV-Master-Computer-Vision-in-Python/SourcecodeOCR/PDF_scan_binary"
    saving_folder = r"H:/APP UNIVERSITY/CODE PYTHON/OpenCV-Master-Computer-Vision-in-Python/SourcecodeOCR/Source_Compare_scan_binary"
    poppler_path = r'H:/OCR/Popler/poppler-23.07.0/Library/bin'
    pdf_files = [file for file in os.listdir(pdf_folder) if file.lower().endswith(".pdf")]
    if not pdf_files:
        print("Không có file PDF trong thư mục. Vui lòng thêm file và chạy lại chương trình.")
    else:
        for pdf_file in os.listdir(pdf_folder):
            if pdf_file.lower().endswith(".pdf"):
                pdf_path = os.path.join(pdf_folder, pdf_file)
                pages = convert_from_path(pdf_path=pdf_path, poppler_path=poppler_path)
                c = 1
                for page in pages:
                    img_name = f"{os.path.splitext(pdf_file)[0]}-page{c}.png"
                    img_path = os.path.join(saving_folder, img_name)
                    if not os.path.exists(img_path):
                        page.save(img_path, "png")
                    else:
                        None
                    c += 1
                os.remove(pdf_path)

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

def process_all_worksheets(excel_file_path):
    wb = xw.Book(excel_file_path)
    for sheet in wb.sheets:
        fit_column_widths_for_one_sheet(sheet)
        sheet.range("A9:F9").api.HorizontalAlignment = -4108  # -4108 tương ứng với giá trị xlCenter trong Excel
        sheet.range("A10:A20").api.HorizontalAlignment = -4108
        sheet.range("A1:B7").api.HorizontalAlignment = -4131
        sheet.range("C10:E20").api.HorizontalAlignment = -4108
        sheet.range("F4:F7").api.HorizontalAlignment = -4152
        sheet.range("F10:F20").api.HorizontalAlignment = -4152
    wb.save()
    wb.close()

def split_text_function(text):
    split_text = text.split("\n")
    split_text = list(filter(lambda item: item.strip(), split_text))  
    result_string = "\n".join(split_text)
    return result_string

def del_space(result):
    indexes_to_remove = []
    for i in range(len(result["conf"])):
        if result["conf"][i] == -1 or result["text"][i] == '' or result["text"][i] == ' ' or  result["text"][i] == '  ' :
            indexes_to_remove.append(i)

    for key, value in result.items():
        result[key] = [value[i] for i in range(len(value)) if i not in indexes_to_remove]

    data = result["text"]
    row_size = 1  # Kích thước của từng dòng
    result1 = [data[i:i + row_size] for i in range(0, len(data), row_size)]
    return result1

def text_return_fulltext(imgCrop):
    imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
    text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
    result_string = split_text_function(text)
    lines = result_string.split("\n")
    full_text = " ".join(lines)
    return full_text

def text_return_lines(imgCrop):
    imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
    text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
    result_string = split_text_function(text)
    lines = result_string.split("\n")
    return lines

def verify_invoice(mst1,ms1,kh1,shd1):
    chrome_options = webdriver.ChromeOptions()  # Đổi tên biến thành chrome_options
    chrome_options.add_argument('--headless')
    web = webdriver.Chrome(options=chrome_options)  # Chỉ sử dụng options=chrome_options
    web.get("https://tracuuhoadon.gdt.gov.vn/search1hd.html")
    sleep(1)
    mst = web.find_element(By.ID,"tin")
    ms = web.find_element(By.ID,"mau")
    kh = web.find_element(By.ID,"kyhieu")
    shd = web.find_element(By.ID,"so")
    btn = web.find_element(By.ID,"searchBtn")
    capcha = web.find_element(By.ID,"captchaCodeVerify")
    sleep(1)
    mst.send_keys(mst1)   
    ms.send_keys(ms1)
    kh.send_keys(kh1)
    shd.send_keys(shd1)
    sleep(1)
    web.get_screenshot_as_file("capcha.png")
    sleep(1)
    img_web = cv2.imread("capcha.png")
    imgcrop = img_web[386:413,378:500]
    #imgcrop = img_web[470:505,504:658]
    cv2.imwrite("Captcha123.png",imgcrop)
    new_image_path = "H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\Captcha123.png"
    predicted_text = predict_captcha(new_image_path)
    capcha.send_keys(predicted_text)
    btn.click()
    sleep(1)
    mXt = web.find_element(By.ID,"lbLoiCode")
    if mXt == "Sai mã xác thực!":
        for i in range(1,10):
            web.get("https://tracuuhoadon.gdt.gov.vn/search1hd.html")
            sleep(1)
            mst = web.find_element(By.ID,"tin")
            ms = web.find_element(By.ID,"mau")
            kh = web.find_element(By.ID,"kyhieu")
            shd = web.find_element(By.ID,"so")
            btn = web.find_element(By.ID,"searchBtn")
            capcha = web.find_element(By.ID,"captchaCodeVerify")
            sleep(1)
            mst.send_keys(mst1)   
            ms.send_keys(ms1)
            kh.send_keys(kh1)
            shd.send_keys(shd1)
            sleep(1)
            web.get_screenshot_as_file("capcha.png")
            sleep(1)
            img_web = cv2.imread("capcha.png")
            imgcrop = img_web[386:413,378:500]
            cv2.imwrite("Captcha123.png",imgcrop)
            new_image_path = "H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\Captcha123.png"
            predicted_text = predict_captcha(new_image_path)
            capcha.send_keys(predicted_text)
            btn.click()
            sleep(1)
            mXt = web.find_element(By.ID,"lbLoiCode")
            if mXt != "Sai mã xác thực!":
                web.get_screenshot_as_file("new_page_screenshot.png")
                sleep(1)
                img_web1 = cv2.imread("new_page_screenshot.png")
                imgcrop1 = img_web1[314:337,293:685]
                cv2.imwrite("Xacthuc.png",imgcrop1)
                text1 = pytesseract.image_to_string(imgcrop1, lang='vie', config='--oem 3 --psm 6')
                text1 = text1.strip().replace('\n', '')
                return text1
    else:        
        web.get_screenshot_as_file("new_page_screenshot.png")
        sleep(1)
        img_web1 = cv2.imread("new_page_screenshot.png")
        imgcrop1 = img_web1[314:337,293:685]
        cv2.imwrite("Xacthuc.png",imgcrop1)
        text1 = pytesseract.image_to_string(imgcrop1, lang='vie', config='--oem 3 --psm 6')
        text1 = text1.strip().replace('\n', '')
        return text1

def pp_ms(img):
    img = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    img = cv2.resize(img,None,fx=1.63,fy=1.63,interpolation=cv2.INTER_BITS) #1.63 
    img = cv2.GaussianBlur(img, (5, 5), 0)
    img = cv2.convertScaleAbs(img, alpha=0.8, beta=0)
    kernel = np.array([[-1, -1, -1],
                    [-1,  9, -1],
                    [-1, -1, -1]])
    img = cv2.filter2D(img, -1, kernel)
    return img

def clear_image_files(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
                None

def convert_png_to_pdf():
    input_folder = "SourcecodeOCR\Source_Compare_scan_binary"
    output_folder = "SourcecodeOCR/PDF_scan_binary"
    png_files = [file for file in os.listdir(input_folder) if file.lower().endswith(".png")]
    if not png_files:
        print("Tất cả các file .pdf hợp lệ.")
    else:
        for png_file in png_files:
            if png_file.lower().endswith(".png"):
                png_path = os.path.join(input_folder, png_file)
                output_pdf_path = os.path.join(output_folder, os.path.splitext(png_file)[0] + ".pdf")
                with open(output_pdf_path, "wb") as pdf_file:
                    pdf_file.write(img2pdf.convert(png_path))

def invoice_cocacola(img,dem,tkh,dc,mst,ms,kh,shd,ngay):
    wb = xw.Book()
    sht = wb.sheets.active
    
    sht.range("A1").value = 'Tên khách hàng'
    sht.range("B1").value = tkh

    sht.range("A2").value = 'Địa chỉ'
    sht.range("B2").number_format = "@"
    sht.range("B2").value = dc

    sht.range("A3").value = 'Mã số thuế'
    sht.range("B3").number_format = "@"
    sht.range("B3").value = mst

    sht.range("A4").value = 'Mẫu số'
    sht.range("B4").number_format = "@"
    sht.range("B4").value = ms

    sht.range("A5").value = 'Ký hiệu'
    sht.range("B5").number_format = "@"
    sht.range("B5").value = kh

    sht.range("A6").value = 'Số hóa đơn'
    sht.range("B6").number_format = "@"
    sht.range("B6").value = shd

    sht.range("A7").value = 'Ngày'
    sht.range("B7").number_format = "@"
    sht.range("B7").value = ngay
    
    
    roi2 = [[(12, 158), (63, 347), 'text', 'STT'], 
            [(70, 155), (741, 347), 'text', 'Tên hàng hóa, dịch vụ'], 
            [(742, 155), (866, 347), 'text', 'Mã hàng'], 
            [(870, 156), (995, 347), 'text', 'Đơn vị tính'], 
            [(1050, 157), (1121, 347), 'text', 'Số lượng'], 
            [(1127, 155), (1245, 347), 'text', 'Đơn giá'], 
            [(1250, 154), (1457, 347), 'text', 'Thành tiền'], 
            [(1175, 354), (1463, 398), 'text', 'Cộng tiền hàng chưa có thuế GTGT'], 
            [(1186, 407), (1463, 452), 'text', 'Tiền thuế GTGT'], 
            [(1181, 464), (1462, 510), 'text', 'Tổng cộng tiền thanh toán'],
            [(323, 518), (1249, 565), 'text', 'Số tiền viết bằng chữ']]
    
    start_row, end_row = 985, 1600
    start_col, end_col = 95, 1560
    half_image2 = img[start_row:end_row, start_col:end_col]
    imgShow = half_image2.copy()
    imgMask = np.zeros_like(imgShow)
    for z,r in enumerate(roi2):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,255,0),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = half_image2[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        if r[3] == 'STT':
            lines = text_return_lines(imgCrop)
            sht.range("A9").value = r[3]
            for index, line in enumerate(lines, start=10):
                line = ''.join(char for char in line if char.isnumeric())
                sht.range("A" + str(index)).number_format = "@" 
                sht.range("A" + str(index)).value = line
        elif r[3] == 'Tên hàng hóa, dịch vụ':
            imgCrop = cv2.resize(imgCrop,None,fx=0.9,fy=0.9,interpolation=cv2.INTER_BITS)
            lines = text_return_lines(imgCrop)
            sht.range("B9").value = r[3]
            for index, line in enumerate(lines, start=10):
                sht.range("B" + str(index)).number_format = "@" 
                sht.range("B" + str(index)).value = line
        elif r[3] == 'Đơn vị tính':
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang = 'eng', config= 'tessdata',output_type= Output.DICT)
            result_2_1_cleaned = [sublist for sublist in result_2_1['text'] if not all(item == '' for item in sublist)]
            sht.range("C9").value = r[3]
            for index, line in enumerate(result_2_1_cleaned, start=10):
                sht.range("C" + str(index)).value = line
        elif r[3] == 'Số lượng':
            lines = text_return_lines(imgCrop)
            sht.range("D9").value = r[3]
            for index, line in enumerate(lines, start=10):
                sht.range("D" + str(index)).number_format = "@" 
                sht.range("D" + str(index)).value = line
            cv2.waitKey(0)
        elif r[3] == 'Đơn giá':
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang = 'eng', config= 'tessdata',output_type= Output.DICT)
            result_2_1_cleaned = [sublist for sublist in result_2_1['text'] if not all(item == '' for item in sublist)]
            sht.range("E9").value = r[3]
            for index, line in enumerate(result_2_1_cleaned, start=10):
                sht.range("E" + str(index)).value = line
        elif r[3] == 'Thành tiền':
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang = 'eng', config= 'tessdata',output_type= Output.DICT)
            result_2_1_cleaned = [sublist for sublist in result_2_1['text'] if not all(item == '' for item in sublist)]
            sht.range("F9").value = r[3]
            for index, line in enumerate(result_2_1_cleaned, start=10):
                sht.range("F" + str(index)).number_format = "#,##0"
                text_value = line
                numeric_value = float(text_value.replace(',', ''))
                sht.range("F" + str(index)).value = numeric_value
        elif r[3] == 'Cộng tiền hàng chưa có thuế GTGT':
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang = 'eng', config= 'tessdata',output_type= Output.DICT)
            result_2_1_cleaned = [sublist for sublist in result_2_1['text'] if not all(item == '' for item in sublist)]
            sht.range("E4").value = r[3]
            sht.range("F4").number_format = "#,##0"
            numeric_value = [float(text.replace(',', '')) for text in result_2_1_cleaned]
            sht.range("F4").value = numeric_value
        elif r[3] == 'Tiền thuế GTGT':
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang = 'eng', config= 'tessdata',output_type= Output.DICT)
            result_2_1_cleaned = [sublist for sublist in result_2_1['text'] if not all(item == '' for item in sublist)]
            sht.range("E5").value = r[3]
            sht.range("F5").number_format = "#,##0"
            numeric_value = [float(text.replace(',', '')) for text in result_2_1_cleaned]
            sht.range("F5").value = numeric_value
        elif r[3] == 'Tổng cộng tiền thanh toán':
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_BITS)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang = 'eng', config= 'tessdata',output_type= Output.DICT)
            result_2_1_cleaned = [sublist for sublist in result_2_1['text'] if not all(item == '' for item in sublist)]
            sht.range("E6").value = r[3]
            sht.range("F6").number_format = "#,##0"
            numeric_value = [float(text.replace(',', '')) for text in result_2_1_cleaned]
            sht.range("F6").value = numeric_value
        if r[3] == 'Số tiền viết bằng chữ':
            imgCrop = cv2.cvtColor(imgCrop,cv2.COLOR_BGR2GRAY)
            imgCrop = cv2.resize(imgCrop,None,fx=1.3,fy=1.3,interpolation=cv2.INTER_BITS)
            _, imgCrop = cv2.threshold(imgCrop, 128, 255, cv2.THRESH_BINARY)
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            full_text = " ".join(lines)
            sht.range("E7").value = r[3]
            sht.range("F7").value = full_text
    
    file_name = f"Cocacola_B_{dem}.xlsx"
    wb.save(fr'SourcecodeOCR\Data_Excel\{file_name}')
    excel_file_path = fr"SourcecodeOCR\Data_Excel\{file_name}"  
    make_columns_bold(excel_file_path, 1,7,1)
    make_columns_bold(excel_file_path, 4,7,5)
    make_rows_bold(excel_file_path, 1, 6, 9)
    wb = xw.Book(excel_file_path)
    change_font_size_and_save(wb, 12)
    process_all_worksheets(excel_file_path)
    
def invoice_linhkhoa(img,dem,tkh,dc,mst,ms,kh,shd,ngay):
    wb = xw.Book()
    sht = wb.sheets.active
    
    sht.range("A1").value = 'Tên khách hàng'
    sht.range("B1").value = tkh

    sht.range("A2").value = 'Địa chỉ'
    sht.range("B2").number_format = "@"
    sht.range("B2").value = dc

    sht.range("A3").value = 'Mã số thuế'
    sht.range("B3").number_format = "@"
    sht.range("B3").value = mst

    sht.range("A4").value = 'Mẫu số'
    sht.range("B4").number_format = "@"
    sht.range("B4").value = ms

    sht.range("A5").value = 'Ký hiệu'
    sht.range("B5").number_format = "@"
    sht.range("B5").value = kh

    sht.range("A6").value = 'Số hóa đơn'
    sht.range("B6").number_format = "@"
    sht.range("B6").value = shd

    sht.range("A7").value = 'Ngày'
    sht.range("B7").number_format = "@"
    sht.range("B7").value = ngay
        
    start_row, end_row = 801 , 1474
    start_col, end_col = 78, 1573
    image2 = img[start_row:end_row, start_col:end_col]
    #image2 = cv2.cvtColor(image2,cv2.COLOR_BGR2GRAY)
    roi2 = [[(7, 98), (96, 665), 'text', 'STT'],
            [(101, 100), (557, 143), 'HH', '1'],
            [(101, 147), (557, 192), 'HH', '2'], 
            [(101, 194), (557, 237), 'HH', '3'], 
            [(101, 240), (557, 285), 'HH', '4'], 
            [(101, 292), (557, 333), 'HH', '5'], 
            [(101, 340), (557, 380), 'HH', '6'], 
            [(101, 387), (557, 428), 'HH', '7'], 
            [(101, 435), (557, 477), 'HH', '8'], 
            [(101, 483), (557, 521), 'HH', '9'], 
            [(101, 532), (557, 571), 'HH', '10'], 
            [(101, 581), (557, 617), 'HH', '11'], 
            [(101, 627), (557, 666), 'HH', '12'],
            [(701, 105), (752, 664), 'text', 'Đơn vị tính'],
            [(941, 104), (1035, 665), 'text', 'Số lượng'], 
            [(1054, 101), (1245, 667), 'text', 'Đơn giá'], 
            [(1303, 101), (1486, 666), 'text', 'Thành tiền']]
    imgShow = image2.copy()
    imgMask = np.zeros_like(imgShow)
    table = []
    line_count = 0
    for x,r in enumerate(roi2):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,0,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = image2[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        if r[3] == 'STT':
            _, imgCrop = cv2.threshold(imgCrop, 86, 255, cv2.THRESH_BINARY)
            imgCrop = cv2.resize(imgCrop,None,fx=1.3,fy=1.3,interpolation=cv2.INTER_LINEAR)
            result = pytesseract.image_to_data(imgCrop, lang = 'eng', config= '--oem 3 --psm 6 -c preserve_interword_spaces=1 -c language_model_penalty_non_dict_word=1 -c language_model_penalty_non_freq_dict_word=1 -c language_model_penalty_dict_non_word=1',output_type= Output.DICT) #,output_type= Output.DICT
            filtered_result = [item for item in result['text'] if item]
            count_list = [item for item in filtered_result if item.isnumeric()]
            sht.range("A9").value = r[3]          
            for index, line in enumerate(count_list, start=10):
                sht.range("A" + str(index)).number_format = "@" 
                sht.range("A" + str(index)).value = line
                if line.strip():
                    line_count += 1

        elif r[3] == "Đơn vị tính":
            _, imgCrop = cv2.threshold(imgCrop, 77, 255, cv2.THRESH_BINARY)
            imgCrop = cv2.resize(imgCrop,None,fx=1.3,fy=1.3,interpolation=cv2.INTER_LINEAR)
            imgCrop = cv2.convertScaleAbs(imgCrop, alpha=1.67, beta=0)
            kernel = np.array([[-1, -1, -1],
                    [-1,  10, -1],
                    [-1, -1, -1]])
            imgCrop = cv2.filter2D(imgCrop, -1, kernel)
            text = pytesseract.image_to_string(imgCrop, lang = 'vie+eng',config= '--oem 3 --psm 6')
            lines = text.splitlines()
            sht.range("C9").value = r[3]
            for i in range(len(lines)):
                if lines[i].startswith('k'):
                    lines[i] = 'kg'
            for index, line in enumerate(lines, start=10):
                sht.range("C" + str(index)).number_format = "@" 
                sht.range("C" + str(index)).value = line
        elif r[3] == "Số lượng":
            _, imgCrop = cv2.threshold(imgCrop, 128, 255, cv2.THRESH_BINARY) 
            imgCrop = cv2.convertScaleAbs(imgCrop, alpha=0.81, beta=0) 
            imgCrop = cv2.resize(imgCrop,None,fx=0.95,fy=0.95,interpolation=cv2.INTER_BITS) 
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            decimal_pattern = re.compile(r'\d+,\d+')
            filtered_data = [item for item in rows if decimal_pattern.match(item[0])]
            sht.range("D9").value = r[3]
            for index, row in enumerate(filtered_data, start=10):
                sht.range("D" + str(index)).number_format = "@"
                sht.range("D" + str(index)).value = row
        elif r[3] == "Đơn giá":
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            sht.range("E9").value = r[3]
            for index, row in enumerate(rows, start=10):
                sht.range("E" + str(index)).number_format = "@"
                sht.range("E" + str(index)).value = row
        elif r[3] == "Thành tiền":
            _, imgCrop = cv2.threshold(imgCrop, 129, 255, cv2.THRESH_TRUNC) 
            imgCrop = cv2.convertScaleAbs(imgCrop, alpha=1.6, beta=0) 
            imgCrop = cv2.GaussianBlur(imgCrop, (3, 3), 0)
            result_2_1 = pytesseract.image_to_data(imgCrop, lang='eng', config='tessdata', output_type=Output.DICT)
            rows = del_space(result_2_1)
            output_list = [[item.replace(',', '.').replace(' ', '') for item in sublist] for sublist in rows]
            decimal_pattern = re.compile(r'\d+.\d+|\d+.\d+.\d')
            filtered_data = [item for item in output_list if decimal_pattern.match(item[0])]
            sht.range("F9").value = r[3]
            for index, row in enumerate(filtered_data, start=10):
                sht.range("F" + str(index)).number_format = "#,##0"
                numeric_value = [float(item.replace('.', '').replace(',', '')) for item in row]
                sht.range("F" + str(index)).value = numeric_value
        elif r[2] == "HH":
            _, imgCrop = cv2.threshold(imgCrop, 128, 255, cv2.THRESH_BINARY)
            imgCrop = cv2.resize(imgCrop,None,fx=0.85,fy=0.85,interpolation=cv2.INTER_LINEAR)
            text = pytesseract.image_to_string(imgCrop, lang='vie', config='--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")
            for i in range(len(lines)):
                lines[i] = lines[i].replace("NÁM","NẤM")
            sht.range("B9").value = "Tên hàng hóa, dịch vụ"
            for line in lines:
                table.append(line)
    start_row = 10            
    for i,row in enumerate(table):
            if index == line_count:
                break
            else:
                cell_address = f"B{start_row + i}"
                sht.range(cell_address).value = table[i]

    start_row, end_row = 1470 , 1665
    start_col, end_col = 78, 1573
    image3 = img[start_row:end_row, start_col:end_col]
    image3 = cv2.cvtColor(image3,cv2.COLOR_BGR2GRAY)
    
    roi3 = [ [(247, 145), (1236, 190), 'text', 'Số tiền viết bằng chữ']]

    result_3 = pytesseract.image_to_data(image3, lang='eng', config='tessdata', output_type=Output.DICT) #, output_type=Output.DICT
    rows = del_space(result_3)
    selected_numbers = [item for sublist in rows for item in sublist if isinstance(item, str) and item.count('.') >=1  and item.replace('.', '').isdigit()]
    formatted_numbers = [number.replace('.', ',') for number in selected_numbers]
    sht.range("E4").value = "Cộng tiền hàng"
    sht.range("E5").value = "Tiền thuế GTGT"
    sht.range("E6").value = "Tổng tiền thanh toán"
    for index, line in enumerate(formatted_numbers, start=4):
        cell = sht.range("F" + str(index))
        cell.number_format = "#,##0" 
        cell.value = float(line.replace(",", ""))

    imgShow = image3.copy()
    imgMask = np.zeros_like(imgShow)
    for z,r in enumerate(roi3):
        cv2.rectangle(imgMask, ((r[0][0]),r[0][1]),((r[1][0]),r[1][1]),(0,255,255),cv2.FILLED)
        imgShow = cv2.addWeighted(imgShow,0.99,imgMask,0.1,0)
        imgCrop = image3[r[0][1]:r[1][1],r[0][0]:r[1][0]]
        if r[3] == 'Số tiền viết bằng chữ':
            imgCrop = cv2.GaussianBlur(imgCrop, (3, 3), 0)
            imgCrop = cv2.convertScaleAbs(imgCrop, alpha=1.1, beta=0)
            imgCrop = cv2.resize(imgCrop,None,fx=1.2,fy=1.2,interpolation=cv2.INTER_LINEAR_EXACT)
            text = pytesseract.image_to_string(imgCrop, lang = 'vie',config= '--oem 3 --psm 6')
            result_string = split_text_function(text)
            lines = result_string.split("\n")   
            full_text = " ".join(lines)
            full_text = full_text.replace('muơi', 'mươi').replace('bồn','bốn')
            sht.range("E7").value = r[3]
            sht.range("F7").value = full_text 
        
    file_name = f"LinhKhoa_B_{dem}.xlsx"
    wb.save(fr'SourcecodeOCR\Data_Excel\{file_name}')  
    excel_file_path = fr"SourcecodeOCR\Data_Excel\{file_name}" 
    make_columns_bold(excel_file_path, 1,7,1)
    make_columns_bold(excel_file_path, 4,7,5)
    make_rows_bold(excel_file_path, 1, 6, 9)
    wb = xw.Book(excel_file_path)
    change_font_size_and_save(wb, 12)
    process_all_worksheets(excel_file_path)
    
def main():
    pdf_to_png()
    verify_executed_Cocacola = False
    verify_executed_LinhKhoa = False
    path_compare = 'SourcecodeOCR\Source_Compare_scan_binary'
    path_source = 'SourcecodeOCR/Source_Images'
    orb = cv2.ORB_create(nfeatures = 1000)
    images_compare = []
    images_source = []
    myList_compare = os.listdir(path_compare)
    mylist_source = os.listdir(path_source)
    print('Số lượng hóa đơn được trích xuất:',len(myList_compare))
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
            matchList.append(len(good))
            for i, num_matches in enumerate(matchList):
                if num_matches > 60:
                    finalVal = i
                    break
        if finalVal == 0:
            img_path = os.path.join(path_compare, myList_compare[index_1])
            img_source = cv2.imread("H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Images\Invoice_Cocacola_Ori.png")
            imgScan = scan(img_source,img_path)
            ngay_gio_thuc_te = datetime.now()
            dem_0 = ngay_gio_thuc_te.strftime("%Y-%m-%d_%H.%M.%S")
            start_row, end_row = 45,430
            start_col, end_col = 95, 1560
            half_image1 = imgScan[start_row:end_row, start_col:end_col]
            half_image2 = half_image1[235:330,1011:1455]
            data2 = pytesseract.image_to_data(half_image2, lang = 'eng', config= 'tessdata',output_type= Output.DICT) #,output_type= Output.DICT
            pattern = re.compile(r'\d+')
            elements_with_numbers = [element for element in data2['text'] if pattern.search(element)]
            data1 = pytesseract.image_to_data(half_image1, lang = 'eng', config= 'tessdata',output_type= Output.DICT) #,output_type= Output.DICT
            string1 =  pytesseract.image_to_string(half_image1, lang = 'vie',config= '--oem 3 --psm 6 preserve_interword_spaces=1')
            number_pattern = re.compile(r'^\d+$')
            numbers_only = [element for element in data1["text"] if number_pattern.match(element)]
            date_obj = datetime.strptime(numbers_only[-1], '%d%m%Y')
            formatted_date = date_obj.strftime('%d/%m/%Y')
            lines = string1.strip().splitlines()
            number_elements = [element for element in data1["text"] if re.match(r'^\d+$', element) and len(element) >= 5]
            while True:
                    if not verify_executed_Cocacola:  
                        text_verify = verify_invoice(number_elements[0],number_elements[1],elements_with_numbers[1],elements_with_numbers[0])
                        verify_executed_Cocacola = True  
                    else:
                        break
                    if text_verify == "NNT tạm nghỉ kinh doanh có thời hạn":
                        print("Công ty Cocacola đã thông báo về việc tạm ngừng hoạt động có thời hạn và được cơ quan có thẩm quyền chấp thuận")
                        sys.exit()
                    elif text_verify == "NNT ngừng hoạt động nhưng chưa hoàn thành thủ tục đóng MST":
                        print("Doanh nghiệp Cocacola không hoàn thành các nghĩa vụ thuế")
                        sys.exit()
                    elif text_verify == "NNT không hoạt động tại địa chỉ đã đăng ký":
                        print("Công ty Cocacola đang tra cứu đã bị cơ quan thuế quản lý khóa mã số thuế do doanh nghiệp không hoạt động tại địa điểm như đã đăng ký trên Giấy chứng nhận đăng ký kinh doanh")
                        sys.exit()
                    elif text_verify == "NNT đang hoạt động (đã được cấp GCN ĐKT)":
                        print("Nhà cung cấp Cocacola hợp lệ")
                        break
            os.remove(img_path)
            invoice_cocacola(imgScan, dem_0,lines[0],lines[2],number_elements[0],number_elements[1],elements_with_numbers[1],elements_with_numbers[0],formatted_date)
        elif finalVal == 1: 
            img_path = os.path.join(path_compare, myList_compare[index_1])
            img_source = cv2.imread("SourcecodeOCR/Source_Images/Invoice_LinhKhoa_Ori.png")
            imgScan = scan(img_source,img_path)
            ngay_gio_thuc_te = datetime.now()
            dem_1 = ngay_gio_thuc_te.strftime("%Y-%m-%d_%H.%M.%S")
            start_row, end_row = 95 , 412
            start_col, end_col = 75, 1530
            image1 = imgScan[start_row:end_row, start_col:end_col]
            image2 = image1[173:316,5:1450]
            img1 = image1[100:146,770:1062]   
            img1 = pp_ms(img1)
            text11 = pytesseract.image_to_string(img1,lang= 'eng',config= '--psm 6')
            text11 = text11.strip().replace('\n', '')
            img1 = image1[109:140,542:776]  
            img1 = pp_ms(img1)
            text21= pytesseract.image_to_string(img1,lang= 'eng',config= '--psm 6')
            text21 = text21.strip().replace('\n', '')
            text = text21 + text11
            
            string1 =  pytesseract.image_to_string(image2, lang = 'vie',config= '--oem 3 --psm 6 preserve_interword_spaces=1')
            lines = string1.strip().splitlines()
            words = lines[2].split()
            index_of_colon = words.index('chỉ:')
            text2 = ' '.join(words[index_of_colon + 1:])
            text1 = lines[0].replace('PHÁM', 'PHẨM')
            data1 = pytesseract.image_to_data(image1, lang = 'vie+eng', config= 'tessdata',output_type= Output.DICT)
            pattern = re.compile(r'\d+')
            elements_with_numbers = [element for element in data1['text'] if pattern.search(element)]
            elements_to_merge = elements_with_numbers[1:4] 
            merged_element = '/'.join(elements_to_merge)
            while True:
                    if not verify_executed_LinhKhoa:  
                        text_verify = verify_invoice(elements_with_numbers[6],text,elements_with_numbers[0],elements_with_numbers[4])
                        verify_executed_LinhKhoa = True  
                    else:
                        break
                    if text_verify == "NNT tạm nghỉ kinh doanh có thời hạn":
                        print("Công ty Linh Khoa đã thông báo về việc tạm ngừng hoạt động có thời hạn và được cơ quan có thẩm quyền chấp thuận")
                        sys.exit()
                    elif text_verify == "NNT ngừng hoạt động nhưng chưa hoàn thành thủ tục đóng MST":
                        print("Doanh nghiệp Linh Khoa không hoàn thành các nghĩa vụ thuế")
                        sys.exit()
                    elif text_verify == "NNT không hoạt động tại địa chỉ đã đăng ký":
                        print("Công ty Linh Khoa đang tra cứu đã bị cơ quan thuế quản lý khóa mã số thuế do doanh nghiệp không hoạt động tại địa điểm như đã đăng ký trên Giấy chứng nhận đăng ký kinh doanh")
                        sys.exit()
                    elif text_verify == "NNT đang hoạt động (đã được cấp GCN ĐKT)":
                        print("Nhà cung cấp Linh Khoa hợp lệ")
                        break
            os.remove(img_path)
            invoice_linhkhoa(imgScan,dem_1,text1,text2,elements_with_numbers[6],text,elements_with_numbers[0],elements_with_numbers[4],merged_element)
        else:
            None
    convert_png_to_pdf()
    folder_path = r"SourcecodeOCR\Source_Compare_scan_binary"
    clear_image_files(folder_path)
    file_count = len(os.listdir("SourcecodeOCR\PDF_scan_binary"))
    print(f"Số lượng hóa đơn không thể trích xuất là: {file_count}")
    for filename in os.listdir("SourcecodeOCR\PDF_scan_binary"):
        print(filename)
    Excel_Jusified()
    print("Quá trình trích xuất hóa đơn hoàn tất")
if __name__ == "__main__":
    main()