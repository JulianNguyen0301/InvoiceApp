import pytesseract 
from pytesseract import Output
import numpy as np
import cv2
import os
import xlwings as xw
from pdf2image import convert_from_path
import re
from datetime import datetime
from selenium import webdriver
from time import sleep,time
from selenium.webdriver.common.by import By
import img2pdf
import glob
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment,Font,numbers,NamedStyle,Color
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl    
from pathlib import Path
import tensorflow as tf
from tensorflow import keras
from random import *
import pandas as pd

pytesseract.pytesseract.tesseract_cmd =r'H:\\APP UNIVERSITY\\CODE PYTHON\\Tesseract-ocr\\tesseract.exe'

img = cv2.imread(r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Compare_Invoice\Pepsico_Page1.png")
#STT
list_stt = []
img_stt = img[1098:2200, 98:150]
#cv2.imshow("img_stt",img_stt)
data_stt = pytesseract.image_to_string(img_stt, lang = 'eng',config= '--oem 3 --psm 6')
lines_data_stt = data_stt.strip().splitlines()
pattern = re.compile(r'^\d+$')
numbers_stt = [element for element in lines_data_stt if pattern.search(element)]
for i in numbers_stt:
    list_stt.append(i)
    count = len(list_stt)

#Tên hàng hóa, dịch vụ
# count_hh1 = count
# img_hh = img[1098:2200, 318:765]
# #cv2.imshow("img_hh",img_hh)
# data_hh = pytesseract.image_to_string(img_hh, lang = 'vie',config= '--oem 3 --psm 6')
# data_hh =  data_hh.replace("NHE","NHF").replace("tỉnh","tinh").replace("muôi","muối")
# lines_data_hh = data_hh.strip().splitlines()
# for index,item in enumerate(lines_data_hh):
#     if len(item) <= 20 and index > 0:
#         lines_data_hh[index-1] = lines_data_hh[index-1] + " " + lines_data_hh[index]
#         lines_data_hh.remove(lines_data_hh[index])
#Đơn vị tính, số lượng, đơn giá, thành tiền
# list_dvt = []
# list_sl = []
# list_dg = []
# list_tt = []
# temp_data = []
# img_complex = img[1098:2200, 766:1570]
# cv2.imshow("img_dvt",img_complex)
# data_dvt = pytesseract.image_to_string(img_complex, lang = 'vie',config= '--oem 3 --psm 6')
# lines_data_dvt = data_dvt.strip().splitlines()
# for i in range(0,len(lines_data_dvt[:count])):
#     temp_data.append(lines_data_dvt[:count][i].split())
#     list_dvt.append(temp_data[i][0])
#     list_sl.append(temp_data[i][1])
#     list_dg.append(temp_data[i][2])
#     list_tt.append(temp_data[i][3])
#Total
img_total = img[1628:2200, 763:1565]
cv2.imshow("img_total",img_total)
data_total = pytesseract.image_to_string(img_total, lang = 'eng',config= '--oem 3 --psm 6')
lines_data_stt = data_total.strip().splitlines()
for item in lines_data_stt:
    if "(Total amount)" in item:
        colon_index = item.index(':')
        untax = item[colon_index + 1:].strip()
    if "(VAT amount)" in item:
        colon_index = item.index(':')
        tax = item[colon_index + 1:].strip()
    if "(Total amount due)" in item:
        colon_index = item.index(':')
        total = item[colon_index + 1:].strip()



cv2.waitKey(0)