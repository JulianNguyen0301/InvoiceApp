import pytesseract 
from pytesseract import Output
import numpy as np
import cv2
from random import *
import re
pytesseract.pytesseract.tesseract_cmd =r'H:\\APP UNIVERSITY\\CODE PYTHON\\Tesseract-ocr\\tesseract.exe'

#invoice_cocacola(img,supplier,mst1,address1,consumer,mst2,address2,ms,kh,so,ngaygiao,ngayki):
# img = cv2.imread(r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Compare_Invoice\Pepsico_Page1.png")
# img_gray = cv2.cvtColor(img,cv2.COLOR_RGB2GRAY)
# img1 = img_gray[116:303,388:1235]
# #cv2.imshow("img1",img1)
# string1 =  pytesseract.image_to_string(img1, lang = 'vie',config= '--oem 3 --psm 6 ')
# lines1 = string1.strip().splitlines()
# data_cleaned1 = [item for item in lines1 if item != '']
# #print(data_cleaned1)
# supplier = data_cleaned1[0] + ' '+ data_cleaned1[1]
# colon_index = data_cleaned1[2].index(':')
# mst1 = data_cleaned1[2][colon_index + 1:].strip()
# colon_index = data_cleaned1[3].index(':')
# address1 = data_cleaned1[3][colon_index + 1:].strip()
# address1 = address1 + " " + data_cleaned1[4]
# address1 = address1.replace("Sô","Số")

# img2 = img_gray[450:671,93:1580]
# #cv2.imshow("img2",img2)
# string2 =  pytesseract.image_to_string(img2, lang = 'vie',config= '--oem 3 --psm 6 ')
# lines2 = string2.strip().splitlines()
# #print(lines2)
# data2 = []
# for line in lines2:
#     if ':' in line:
#         colon_index = line.index(':')
#         data = line[colon_index + 1:].strip()
#         data2.append(data)
# #print(data2)
# consumer = data2[1]
# address2 = data2[2]
# mst2 = data2[3]
# img3 = img_gray[116:250,1235:1520]
# #cv2.imshow("img3",img3)
# string3 =  pytesseract.image_to_string(img3, lang = 'eng',config= '--oem 3 --psm 6 ')
# lines3 = string3.strip().splitlines()
# data3 = []
# for i in range(0,len(lines3)):
#     colon_index = lines3[i].index(':')
#     data = lines3[i][colon_index + 1:].strip()
#     data3.append(data)
# ms = data2[0]
# kh = data3[0]
# so = data3[1]
# ngaygiao = data3[2]

# img4 = img_gray[1500:2200,1120:1540]
# cv2.imshow("img4",img4)
# string4 =  pytesseract.image_to_string(img4, lang = 'eng',config= '--oem 3 --psm 6 ')
# lines4 = string4.strip().splitlines()
# colon_index = lines4[-1].index(':')
# data4 = lines4[-1][colon_index + 1:].strip()
# ngayki = data4
# print(f"Supplier: {supplier}\nMST1: {mst1}\nAddress1: {address1}\nConsumer: {consumer}\nMST2: {mst2}\nAddress2: {address2}\nMS: {ms}\nKH: {kh}\nSo: {so}\nNgaygiao: {ngaygiao}\nNgayki: {ngayki}")

# cv2.waitKey(0)
img = cv2.imread(r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Images_Invoice\Sample_6.png")
img[579:745,1379:1540] = 255
img1 = img[270:791, 66:1534]
#cv2.imshow("img1",img1)
string1 =  pytesseract.image_to_string(img1, lang = 'vie',config= '--oem 3 --psm 6 preserve_interword_spaces=1')
lines1 = string1.strip().splitlines()
data_cleaned1 = [item for item in lines1 if item != '']
supplier = data_cleaned1[0]

colon_index = data_cleaned1[1].index(':')
mst1 = data_cleaned1[1][colon_index + 1:].strip()

colon_index = data_cleaned1[2].index(':')
address1 = data_cleaned1[2][colon_index + 1:].strip()

colon_index = data_cleaned1[7].index(':')
consumer = data_cleaned1[7][colon_index + 1:].strip()

colon_index = data_cleaned1[8].index(':')
mst2 = data_cleaned1[8][colon_index + 1:].strip()

colon_index = data_cleaned1[9].index(':')
address2 = data_cleaned1[9][colon_index + 1:].strip()

img2 = img[70:200,300:1600]
string2 =  pytesseract.image_to_string(img2, lang = 'eng',config= '--oem 3 --psm 6 preserve_interword_spaces=1')
lines2 = string2.strip().splitlines()
colon_index = lines2[0].index(':')
kh = lines2[0][colon_index + 1:].strip()

temp_lines2 = lines2[1].split()
ngaygiao = temp_lines2[1] + "/" + temp_lines2[3] + "/" + temp_lines2[5]

colon_index = lines2[1].index(':')
so = lines2[1][colon_index + 1:].strip()

img3 = img[190:247,614:1135]
img3 = cv2.cvtColor(img3,cv2.COLOR_BGR2GRAY)
img3 = cv2.resize(img3,None,fx=1.24,fy=1.24,interpolation=cv2.INTER_BITS)
gamma = 2
img3 = np.uint8(np.power(img3 / float(np.max(img3)), gamma) * 255)
ms = pytesseract.image_to_string(img3,lang= "vie", config='--oem 3 --psm 7')
ms = ms.replace("&","").replace("O","0").replace("I","1").replace("Q","0").replace("68B","6B").replace("383","3B3").replace("984D","98AD").replace("340","3A0").replace("085A","08BA").replace("946","9A6").replace("984","98A").replace("346","3A6")
ms = re.sub(r'[^a-fA-F0-9]', '', ms)
if not ms.startswith("00"):
    ms = "0" + ms

img4 = img[1500:2120,945:1450]
string4 =  pytesseract.image_to_string(img4, lang = 'eng',config= '--oem 3 --psm 6 preserve_interword_spaces=1')
lines4 = string4.strip().splitlines()
colon_index = lines4[-1].index(':')
ngayki = lines4[-1][colon_index + 1:].strip()
#print(f"Supplier: {supplier}\nMST1: {mst1}\nAddress1: {address1}\nConsumer: {consumer}\nMST2: {mst2}\nAddress2: {address2}\nMS: {ms}\nKH: {kh}\nSo: {so}\nNgaygiao: {ngaygiao}\nNgayki: {ngayki}")
cv2.waitKey(0)
