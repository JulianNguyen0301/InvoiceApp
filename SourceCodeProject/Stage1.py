import cv2
import os
from pdf2image import convert_from_path
from time import sleep,time
from random import *

def main():
    path_compare_INV = 'H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Compare_Invoice'
    path_source_INV = 'H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Images_Invoice'
    orb = cv2.ORB_create(nfeatures = 1000)
    images_compare_INV = []
    images_source_INV = []
    myList_compare_INV = os.listdir(path_compare_INV)
    mylist_source_INV = os.listdir(path_source_INV)
    print('Số lượng hóa đơn được trích xuất:',len(myList_compare_INV))
    for index in range(len(myList_compare_INV)):
        if len(myList_compare_INV) > 0:
            img1 = myList_compare_INV[0]
            images_compare_INV = cv2.imread(f'{path_compare_INV}\{img1}')
            images_compare_INV = cv2.cvtColor(images_compare_INV,cv2.COLOR_BGR2GRAY)
            kp1, des1 = orb.detectAndCompute(images_compare_INV,None)
            matchList = []
            for img2 in mylist_source_INV:
                index_1 = myList_compare_INV.index(img1)
                images_source_INV = cv2.imread(f'{path_source_INV}/{img2}')
                kp2, des2 = orb.detectAndCompute(images_source_INV,None)
                bf = cv2.BFMatcher()
                matches = bf.knnMatch(des1,des2,k=2)
                good = []
                for m,n in matches:
                    if m.distance < 0.75*n.distance:
                        good.append([m])  
                matchList.append(len(good))
            print(matchList)
if __name__ == "__main__":
    main() 