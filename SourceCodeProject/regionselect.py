import cv2
import random

scale = 0.5
circles = []
counter = 0
counter2 = 0
point1 = []
point2 = []
myPoints = []
myColor = []

def mousePoints(event,x,y,flags,params):
    global counter,point1,point2,counter2,circles,myColor
    if event == cv2.EVENT_LBUTTONDOWN:
        if counter == 0:
            point1 = int(x), int(y)
            counter +=1
            myColor = (random.randint(0,2)*200,random.randint(0,2)*200,random.randint(0,2)*200)
        elif counter == 1:
            point2 = int(x), int(y)
            region = input('Enter Region ')
            table = input('Enter table')
            myPoints.append([point1,point2,region,table])
            counter = 0
        circles.append([x,y,myColor])
        counter2 +=1

def main():
    img = cv2.imread(r"H:\APP UNIVERSITY\CODE PYTHON\OpenCV-Master-Computer-Vision-in-Python\SourcecodeOCR\Source_Images_Invoice\Sample_6.png")
    
    while True:
        for x,y,color in circles:
            cv2.circle(img,(x,y),3,color,cv2.FILLED)
        cv2.imshow("Original Image",img)
        cv2.setMouseCallback("Original Image", mousePoints)
        if cv2.waitKey(1) & 0xFF == ord('s'):
            print(myPoints)
            break

if __name__ == "__main__":
    main()

