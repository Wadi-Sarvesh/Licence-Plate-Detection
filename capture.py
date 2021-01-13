from PIL import Image
import numpy as np
import pytesseract
import re
import argparse
import cv2
import os
import openpyxl
import xlrd
from xlutils.copy import copy
from datetime import datetime
from datetime import date
import  imutils

cap = cv2.VideoCapture('sample_video9.mp4')
img_counter = 0
while True:
    ret, frame = cap.read()
    if not ret:
        break
    k = cv2.waitKey(25)
    cv2.imshow("test",frame)

    if k%256 == 27:
        # ESC pressed
        print("Escape hit, closing...")
        break
    elif k%256 == 32:
        # SPACE pressed
        img_name = "opencv_frame_{}.png".format(img_counter)
        cv2.imwrite(img_name, frame)
        print("{} written!".format(img_name))
        img_counter += 1
       
cap.release()

cv2.destroyAllWindows()




