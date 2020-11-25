""" 
Author: Prashanth Prince
Date : 25 Nov 2020

"""

import cv2
import numpy as np
import pandas as pd
import pytesseract
import win32api
import pyautogui
import time
import math

df = pd.read_excel('dataset_coordinates.xlsx')

left_click_state = win32api.GetKeyState(0x01)  # Left button of Mouse 
right_click_state = win32api.GetKeyState(0x02)  # Right button of Mouse

""" 
 Description: 
 The function "detecting_text" is used to detect the text from the given image. 
 
 Input: Image 
 Output: Text

"""

def detecting_text(image):
    #Converting the image to gray scale 
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) 

    # Performing OTSU threshold 
    # OTSU's threshold technique helps us to do automatic thresholding (automatically determines the threshold)
    ret, thresh1 = cv2.threshold(gray_image, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)

    # Specify structure shape and kernel size. Kernel size increases or decreases the area of the rectangle to be detected.  
    rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (18, 18)) 

    # Applying dilation on the threshold image 
    # Dilation is used for edge detection. Taking the dilation of an image and then subtracting away the original image gives us the edges.
    dilated_image = cv2.dilate(thresh1, rect_kernel, iterations = 1) 
  
    # Finding contours in the image
    contours, hierarchy = cv2.findContours(dilated_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE) 

    # Creating a copy of the image 
    image_copy = image.copy()

    # Looping through the identified contours. Then rectangular part is cropped and passed on to pytesseract for extracting text from it.  
    for cnt in contours: 
        x, y, w, h = cv2.boundingRect(cnt) 

        # Drawing a rectangle on the copied image 
        rect = cv2.rectangle(image_copy, (x, y), (x + w, y + h), (0, 0, 255), 2) 

        # Cropping the text block from the image
        cropped = image_copy[y:y + h, x:x + w] 
    
    # Applying Optical Character Recognition using Tesseract.
    # Setting the config based on the differect working modes of pytesseract.      
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(image, config=custom_config)
    print(text)
    cv2.imshow("Captured texts", image_copy)
    cv2.waitKey(0)

""" 
 Description:
 The function "take_screenshot" is used to take a screenshot at the position of mouse cursor. 
 It takes the page name, position of the label from the user. Then it calculates the distance between the coordinates of the mouse cursor and the position of userid/password text boxes.
 The box with minimum distance from the mouse cursor is selected and then it is cropped from the original screenshot.
 The approach of calculating the minimum distance is carried out to know which text box is selected.

 Input: x and y coordinates of the Mouse Cursor 
 Output: Image

"""

def take_screenshot(x,y):
    # Taking screenshot at the position of mouse pointer.
    image = pyautogui.screenshot()
    page = input("Enter the page you are in: ")
    pos = input("Enter the position of the label(top, bottom, left, right or none): ")

    pageloc2 = [i for i in range(df.index.start, df.index.stop) if df['Page'][i] == page.lower()]
    diff = []

    for i in range(len(pageloc2)):
        diff.append(math.sqrt((x - df.loc[pageloc2[i]].at['left'])**2 + (y - df.loc[pageloc2[i]].at['top'])**2))
    
    image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    image_cropped = image[int(df.loc[pageloc2[diff.index(min(diff))]].at['top']):int(df.loc[pageloc2[diff.index(min(diff))]].at['top'] + df.loc[pageloc2[diff.index(min(diff))]].at['height']), int(df.loc[pageloc2[diff.index(min(diff))]].at['left']):int(df.loc[pageloc2[diff.index(min(diff))]].at['left'] + df.loc[pageloc2[diff.index(min(diff))]].at['width'])]
    cv2.imshow("Captured texts", image_cropped)
    cv2.waitKey(0)

    detecting_text(image_cropped)

try:
    while True: 
        a = win32api.GetKeyState(0x01) 
        b = win32api.GetKeyState(0x02) 
 
        if a != left_click_state or b != right_click_state:  # To check if the Button state has changed 
            left_click_state = a  
            right_click_state = b 
            if a < 0 or b < 0: 
                x, y = pyautogui.position()
                take_screenshot(x,y)    
        time.sleep(0.001)
except KeyboardInterrupt:
    pass

cv2.destroyAllWindows()