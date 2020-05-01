# -*- coding: utf-8 -*-
"""
Created on Thu Apr 23 22:40:15 2020

@author: nandini bansal
"""

# pytesseract is for OCR from images
# Check PyPi Project - https://pypi.org/project/pytesseract/
import pytesseract

# OpenCV for pre-processing images before feeding to the tesseract OCR.
# Check PyPi Project - https://pypi.org/project/opencv-python/
import cv2

# Converting every page of PDF into images for OCR
# Check PyPi Project - https://pypi.org/project/pdf2image/
import pdf2image

# For reading tabular data from pdf in csv/json/tsv format
# Check PyPi Project - https://pypi.org/project/tabula-py/ 
import tabula

# A Python binding for MuPDF - a lightweight PDF and XPS viewer
# MuPDF can access files in PDF, XPS, OpenXPS, epub, comic and fiction book formats, and it is known for both, its top performance and high rendering quality.
# Check PyPi Project - https://pypi.org/project/PyMuPDF/ (previously known as py-fitz)
import fitz

import os
import pandas as pd
import numpy as np
import re

###### Global Variables ######

# Including the path to the tesseract.exe for PyTesseract to work
pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'

## BASE_DIR and pdf_file can also be stored in a JSON file
BASE_DIR = os.getcwd()
pdf_name = 'bank.pdf'

file_name = pdf_name.split('.')[0]

# PyTesseract Config Options
# language = English
# psm or Page segmentation modes = 1 represents Automatic page segmentation with Orientation and Script Detection (OSD).
# oem or Engine Mode - Nueral Nets LSTM engine only
config = ('-l eng --oem 1 --psm 1')

################################

def range_split(t):
    range_values = t.split(' from ')[-1]
    if len(range_values.split(' to ')) == 2:
        first_value, last_value = range_values.split(' to ')
    else:
        first_value = "NA"
        last_value = "NA"
        
    if first_value == " ":
        first_value = "NA"
    if last_value == " ":
        last_value = "NA"
    return first_value, last_value

def extract_text_from_pdf():
    # Converting the pages of PDF file into images 
    images = pdf2image.convert_from_path(os.path.join(BASE_DIR, pdf_name))
    images_folder = os.path.join(BASE_DIR, file_name)
    
    # Creating a folder with the name of pdf file for storing images
    if not os.path.exists(images_folder):
        os.mkdir(images_folder)
    
    # Saving images
    pg_cnt = 1
    for image in images:
        image.save(os.path.join(images_folder, file_name+'_'+str(pg_cnt)+'.png'), 'PNG')
    # Using PyTesseract, the text is extracted from the images.
    # Before feeding the image file to the "image_to_string" function
    # we have perfomed certain pre-processing of the image file to increase the readability of the text
    # Image pre-processing helps in removing the noise and enhance the quality of image for extraction of the text
    
    ## We can also use direct pdf to text conversion libraries of Python for extracting text.
    
    for file in os.listdir(images_folder):
        if file.endswith('.png') and file_name in file:
            # Opening the image file in Gray Scale Mode
            img = cv2.imread(os.path.join(images_folder, file), 0)
            # Performing Binary Thresholding and Erosion to enhance the image quality
            r, threshold = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY)
            kernel = np.ones((2,2),np.uint8)
            erosion = cv2.erode(threshold,kernel,iterations = 1)
            # Extracting the text
            text = pytesseract.image_to_string(erosion, config=config)
    
    date_flag = False
    amount_flag = False
    # Parsing the text extracted for relevant information
    for t in text.split('\n'):
        if 'Account Number' in t:
            acc_number = re.findall('\d+',t)[0]
            owner_name = t.split(' - ')[-1]
            firstname, lastname = owner_name.split(' ')
            
        elif 'Transaction Date' in t and not date_flag:
            first_date, last_date = range_split(t)
            date_flag = True
            
        elif 'Amount' in t and not amount_flag:
            amount_start, amount_last = range_split(t)
            amount_flag = True
            
        elif 'Cheque number' in t:
            cheque_start, cheque_last = range_split(t)
    return acc_number, firstname, lastname, first_date, last_date, amount_start, amount_last, cheque_start, cheque_last
    
def extract_image_text():
    
    # From the pdf, we ares searching for the Image elements before extracting them and saving their PNG files.
    # Text can be extracted from these "PNG" image files using PyTesseract module
    
    doc = fitz.open(os.path.join(BASE_DIR, pdf_name))
    image_list = []
    for i in range(len(doc)):
        for img in doc.getPageImageList(i):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            if pix.n < 5: # this is GRAY or RGB
                pix.writePNG("p%s-%s.png" % (i, xref))
                image_list.append("p%s-%s.png" % (i, xref))
            else: # CMYK: convert to RGB first
                pix1 = fitz.Pixmap(fitz.csRGB, pix)
                pix1.writePNG("p%s-%s.png" % (i, xref))
                image_list.append("p%s-%s.png" % (i, xref))
                pix1 = None
            pix = None
    
    text_from_images = []
    for image in image_list:        
        bank_name = pytesseract.image_to_string(cv2.imread(image))
        text_from_images.append(bank_name)
        
    return text_from_images

def extract_table_from_pdf():    
    # Extracting Tabular Information from the pdf
    # Tabula can extract tables with or without borders - It is wrapper built over tabula-java
    # To run tabula-py, JAVA runtime set up is a pre-requisite  
    tabula.convert_into(os.path.join(BASE_DIR, pdf_name), os.path.join(BASE_DIR, file_name+".csv"), output_format="csv")
    table = pd.read_csv(os.path.join(BASE_DIR, file_name+".csv"))
    return table

def ocr():
    acc_number, firstname, lastname, first_date, last_date, amount_start, amount_last, cheque_start, cheque_last = extract_text_from_pdf()
    text_from_images = extract_image_text()
    table = extract_table_from_pdf()
    
    print("Bank Name ", text_from_images[0])
    print()
    print("Account Number ", acc_number)
    print("Account Holder Information")
    print("First Name: ", firstname)
    print("Last Name: ", lastname)
    print()
    print("Transaction Information")
    print("Transaction Dates from {} to {}".format(first_date, last_date))
    print("Transaction Amount from {} to {}".format(amount_start, amount_last))
    print("Cheque Number from {} to {}".format(cheque_start, cheque_last))
    print(table.head())


if __name__ == '__main__':
	ocr()