# main.ipynb
#
# Uses relative directories "input/Check_Copy" and "input/TC_Report", to scan crewmember's names by cropping sections of PDFs dropped inside those directories.
# Saves in "output", PDFs with each crewmember's Timecard on top and Check Copies behind them according to the detected name.
# Saves in "output" also two XLSX files with errors and info about timecards and check copies that were unable to be matched with eachother / anything.

import numpy as np;
import pandas as pd;
import pytesseract;
import sys, os;
import cv2;
import math;
from PIL import Image; 
from pdf2image import convert_from_path, convert_from_bytes
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)

# use homebrew to install tesseract
# brew install tesseract
# brew info tesseract
# pip3 install numpy
# pip3 install pandas
# pip3 install Pillow
# pip3 install opencv-python
# pip3 install pdf2image
# pip3 install openpyxl

#files_tc = os.listdir(path_to_tc)

class PDF_FILE:
    def __init__(self,dir,name):
        self.path = dir+r'/'+name
        self.pages = []
        self.name = name
    def scan(self):
        self.pages = []
        for ind, pil_container in enumerate(convert_from_path(self.path)):
            pil_image = np.asarray(pil_container)
            pil_image_crop = pil_image[self.crop_coords["TOP"]:self.crop_coords["BOTTOM"],self.crop_coords["LEFT"]:self.crop_coords["RIGHT"]]
            pil_image_crop = cv2.resize(pil_image_crop, (0,0), fx=2, fy=2)
            hsv = cv2.cvtColor(pil_image_crop, cv2.COLOR_BGR2HSV)
            msk = cv2.inRange(hsv, np.array([0,0,120]), np.array([179,255,255]))
            thresh = cv2.threshold(msk, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
            inv = 255 - thresh
            inv = msk
            # pytesseract config
            languages_ = "eng"
            special_config = "--psm 6 --oem 3"
            # pytesseract search for text
            txt = pytesseract.image_to_string(inv, lang=languages_, config=special_config)
            txt_strp = txt.replace("*","").replace("_","").replace("(","").replace(")","").replace("-","").replace("0","").replace("1","").replace("2","").replace("3","").replace("4","").replace("5","").replace("5","").replace("6","").replace("7","").replace("8","").replace("9","")
            txt_strp_list_a = ""
            txt_strp_list_b = ""
            if "\n" in txt_strp[:-1]:
                txt_strp_list_b = txt_strp.split("\n")[1].split(" ")[1].split("/")[0]
                txt_strp_list_a = txt_strp.split("\n")[1].split(" ")[1].split("/")[1]
                txt_strp = " ".join([txt_strp_list_a,txt_strp_list_b])
            self.addPage(pil_container,txt_strp.rstrip(), ind)
            # preview the image
            #cv2.imshow("msk",inv)
            #cv2.waitKey(0)
    def addPage(self, text, index):
        # this gets overloaded by class TC_FILE and class CC_FILE definitions anyway
        pass
    def print(self):
        print("PRINTING: "+self.path)
        for el in self.pages:
            el.print()

class TC_FILE(PDF_FILE):
    def __init__(self, path, name):
        super().__init__(path, name)
        self.crop_coords = { "TOP": 190, "BOTTOM": 250, "LEFT": 60, "RIGHT": 700 }
    def addPage(self, pil_image, text, index):
        first = ""
        last = ""
        if len(text.split(",")) > 1:
            last = text.split(",")[0]
            first = text.split(",")[1]
            first = first.split(" ")[0]
        self.pages.append(PDF_FILE_PAGE(pil_image, first,last,index))

class CC_FILE(PDF_FILE):
    def __init__(self, path, name):
        super().__init__(path, name)
        self.crop_coords = { "TOP": 50, "BOTTOM": 100, "LEFT": 560, "RIGHT": 900 }

    def addPage(self, pil_image, text, index):
        first = ""
        last = ""
        if len(text.split(" ")) > 1:
            last = text.split(" ")[-1]
            first = text.split(" ")[0]
        self.pages.append(PDF_FILE_PAGE(pil_image, first,last,index))

class PDF_FILE_PAGE:
    def __init__(self, pil_img, first, last, num):
        self.first = first
        self.last = last
        self.index = num
        self.image = pil_img
        self.isMatched = False
    def print(self):
        print(str(self.index)+"\t"+self.first+" "+self.last)
    def get_image(self):
        return self.image
    def get_name(self):
        return str(self.first+" "+self.last)
    def get_filename(self):
        return str(self.last+", "+self.first).upper()
    def get_isMatched(self):
        return self.isMatched
    def set_isMatched(self, match):
        self.isMatched = match

class SEARCH_AREA:
    def __init__(self,pwd):
        self.path = pwd
        self.files = []
        pass
    def print(self):
        for el in self.files:
            el.print()
    def get_files(self):
        return self.files

class TIMECARDS(SEARCH_AREA):
    def __init__(self, pwd):
        super().__init__(pwd)
    def scan(self):
        self.files = []
        for filename in os.listdir(self.path):
            # validate input
            if filename[0] != ".":
                self.files.append(TC_FILE(self.path,filename))
        for file in self.files:
            file.scan()

class CHECKCOPIES(SEARCH_AREA):
    def __init__(self, pwd):
        super().__init__(pwd)
    def scan(self):
        self.files = []
        for filename in os.listdir(self.path):
            # validate input
            if filename[0] != ".":
                self.files.append(CC_FILE(self.path,filename))
        for file in self.files:
            file.scan()

def get_error_data(filesobj):
    data = { "filename": [], "page": [], "unmatched_name": [], "isMatched": [] }
    for file in filesobj.files:
        for page in file.pages:
            if page.get_isMatched() == False:
                data["filename"].append(file.name)
                data["page"].append(int(page.index)+1)
                data["unmatched_name"].append(page.get_name())
                data["isMatched"].append(page.get_isMatched())
    return data

def main_file_program(TIMECARDS, CHECKCOPIES):
    for tcfile in TIMECARDS.files:
        for tc in tcfile.pages:
            matching_images = []
            for ccfile in CHECKCOPIES.files:
                for cc in ccfile.pages:
                    if tc.get_name() == cc.get_name():
                        matching_images.append(cc.get_image())
                        cc.set_isMatched(True)
                        tc.set_isMatched(True)
            # append matches
            tc.get_image().save(root_folder+"/output/"+tc.get_filename()+".pdf", save_all=True, append_images=matching_images)

def main_file_report(TIMECARDS, CHECKCOPIES):
    pd.DataFrame(get_error_data(CHECKCOPIES)).to_excel(root_folder+"/output/# CHECKCOPIES - ERRORS.xlsx")
    pd.DataFrame(get_error_data(TIMECARDS)).to_excel(root_folder+"/output/# TIMECARDS - ERRORS.xlsx")




#########                         MAIN                        #########

root_folder = os.path.dirname(os.path.abspath(__file__))

main_timecards = TIMECARDS(root_folder+"/input/TC_Report")
main_checkcopies = CHECKCOPIES(root_folder+"/input/Check_Copy")

main_timecards.scan()
main_checkcopies.scan()

main_file_program(main_timecards,main_checkcopies)
main_file_report(main_timecards,main_checkcopies)

