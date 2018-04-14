import os
import cv2
try:
    import fitz
except:
    print("Failed to import fitz.  Reading images from PDF will not be available.")
import sys,  re

from PIL import Image

import pytesseract
from pytesseract import image_to_string

from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine


###################
## Requires Fitz
##  
## Get Images out of PDF.  Takes in path to PDF
##  Returns list of path to images, with the same name as the PDF, with '.pdf' converted to image count and ".png"
###################
def get_pdf_images(pdf_file):
    checkXO = r"/Type(?= */XObject)"       # finds "/Type/XObject"   
    checkIM = r"/Subtype(?= */Image)"      # finds "/Subtype/Image"

    doc = fitz.open(pdf_file)
    imgcount = 0
    lenXREF = doc._getXrefLength()         # number of objects - do not use entry 0!
    
    img_files = []

    for i in range(1, lenXREF):            # scan through all objects
        text = doc._getObjectString(i)     # string defining the object
        isXObject = re.search(checkXO, text)    # tests for XObject
        isImage   = re.search(checkIM, text)    # tests for Image
        if not isImage:
            isImage = "/Subtype/Image" in text
        if not isXObject and not isImage:   # not an image object if not both True
            continue
        imgcount += 1
        pix = fitz.Pixmap(doc, i)          # make pixmap from image
        
        filename = pdf_file.replace(".pdf", "") + "-%s.png" % (i,)
        if pix.n < 5:                      # can be saved as PNG
            pix.writePNG(filename)
        else:                              # must convert the CMYK first
            pix0 = fitz.Pixmap(fitz.csRGB, pix)
            pix.writePNG(filename)
            pix0 = None                    # free Pixmap resources
            
        reverse_image(filename)
        img_files.append(filename)
        pix = None                         # free Pixmap resources

    return img_files
    
    
###############
## Requires cv2
##  The Images read in from PDF are reversed in the y direction
##  Reads in the provided PNG file, reverses the axis, and writes to the same filename
###############    
def reverse_image(filename):
    img = cv2.imread(filename)
    
    #rimg=cv2.flip(img,1)
    fimg=cv2.flip(img,0)
    cv2.imwrite(filename, fimg)
    

#################
## Requires tesseract to be installed, as well as the pytesseract library
##   Alter the tesseract install location if required
##  
##  Input is an image
##  Output is a string of all text in image
#################
def get_text_from_image(image_in):
    tessdata_dir_config = r'--tessdata-dir "C:\Program Files (x86)\Tesseract-OCR\tessdata"'
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    
    if os.path.isfile(image_in):
        image_data = Image.open(image_in)
        return image_to_string(image_data, lang='eng',  config=tessdata_dir_config)
        
        
######################
## Requires PDFMiner3k
##
##  Takes a PDF file in, and returns a string of all the text.
######################
def get_pdf_text(pdf_file):
    pdf_text = ""
    
    with open(pdf_file, 'rb') as file_hdl:
                
        parser = PDFParser(file_hdl)
        doc = PDFDocument()
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize('')
        
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        laparams.char_margin = 1.0
        laparams.word_margin = 1.0
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        for page in doc.get_pages():
            interpreter.process_page(page)
            layout = device.get_result()
            for lt_obj in layout:
                if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                    pdf_text += lt_obj.get_text()
    
    if len(pdf_text) == 0:
        img_files = get_pdf_images(pdf_file)
        for img_file in img_files:
            pdf_text += get_text_from_image(img_file)
    
    return pdf_text
