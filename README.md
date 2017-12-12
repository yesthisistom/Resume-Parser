# Resume Parser

Takes a folder of resumes, whether as standalone PDFs and Word Documents, or as attachments in .msg files (the way that Ziprecruiter sends them), and creates an excel spreadsheet of results. 

The spreadsheet contains the applicants contact information, as well as developer input desired key words and red flags, as well as the most frequently used words in the resume. 

The PDF reading software will work on both text PDFs and PDFs with the text as images if the machine running it has Google's [Tesseract OCR ](https://github.com/tesseract-ocr/tesseract) installed. 

## Required Software

See 'requirements.txt' for required packages.  This script and libraries were written using python 3.6.  

If you intend to process PDFs with images and no text, please install Google's [Tesseract OCR ](https://github.com/tesseract-ocr/tesseract).

'requirements.txt' created with pipreqs.

Please note, if using python 3, install the docx library using 

```
pip install python-docx
```

rather than the python 2 compatible version

```
pip install docx
```

## Included libraries

### readmsg.py

The readmsg.py file contains one function

get_msg_attachment(msg_in)
```
##########
## Takes a message file (.msg) in, and pull out all attachments.  
##  Attachements will have the same name as the message file, with the count and correct extension
##  Returns a list of attachments found
##########
```

Given a .msg file "test.msg" containing two PDF attachments, get_msg_attachment will extract those attachments to the same folder as the .msg file, and rename them "test_0.pdf" and "test_1.pdf". 

It will extract attachments of any extension. 

### readpdf.py

The main entry point into readpdf is the function 'get_pdf_text', which takes a path to a PDF as input. 

It will utilize the helper functions in the readpdf to extract text from the input PDF.  If there is no text, it will attempt to extract any images from the PDF and read the image text.  

If Tesseract is installed, edit the 'tessdata_dir_config' variable and 'pytesseract.pytesseract.tesseract_cmd' variable to reflect your install. These are in the 'get_text_from_image' function.

```
def get_text_from_image(image_in):
    tessdata_dir_config = '--tessdata-dir "C:\\Program Files (x86)\\Tesseract-OCR\\tessdata"'
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
```

### readdocx.py

The 'readdocx' library contains a single function, getDocxText, which takes the path to a single docx file as input.  

It will return a string of all the text contained in the docx. 

## How to Run 'resume_parser.py'

Before running, please review the global variables at the top of the file. 

```
keywords = ["java", "python", "spark", "hadoop", "mapreduce", "reduce"]
undesirable = ["highschool", "high school"]
```

In its current state, desirable keywords are development specific, but this list can be modified to fit your needs. The undesirable key words list should be modified to reflect the requirements of the job you are hiring for. 

### Running as Standalone

```
> python resume_parser.py -h

usage: resume_parser.py [-h] [-x EXISTINGEXCEL] -i INPUTDIR

Create Resume Triage Spreadsheet

optional arguments:
  -h, --help            show this help message and exit
  -x EXISTINGEXCEL, --existingExcel EXISTINGEXCEL
                        Previously created excel to update

required named arguments:
  -i INPUTDIR, --inputDir INPUTDIR
                        Directory containing resumes (pdf and .docx) or .msg
                        files

```


### Using the API

Running as standalone, the main() function in resume_parser.py uses argparse to get inputs from the user, and passes them to the resume_parser function. 

This function takes two arguments: The first is a list of files (.msg, .pdf, or .docx) to be parsed, and the second is an existing output from this script.  If no output already exists, the default is None. 


