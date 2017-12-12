
import docx

###########################
## Given a path to a word document, extracts the text.  
##  No error checking is done to confirm the file exists
##  For python 3, to import docx you need to run pip install python-docx, NOT pip install docx
############################
def getDocxText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)