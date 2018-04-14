
try:
    import readmsg
except:
    print("Failed to import 'readmsg.py'.  Extracting attachments from .msg files will not be available.")
import readpdf
import readdocx

import os
import re
import sys
import glob
import time
import string
import ntpath
import argparse
import collections
import pandas as pd

stop_list = [ "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are", "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but", "by", "could", "did", "do", "does", "doing", "down", "during", "each", "few", "for", "from", "further", "had", "has", "have", "having", "he", "he'd", "he'll", "he's", "her", "here", "here's", "hers", "herself", "him", "himself", "his", "how", "how's", "i", "i'd", "i'll", "i'm", "i've", "if", "in", "into", "is", "it", "it's", "its", "itself", "let's", "me", "more", "most", "my", "myself", "nor", "of", "on", "once", "only", "or", "other", "ought", "our", "ours", "ourselves", "out", "over", "own", "same", "she", "she'd", "she'll", "she's", "should", "so", "some", "such", "than", "that", "that's", "the", "their", "theirs", "them", "themselves", "then", "there", "there's", "these", "they", "they'd", "they'll", "they're", "they've", "this", "those", "through", "to", "too", "under", "until", "up", "very", "was", "we", "we'd", "we'll", "we're", "we've", "were", "what", "what's", "when", "when's", "where", "where's", "which", "while", "who", "who's", "whom", "why", "why's", "with", "would", "you", "you'd", "you'll", "you're", "you've", "your", "yours", "yourself", "yourselves" ]

keywords = ["java", "python", "spark", "hadoop", "mapreduce", "reduce"]
undesirable = ["highschool", "high school"]

output_excel_prefix = "Developer_Resumes_"

def get_text_from_files(filelist, existing_df):

    file_text_dict = {}
    
    idx = 0
    while idx < len(filelist):
        file = filelist[idx]
        idx += 1
        
        fn_key = ntpath.basename(file)
        fn_key = fn_key.replace("New Candidate ", "").split (" for ")[0]
        
        #any(df.column == 07311954)
        if (not existing_df is None) and any(existing_df.iloc[:, 0] == fn_key):
            print("Skipping", fn_key)
            continue
        
        text = None
        ###############
        ## if msg, get attachments
        ##  add attachment to filelist
        if file.lower().endswith(".msg"):
            attachments = readmsg.get_msg_attachment(file)
            if attachments and len(attachments) > 0:
                for att in attachments:
                    if not att in filelist:
                        filelist.append(att)
            
        
        ###############
        ## if PDF, read PDF text
        elif file.lower().endswith(".pdf"):
            text = readpdf.get_pdf_text(file)
            
        ###############
        ## if word doc, read word text
        elif file.lower().endswith(".docx"):
            text = readdocx.getDocxText(file)
            
        else: 
            f, f_ext = os.path.splitext(file)
            print ("Skipping file type ", f_ext)
    
    
        if type(text) == str:
            file_text_dict[fn_key] = text
            
    return file_text_dict
    
def get_bag_of_words_from_resume(resume_text):

    global stop_list
    
    #Lower case, remove punctuation
    translator = str.maketrans('', '', string.punctuation)
    resume_text = resume_text.lower()
    
    regex = re.compile('[%s]' % re.escape(string.punctuation))
    resume_text = regex.sub(' ', resume_text)
    
    resume_words = re.split("[ \n]", resume_text)
    
    
    bag_of_words = []
    for word in resume_words: 
        word = word.strip()
        if word not in stop_list and len(word) > 1:
            bag_of_words.append(word)
            
    return bag_of_words
    
def create_dict_for_resume(resume_text, resume_id):
    
    global keywords
    global undesirable
    
    
    output_dict = {}
    output_dict["resume id"] = resume_id
    # Get email
    email_pattern = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
    email_str = ""
    for email in re.findall(email_pattern, resume_text):
        if len(email_str) > 0:
            email_str += ","
        email_str += email
        
    output_dict["email"] = email_str    
    # Get phone
    phone_pattern = r"\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}"
    phone_str = ""
    for phone in re.findall(phone_pattern, resume_text):
        if len(phone_str) > 0:
            phone_str += ","
        phone_str += phone
    
    output_dict["phone"] = phone_str
    
    # Get key words
    keywords_found = ""
    for keyword in keywords:
        if keyword in resume_text.lower():
            if len(keywords_found) > 0:
                keywords_found += ","
            keywords_found += keyword
    
    output_dict["key words"] = keywords_found
    # Search for excluded words
    badwords_found = ""
    for badword in undesirable:
        if badword in resume_text.lower():
            if len(badwords_found) > 0:
                badwords_found += ","
            badwords_found += badword
    output_dict["red flags"] = badwords_found
    
    # get most common words
    bag_of_words = get_bag_of_words_from_resume(resume_text)
    word_counter = collections.Counter(bag_of_words)
    
    common_words = ""
    for common in word_counter.most_common(15):
        if len(common_words) > 0:
            common_words += ","
        common_words += common[0] + ":" + str(common[1])
        
    output_dict['frequently used words'] = common_words
        
    
    return output_dict


def create_excel_output(resume_dict_list, existing_df, folder):
    resume_df = pd.DataFrame(resume_dict_list)

    cols = ["resume id", "email", "phone", "key words", "red flags", "frequently used words"]
    resume_df = resume_df[cols]
    resume_df["reviewed"] = "no"
    resume_df["notes"] = ""
    resume_df["interview"] = ""

    if not existing_df is None:
        resume_df = pd.concat([resume_df, existing_df])

    ##########
    ##Filter and sort
    ##########
    s = resume_df["key words"].str.split(",").apply(len).sort_values(ascending = False).index

    reindexed = resume_df.reindex(s)

    filtered_sorted = reindexed[reindexed["red flags"].apply(len) == 0]
    filtered_sorted = filtered_sorted[filtered_sorted["key words"].str.split(",").apply(len) > 1]
    
    timestr = time.strftime("%Y%m%d-%H%M%S")

    filename = output_excel_prefix + timestr + ".xlsx"

    filename = os.path.join(folder, filename)
    writer = pd.ExcelWriter(filename)

    resume_df.to_excel(writer, 'All Resumes', index=False)
    filtered_sorted.to_excel(writer, 'Filtered Resumes', index=False)

    writer.save()
    
def resume_parser(file_list, input_dir, existing_excel=None):
    
    #######################
    ## Read existing file
    #######################
    existing_df = None
    if existing_excel:
        existing_df = pd.read_excel(existing_excel, sheet_name="All Resumes")
    
    #######################
    ## Process files
    print("Reading Resume Text for", len(file_list), "files")
    file_text_map = get_text_from_files(file_list, existing_df)
    
    resume_dict_list = []
    for resume_id, resume in file_text_map.items():
        resume_dict_list.append(create_dict_for_resume(resume, resume_id))
    
    ########################
    ## Write results
    print ("Writing results")
    create_excel_output(resume_dict_list, existing_df, input_dir)
    
    return


def main(argv):

    ########################
    ## Parsing arguments
    parser = argparse.ArgumentParser(description='Create Resume Triage Spreadsheet')
    parser
    parser.add_argument("-x", "--existingExcel", help="Previously created excel to update", default=None)
    
    requiredNamed = parser.add_argument_group('required named arguments')
    requiredNamed.add_argument("-i", "--inputDir", help="Directory containing resumes (pdf and .docx) or .msg files", required=True)
                    
    args = parser.parse_args()
    
    input_dir = args.inputDir
    existing_excel = args.existingExcel
    
    if input_dir is None:
        parser.print_help()
        return 
        
    if not os.path.exists(os.path.dirname(input_dir)):
        print ("Directory not found: " + input_dir)
        return
        
    if existing_excel:
        if s.path.splitext(file)[1].lower() != ".xlsx":
            print ("Input for existing excel is not an excel file")
            return
        if not os.path.exists(existing_excel):
            print ("Unable to find existing excel file")
            return
        
   
    ########################
    ## Get files to process
    print ("Finding Files")
    file_list = glob.glob(os.path.join(input_dir, "*.*"))
    supported_exts = [".docx", ".pdf", ".msg"]
    
    #########
    ## If no excel file was input, try to find one
    if not existing_excel:
        excel_files = []
        for file in file_list:
            if os.path.splitext(file)[1].lower() == ".xlsx" and os.path.basename(file).startswith(output_excel_prefix):
                excel_files.append(file)
        if len(excel_files) > 0:
            existing_excel = sorted(excel_files)[-1]
            print ("Using Excel " + os.path.basename(file))
    
    for file in file_list:
        if not os.path.splitext(file)[1].lower() in supported_exts:
            file_list.remove(file)
    
    if len(file_list) == 0:
        print("No supported files found in directory: " + input_dir)
        return
    
    #####################
    ## Call the API
    #####################
    resume_parser(file_list, input_dir, existing_excel)
    
    return 

if __name__ == '__main__':
    main(sys.argv[1:])
