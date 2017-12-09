import readpdf
import readmsg

import glob
import time
import ntpath
import collections
import pandas as pd

stop_list = [ "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are", "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but", "by", "could", "did", "do", "does", "doing", "down", "during", "each", "few", "for", "from", "further", "had", "has", "have", "having", "he", "he'd", "he'll", "he's", "her", "here", "here's", "hers", "herself", "him", "himself", "his", "how", "how's", "i", "i'd", "i'll", "i'm", "i've", "if", "in", "into", "is", "it", "it's", "its", "itself", "let's", "me", "more", "most", "my", "myself", "nor", "of", "on", "once", "only", "or", "other", "ought", "our", "ours", "ourselves", "out", "over", "own", "same", "she", "she'd", "she'll", "she's", "should", "so", "some", "such", "than", "that", "that's", "the", "their", "theirs", "them", "themselves", "then", "there", "there's", "these", "they", "they'd", "they'll", "they're", "they've", "this", "those", "through", "to", "too", "under", "until", "up", "very", "was", "we", "we'd", "we'll", "we're", "we've", "were", "what", "what's", "when", "when's", "where", "where's", "which", "while", "who", "who's", "whom", "why", "why's", "with", "would", "you", "you'd", "you'll", "you're", "you've", "your", "yours", "yourself", "yourselves" ]

keywords = ["java", "python", "spark", "hadoop", "mapreduce", "reduce", "clearance"]
undesirable = ["india", "highschool", "high school"]


def get_text_from_files(filelist):

    file_text_dict = {}
    
    while len(filelist) > 0:
        file = filelist.pop(0).lower()
        
        text = None
        ###############
        ## if msg, get attachments
        ##  add attachment to filelist
        if file.endswith(".msg"):
            attachments = readmsg.get_msg_attachment(file)
            filelist.extend(filelist)
        
        ###############
        ## if PDF, read PDF text
        elif file.endswith(".pdf"):
            text = readpdf.get_pdf_text(file)
            
        ###############
        ## if word doc, read word text
        elif file.endswith(".docx"):
            print ("Word doc")
            
        else: 
            f, f_ext = os.path.splitext(file)
            print ("Skipping file type ", f_ext)
    
    
        if type(text) equals str:
            fn_key = ntpath.basename(pdf)
            fn_key = fn_key.replace("New Candidate ", "").split (" for ")[0]
            
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
    
def create_dict_for_resume(resume_text, filename):
    
    global keywords
    global undesirable
    
    
    output_dict = {}
    output_dict["resume pdf"] = filename
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


def create_excel_output(resume_dict_list, file_out):
    resume_df = pd.DataFrame(resume_dict_list)

    cols = ["resume pdf", "email", "phone", "key words", "red flags", "frequently used words"]
    resume_df = resume_df[cols]


    ##########
    ##Filter and sort
    ##########
    s = resume_df["key words"].str.split(",").apply(len).sort_values(ascending = False).index

    reindexed = resume_df.reindex(s)

    filtered_sorted = reindexed[reindexed["red flags"].apply(len) == 0]
    filtered_sorted = filtered_sorted[filtered_sorted["key words"].str.split(",").apply(len) > 1]
    
    timestr = time.strftime("%Y%m%d-%H%M%S")

    filename = "Developer_Resumes_" + timestr + ".xlsx"
    writer = pd.ExcelWriter(filename)

    resume_df.to_excel(writer, 'All Resumes', index=False)
    filtered_sorted.to_excel(writer, 'Filtered Resumes', index=False)

    writer.save()



def main(argv):
    ########################
    ## Parsing arguments
    
    ########################
    ## Get files to process
    filelist = glob.glob(os.path.join(path_in, "*.*"))
    
    #######################
    ## Process files
    file_text_map = get_text_from_files(filelist)
    
    ########################
    ## Write results


if __name__ == '__main__':
    main(sys.argv[1:])