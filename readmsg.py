import os
import ntpath
import win32com.client

##########
## Takes a message file (.msg) in, and pull out all attachments.  
##  Attachements will have the same name as the message file, with the count and correct extension
##  Returns a list of attachments found
##########
def get_msg_attachment(msg_in):
    msg_in = os.path.abspath(msg_in)
    folder = os.path.dirname(msg_in)
    
    filenames = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem(msg_in)        #filename including path
        att = msg.Attachments

        count = 0;
        for attachment in att:
            filename, file_extension = os.path.splitext(str(attachment))
            file_out = ntpath.basename(msg_in).replace(".msg", "_" + str(count) + file_extension.lower())
            
            filename = os.path.join(folder, file_out)
            attachment.SaveAsFile(filename)
            filenames.append(filename)
            
            count += 1

        del outlook, msg
    except Exception as e:
        print("Failed to read file " + msg_in)
        print(e)
    
    return filenames
    