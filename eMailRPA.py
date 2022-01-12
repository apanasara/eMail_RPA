import imaplib
import pprint
import email
import pandas as pd

imap_host = 'imap.secureserver.net' # server address
imap_user = '_____@______.com' #eMail ID
imap_pass = '____________' # password

search_folder = "INBOX"
search_string ='(FROM "________________")' # eMail search string, for more kindly refer documentation of imaplib library
from_to_line_number = [9,22] # starting line no & ending line no of eMail, in which interested for scrapping

output_folder_file = "D:\Amit\Python\emailExtract.csv" # path where output file to be stored
print("filepath:{}".format(output_folder_file.replace('/','\\\\')))
df_marks = pd.DataFrame()

#-------------- IMAP SSL -----------------------

with imaplib.IMAP4_SSL(host=imap_host, port=imaplib.IMAP4_SSL_PORT) as imap_ssl:
    print("Connection Object : {}".format(imap_ssl))

    #----------------- Login to Mailbox --------------
    print("Logging into mailbox...")
    resp_code, response = imap_ssl.login(imap_user, imap_pass)

    print("Response Code : {}".format(resp_code))
    print("Response      : {}\n".format(response[0].decode()))

    #-------------- Set Mailbox ----------------
    resp_code, mail_count = imap_ssl.select(mailbox=search_folder, readonly=True)

    #---------------- Search mails in a given Directory -------------------   
    resp_code, mails = imap_ssl.search(None, search_string)
    mail_ids = mails[0].decode().split()
    print("Total Mail IDs : {}\n".format(len(mail_ids)))

    for mail_id in mail_ids:
        print("{}\n".format(mail_id))
        #print("================== Start of Mail [{}] ====================".format(mail_id))
        resp_code, mail_data = imap_ssl.fetch(mail_id, '(RFC822)') ## Fetch mail data.
        message = email.message_from_bytes(mail_data[0][1]) ## Construct Message from mail data
        """
        print("From       : {}".format(message.get("From")))
        print("To         : {}".format(message.get("To")))
        print("Bcc        : {}".format(message.get("Bcc")))
        print("Date       : {}".format(message.get("Date")))
        print("Subject    : {}".format(message.get("Subject")))
        print("Body : ")
        """
        for part in message.walk():
            if part.get_content_type() == "text/html":
                body_lines = part.as_string().split("\n")
                # Convert key-value String to dictionary
                # Using map() + split() + loop
                res = []
                for sub in "".join(body_lines[from_to_line_number[0]:from_to_line_number[1]]).split(' <br>'):
                    if ':' in sub:
                        res.append(map(str.strip, sub.split(': ', 1)))
                res = dict(res)
                df_marks = df_marks.append(res, ignore_index=True)
                #for ln in line_to_extract:
                #   print("\n".join(body_lines[ln:ln+1])[:-5].split(sep=': ', maxsplit=1)[1]) 
        #print("================== End of Mail [{}] ====================\n".format(mail_id))


    #------------- Close Selected Mailbox -----------------------
    print("\nClosing selected mailbox....")
    imap_ssl.close()
    #print(df_marks)
    df_marks.to_csv(output_folder_file.replace('/','\\\\'))