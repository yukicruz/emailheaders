import win32com.client
import csv
import re
import mailbox
# from mailbox import Mailbox

dict_header = {
    'body': ''
    }

dict_header_parsed = {
    'Message-ID': ''
}

header_dict = {}
regular_view_dict = {}

# Navigate to folder with phishing email
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
phish_folder = inbox.Folders['PhishTest']
messages = phish_folder.Items
message = messages.GetLast()
mess = message.Body

# Important parts of email pulled from internet header
internet_header = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
message_id = re.search(r'Message-ID:\s<(.*?)>', internet_header).group(1) # <.*?> Nongreedy: matches "<python>" in "<python>perl>" # https://stackoverflow.com/questions/4666973/how-to-extract-the-substring-between-two-markers
subject_id = re.search(r'Subject:\s(.*)', internet_header).group(1) # <.*> Greedy repetition: matches "<python>perl>" Link: https://www.tutorialspoint.com/python/python_reg_expressions.htm

# Add header attributes to header_dict
header_dict['Message ID'] = message_id
header_dict['Subject ID'] = subject_id

# Print sections pulled from internet header
# print(header_dict)
# print(f"Internet Header: {internet_header}")
# print(f"Message ID: {message_id}")
# print(f"Subject: {subject_id}")


# Sections pulled from regular email
sender_name = message.SenderName
sender_email_address = message.SenderEmailAddress
date_sent = message.SentOn
recipient_to = message.To
recipient_cc = message.CC
recipient_bcc = message.BCC
email_subject = message.Subject
email_body = message.Body

# Add attributes of regular view of email to regular_view_dict
regular_view_dict['Sender Name'] = sender_name
regular_view_dict['Sender Email Address'] = sender_email_address
regular_view_dict['Date sent'] = date_sent
regular_view_dict['Recipients (To)'] = recipient_to
regular_view_dict['Recipients (CC)'] = recipient_cc
regular_view_dict['Recipients (BCC)'] = recipient_bcc
regular_view_dict['Subject'] = email_subject
regular_view_dict['Body'] = email_body

# Print sections pulled from regular email
# print(regular_view_dict)
# print(f"Sender Name: {sender_name}")
# print(f"Sender Email Address: {sender_email_address}")
# print(f"Date sent: {date_sent}")
# print(f"To: {recipient_to}")
# print(f"CC: {recipient_cc}")
# print(f"BCC: {recipient_bcc}")
# print(f"Subject: {email_subject}")
# print(f"Body: {email_body}")


########  LATER  ########
# Print parts of email ## use Class, db or CSV, and RE

## module: mailbox
# outlookm = mailbox.mboxMessage(outlook)

## wrapper: mail-parser
# import mailparser

# mail = mailparser.parse_from_bytes(byte_mail)
# mail = mailparser.parse_from_file(f)
# mail = mailparser.parse_from_file_msg(outlook_mail)
# mail = mailparser.parse_from_file_obj(fp)
# mail = mailparser.parse_from_string(raw_mail)

# print(mess)
# with open('headers.csv', 'w', newline='') as csvfile:
#     header_writer = csv.writer(csvfile, delimiter='\t', quotechar='|')
#     # header_writer.writerow(internet_header)
#     header_writer.writerow(["key", "value"]) # https://stackoverflow.com/questions/34283178/typeerror-a-bytes-like-object-is-required-not-str-in-python-and-csv
#     header_writer.writerows(dict_header.items());
    #
    # for key, value in dict_header.items(): # For Loops https://wiki.python.org/moin/ForLoop
    #     if key == 'Message-ID' :
    #         my_regex =
    #         dict_header.update({key:internet_header}); # Updating dictionaries # https://stackoverflow.com/questions/1024847/add-new-keys-to-a-dictionary

# <<<<<<< HEAD
#
# with open('headers.csv', 'w', newline='') as csvfile:
#     header_writer = csv.writer(csvfile, delimiter='\t', quotechar='|')
#     # internet_header = internet_header.replace('\n', '\r')
#     header_writer.writerow(internet_header)
# =======

def parse_results():
    f = open('DMV_driverNameIDCard_bodyRaw.csv','r') # https://stackoverflow.com/questions/9222106/how-to-extract-information-between-two-unique-words-in-a-large-text-file
    data = f.read()

    for key, value in dict_header_parsed.items():
        if key == 'Message-ID' :
            my_regex = r'' + re.escape(key) + r'(.*?)\*'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data, re.DOTALL)[0]}) # Adding [0] calls the first item from the list so it's in string; This line would call the entire list without the [0]
        elif key == 'LICENSE STATUS:' :
            my_regex = r'' + re.escape(key) + r'(.*?)DEPARTMENTAL ACTIONS:'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data, re.DOTALL)[0]})
        elif key == 'DEPARTMENTAL ACTIONS:' :
            my_regex = r'' + re.escape(key) + r'(.*?)CONVICTIONS:'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data, re.DOTALL)[0]})
        elif key == 'CONVICTIONS:' :
            my_regex = r'' + re.escape(key) + r'(.*?)FAILURES TO APPEAR:'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data, re.DOTALL)[0]})
        elif key == 'FAILURES TO APPEAR:' :
            my_regex = r'' + re.escape(key) + r'(.*?)ACCIDENTS:'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data, re.DOTALL)[0]})
        elif key == 'ACCIDENTS:' :
            my_regex = r'' + re.escape(key) + r'(.*?)END'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data, re.DOTALL)[0]})
        else:
            my_regex = r'' + re.escape(key) + r'(.*?)\*'
            dict_DMV_driverNameIDCard_parsed.update({key:re.findall(my_regex, data)[0]})
            # TO DO: Separate out lastname, firstname, middlename


def results_header():
    outfile = open('./DMV_driverNameIDCard_bodyRaw.csv', 'w')
    writer = csv.writer(outfile)
    writer.writerow(["key", "value"]) # https://stackoverflow.com/questions/34283178/typeerror-a-bytes-like-object-is-required-not-str-in-python-and-csv
    writer.writerows(dict_header.items());
