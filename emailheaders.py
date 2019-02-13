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
# message = messages[0]  # messages[0] accesses the oldest email in the folder; messages[1] would access the 2nd oldest email in the folder
message = messages[0]  # Access the [nth] email in the folder; 0 is the oldest, n is the newest
# message = messages.GetLast()  # Access the newest email in the folder
mess = message.Body

# Important parts of email pulled from internet header
internet_header = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
message_id = re.search(r'Message-ID:\s<(.*?)>', internet_header).group(1) # <.*?> Nongreedy: matches "<python>" in "<python>perl>" # https://stackoverflow.com/questions/4666973/how-to-extract-the-substring-between-two-markers
subject_id = re.search(r'Subject:\s(.*)', internet_header).group(1) # <.*> Greedy repetition: matches "<python>perl>" Link: https://www.tutorialspoint.com/python/python_reg_expressions.htm
# origin_ip = re.findall("Received:\sfrom\s.*\((\d+\.\d+\.\d+\.\d+)\)", internet_header)[-1]
try:
    origin_ip = re.findall("Received:\sfrom\s.*\((\d+\.\d+\.\d+\.\d+)\)", internet_header)[-1]  # Create list of instances following "Received from:..." and pull the last instance.
except:
    origin_ip = "Did not find IP"

try:
    origin_smtp = re.findall("Authentication-Results-Original:.*smtp\.mailfrom=([^\s]*);", internet_header, re.DOTALL)[0]  # The re.DOTALL flag tells python to make the ‘.’ special character match all characters, including newline characters
except:
    origin_smtp = "Did not find SMTP"

try:
    email_addresses = re.findall("[\w\.-]+@[\w\.-]+", internet_header, re.DOTALL)
except:
    email_addresses = "Did not find email addresses."

ipAddresses = re.findall(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b", internet_header, re.DOTALL)

#Removing all empty sublists
# ipAddresses = list(filter(None, ipAddresses))



# Add header attributes to header_dict
header_dict['Message ID'] = message_id
header_dict['Subject ID'] = subject_id
header_dict['Originating IP'] = origin_ip
header_dict['Originating SMTP'] = origin_smtp
header_dict['Email Addresses'] = email_addresses
header_dict['IP Addresses'] = ipAddresses

# Print sections pulled from internet header
# print(header_dict)
# print(f"Internet Header: {internet_header}")
# print(f"Message ID: {message_id}")
print(f"Subject: {subject_id}")
print(f"Originating IP: {origin_ip}")
print(f"Originating SMTP: {origin_smtp}")
print(f"ipAddresses: {ipAddresses}")
print(f"Email Addresses: {email_addresses}")

# Sections pulled from regular email (aka User-specified details)
sender_name = message.SenderName
sender_email_address = message.SenderEmailAddress
date_sent = message.SentOn
recipient_to = message.To
recipient_cc = message.CC
recipient_bcc = message.BCC
email_subject = message.Subject
email_body = message.Body  # raw email body
# email_body = ' '.join(email_body.split(sep=None))
email_body = '"' + email_body + '"'

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

# Add regular_view_dict to csv
with open('test.csv', 'w') as f:
    for key in regular_view_dict.keys():
        f.write("%s,%s \n"%(key,regular_view_dict[key]))
    for key in header_dict.keys():
        f.write("%s,%s \n"%(key,header_dict[key]))

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
