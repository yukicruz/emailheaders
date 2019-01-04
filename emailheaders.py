import win32com.client
import csv
import re

dict_header = {
    'body': ''
    }

dict_header_parsed = {
    'Message-ID': ''
}

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetLast()
mess = message.Body

# pull email header
internet_header = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")

## use Class, db or CSV, and RE
# pull specific info from email header
message_id = re.search(r'Message-ID:\s<(.*?)>', internet_header).group(1) # <.*?> Nongreedy: matches "<python>" in "<python>perl>" # https://stackoverflow.com/questions/4666973/how-to-extract-the-substring-between-two-markers
print("Message ID: " + message_id)

# print("\n") #
subject_id = re.search(r'Subject:\s(.*)', internet_header).group(1) # <.*> Greedy repetition: matches "<python>perl>" Link: https://www.tutorialspoint.com/python/python_reg_expressions.htm
print("Subject: " + subject_id)

print(mess)
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
