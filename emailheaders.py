import win32com.client
import csv

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetLast()
mess = message.Body
internet_header = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
print(internet_header)

with open('headers.csv', 'w', newline='') as csvfile:
    header_writer = csv.writer(csvfile, delimiter='\t', quotechar='|')
    header_writer.writerow(internet_header)
