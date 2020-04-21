import imaplib
import email
import os

mail = imaplib.IMAP4_SSL('imap.gmail.com')
# imaplib module implements connection based on IMAPv4 protocol

mail.login('aashay.shah@somaiya.edu', 'moneyheist123')

#print(mail.list())
# Lists all labels in GMail

mail.select('inbox')
# Connected to inbox

result, data = mail.uid('search', None, "ALL")
#print(data)
# Search and returns UID

inbox_item_list = data[0].split()
#print(inbox_item_list)

most_recent = inbox_item_list[-11]
result, email_data = mail.uid('fetch', most_recent, '(RFC822)')
# Fetch the email body

raw_email = email_data[0][1].decode('utf-8')
email_message = email.message_from_string(raw_email)
# Converts byte literal to string

print(f"Recepients: {email_message['To']}\n")
print(f"Sender: {email_message['From']}\n")
print(f"Subject: {email_message['Subject']}\n")
print(f"Date: {email_message['Date']}\n")
print(f"CC: {email_message['CC']}\n")

# The next for loop is used to obtain the contents of the email
for part in email_message.walk():
    if part.get_content_type() == 'text/plain':
        body = part.get_payload(decode=True)
        save_string = str("" + "content" + ".eml")
        myfile = open(save_string, 'w')
        myfile.write(body.decode('utf-8'))
        myfile.close()
    else:
        continue

# The next for loop is used to download the attachments if any
for part in email_message.walk():
    if part.get_content_maintype() == 'multipart':
        continue
    if part.get('Content-Disposition') is None:
        continue
    fileName = part.get_filename()
    print(f"Attachment: {fileName}")
    if bool(fileName):
        filePath = os.path.join('', fileName)
        if not os.path.isfile(filePath):
            fp = open(filePath, 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
        subject = email_message['Subject']
        print('Downloaded "{file}" from email titled "{subject}"'.format(file=fileName, subject=subject,))
