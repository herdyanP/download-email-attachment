import imaplib
import email
import os
import builtins
import datetime

mailSv = "imap.gmail.com"
mailUser = "herdyan.pradana@gmail.com"
# mailPass = "drzsikqpbxcvrkqt"
mailPass = "drzsikqpbxcvrkqt?"
outputDir = os.getcwd()

# connects to email client through IMAP
def connect():
    mail = imaplib.IMAP4_SSL(mailSv)
    mail.login(mailUser, mailPass) # login ke server email
    mail.select() # select folder yang mau dipake (default: inbox)
    return mail

def downloadAttachments(mail, emailid):
    resp, data = mail.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    msg = email.message_from_bytes(email_body)
    if msg.get_content_maintype() != 'multipart':
        return
    for part in msg.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputDir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))

def start():
    try:
        mail = connect()
        mail.select()
        typ, data = mail.search(None, '(SUBJECT "DATA UPLOAD CSV 20220715")')
        print(data)
    except imaplib.IMAP4.error:
        print('errors occured')

    # for emailid in data[0].split():
        # downloadAttachments(mail, emailid)

# def print(*str):
#     now = datetime.datetime.now()
#     stamp = now.strftime('%Y-%m-%d %H:%M:%S')
#     builtins.print(f'{stamp} {str} ')

# start()

def coba(*args, **kwargs):
    print(' '.join(map(str,args)), **kwargs)

abc = 123
coba(f'coba aja {abc}', end = '...')
coba('yahaha')



# mail = imaplib.IMAP4_SSL(mailSv)
# mail.login(mailUser, mailPass)
# mail.select()
# typ, data = mail.search(None, '(SUBJECT "DATA UPLOAD CSV")')
# for num in data[0].split():
#     typ, data = mail.fetch(num, '(RFC822)')
#     print('Message %s\n%s\n' % (num, data[0][1]))

# mail.close()
# mail.logout()