import schedule
import time
import imaplib
import email
import os
import datetime
import csv
import mysql.connector

mailSv = "imap.gmail.com"
mailUser = "herdyan.pradana@gmail.com"
mailPass = "drzsikqpbxcvrkqt"
outputDir = os.getcwd()

def start():
    now = datetime.datetime.now()
    dtformat = now.strftime('%Y%m%d')

    printWithStamp(f'Fetching "Data Upload CSV {dtformat}"', end='... ')
    try:
        mail = connectMail()
        mail.select()
        typ, data = mail.search(None, f'(SUBJECT "Data Upload CSV {dtformat}")')
        for emailid in data[0].split():
            downloadAttachments(mail, emailid)
        
        print('Done')
        uploadCSV(dtformat)
    except imaplib.IMAP4.error:
        print('Fail')

def connectMail():
    mail = imaplib.IMAP4_SSL(mailSv)
    mail.login(mailUser, mailPass)
    mail.select()
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

def uploadCSV(dtformat):
    try:
        printWithStamp('Connecting to database', end = '... ')
        mydb = mysql.connector.connect(
            host="192.168.3.81",
            user="mcollection",
            password="mediapos",
            database="mcollectionsleman"
        )
        print('Done')
    except:
        print('Fail')

    mycur = mydb.cursor()
    try:
        printWithStamp(f'Processing "Data Upload CSV {dtformat}"', end = '... ')
        with open(f'TABUNGAN_{dtformat}.csv') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                namaclear = row[2].replace("'", "")
                alamclear = row[3].replace("'", "")

                sql = f'insert into saldotabungan (CIF, SSREK, SSNAMA, SSALAMAT, SSSALDO, SSTGL, JPINJAMAN, STATUS) values ("{row[0]}", "{row[1]}", "{namaclear}", "{alamclear}", "{row[4]}", "{row[5]}", "{row[8]}", "{row[9]}");'
                mycur.execute(sql)

                line_count += 1

            print('Done')
            printWithStamp(f'Processed {line_count} lines in "Data Upload CSV {dtformat}"')
            mydb.commit()
    except:
        print('Fail')
        printWithStamp(f'File TABUNGAN_{dtformat}.csv not found')

def printWithStamp(*args, **kwargs):
    dt = datetime.datetime.now()
    ts = dt.strftime('%Y-%m-%d %H:%M:%S')
    print(f'({ts}) ' + ' '.join(map(str,args)), **kwargs)



# ========== MAIN CODE HERE ===========
print("====== AUTO-UPLOADER MCOLLECTION SLEMAN ======")
schedule.every().day.at("04:30").do(start)
while True:
    schedule.run_pending()
    time.sleep(300)