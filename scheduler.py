from inspect import Attribute
from typing import final
import schedule
import time
import imaplib
import email
import os
import datetime
import csv
import mysql.connector
import yaml

try:
    config = yaml.safe_load(open('config.yml'))

    # mailSv = "imap.gmail.com"
    # mailUser = "herdyan.pradana@gmail.com"
    # mailPass = "drzsikqpbxcvrkqt"
    outputDir = os.getcwd()
    statSuccess = True

    def init():
        global statSuccess
        statSuccess = False

        start()

    def start():
        global statSuccess
        now = datetime.datetime.now()
        wk = now.isoweekday()

        if(wk <= 5): #biar cuma jalan di weekday
            while(statSuccess == False): #looping selama status masih gagal
                now = datetime.datetime.now()
                dtformat = now.strftime('%Y%m%d')

                printWithStamp(f'Fetching "Data Upload CSV {dtformat}"', end='... ')
                mail = connectMail()
                try:
                    mail.select()
                    # typ, data = mail.search(None, f'(SUBJECT "Data Upload CSV {dtformat}")')
                    typ, data = mail.search(None, f'(SUBJECT "{dtformat}")')
                    for emailid in data[0].split():
                        downloadAttachments(mail, emailid)
                    
                    print('Done')
                    uploadCSV(dtformat)
                except imaplib.IMAP4.error:
                    print('Fail')
                except AttributeError:
                    print('Fail')

                time.sleep(900) #interval pengecekan

    def connectMail():
        try:
            # mail = imaplib.IMAP4_SSL(mailSv)
            mail = imaplib.IMAP4_SSL(config['mailSv'])
            # mail.login(mailUser, mailPass)
            mail.login(config['mailUser'], config['mailPass'])
            mail.select()
            return mail
        except:
            return

    def downloadAttachments(mail, emailid):
        resp, data = mail.fetch(emailid, "(BODY.PEEK[])")
        email_body = data[0][1]
        msg = email.message_from_bytes(email_body)
        if msg.get_content_maintype() != 'multipart':
            return
        for part in msg.walk():
            if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
                # open(outputDir + '/csv/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))
                open(rf'{outputDir}\csv\{part.get_filename()}', 'wb').write(part.get_payload(decode=True))

    def uploadCSV(dtformat):
        global statSuccess
        try:
            printWithStamp('Connecting to database', end = '... ')
            # ====== real ======
            # mydb = mysql.connector.connect(
            #     host="192.168.3.81",
            #     user="mcollection",
            #     password="mediapos",
            #     database="mcollectionsleman"
            # )
            mydb = mysql.connector.connect(
                host=config['dbHost'],
                user=config['dbUser'],
                password=config['dbPass'],
                database=config['dbName']
            )

            # ====== localhost ======
            # mydb = mysql.connector.connect(
            #     host="localhost",
            #     user="root",
            #     password="",
            #     database="mcollectionsleman"
            # )
            print('Done')
        except:
            print('Fail')

        mycur = mydb.cursor()
        try:
            line_count = 0
            printWithStamp(f'Processing "TABUNGAN_{dtformat}.csv"', end = '... ')
            with open(rf'{outputDir}\csv\TABUNGAN_{dtformat}.csv') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=',')
                for row in csv_reader:
                    namaclear = row[2].replace("'", "")
                    alamclear = row[3].replace("'", "")

                    sql = f'insert into saldotabungan (CIF, SSREK, SSNAMA, SSALAMAT, SSSALDO, SSTGL, JPINJAMAN, STATUS) values ("{row[0]}", "{row[1]}", "{namaclear}", "{alamclear}", "{row[4]}", "{row[5]}", "{row[8]}", "{row[9]}");'
                    mycur.execute(sql)

                    line_count += 1

                print('Done')
                printWithStamp(f'Processed {line_count} lines in "TABUNGAN_{dtformat}.csv"')
                mydb.commit()
            
            if(line_count > 0):
                statSuccess = True
            
        except:
            print('Fail')
            printWithStamp(f'File TABUNGAN_{dtformat}.csv not found')
        finally:
            mycur.close()
            mydb.close()

    def printWithStamp(*args, **kwargs):
        dt = datetime.datetime.now()
        ts = dt.strftime('%Y-%m-%d %H:%M:%S')
        print(f'({ts}) ' + ' '.join(map(str,args)), **kwargs)

    def startAgain():
        if(statSuccess == False):
            printWithStamp('Restarting the process since the last attempt failed')
            start()



    # ========== MAIN CODE HERE ===========
    print("====== AUTO-UPLOADER MCOLLECTION SLEMAN ======")
    schedule.every().day.at("04:30").do(init)
    while True:
        schedule.run_pending()
        time.sleep(1)

    # buat compile:
    # pyinstaller --onefile --clean [nama_file]

except:
    print("Cannot read config file. Closing...")
    time.sleep(10)