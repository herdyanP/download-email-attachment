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

import logging

logdate = datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
logging.basicConfig(filename=f'{logdate}.log', filemode='w', format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

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
            # Connect ke database
            printWithStamp('Connecting to database', end = '...', flush=True)
            try:
                mydb = mysql.connector.connect(
                    host=config['dbHost'],
                    user=config['dbUser'],
                    password=config['dbPass'],
                    database=config['dbName']
                )
                print('Done')
            except Exception as e:
                print('Fail')
                logging.exception('Exception occured when connecting to database!')
                quit()

            # Fetch CSV dari email
            dtformat = now.strftime('%Y%m%d')
            printWithStamp(f'Fetching "Data Upload CSV {dtformat}"', end='... ')
            try:
                mail = connectMail()
                mail.select()
                # typ, data = mail.search(None, f'(SUBJECT "Data Upload CSV {dtformat}")')
                typ, data = mail.search(None, f'(SUBJECT "{dtformat}")')
                for emailid in data[0].split():
                    downloadAttachments(mail, emailid)
                
                print('Done')
                uploadCSV(dtformat)
            except imaplib.IMAP4.error:
                logging.exception('Exception on mail')
                print('Fail')
            except AttributeError:
                logging.exception('Exception on mail')
                print('Fail')

            time.sleep(10) #interval pengecekan

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
        now = datetime.datetime.now()
        dtnow = now.strftime('%Y-%m-%d')
        printWithStamp(f'Processing "TABUNGAN_{dtformat}.csv"', end = '... ')

        sql = f'TRUNCATE TABLE saldotabungan_tmp'
        mycur.execute(sql)

        with open(rf'{outputDir}\csv\TABUNGAN_{dtformat}.csv') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            for row in csv_reader:
                namaclear = row[2].replace("'", "")
                alamclear = row[3].replace("'", "")

                # sql = f'insert into saldotabungan (CIF, SSREK, SSNAMA, SSALAMAT, SSSALDO, SSTGL, JPINJAMAN, STATUS) values ("{row[0]}", "{row[1]}", "{namaclear}", "{alamclear}", "{row[4]}", "{row[5]}", "{row[8]}", "{row[9]}");'
                sql = f'insert into saldotabungan_tmp (CIF, SSREK, SSNAMA, SSALAMAT, SSSALDO, SSTGL, JPINJAMAN, STATUS) values ("{row[0]}", "{row[1]}", "{namaclear}", "{alamclear}", "{row[4]}", "{row[5]}", "{row[8]}", "{row[9]}");'
                mycur.execute(sql)

                line_count += 1

            sql = f'insert into saldotabungan (CIF, SSREK, SSNAMA, SSALAMAT, SSSALDO, SSTGL, JPINJAMAN, STATUS) select CIF, SSREK, SSNAMA, SSALAMAT, SSSALDO, SSTGL, JPINJAMAN, STATUS from saldotabungan_tmp where SSREK not in (SELECT SSREK FROM saldotabungan WHERE SSTGL = "{dtnow}");'
            mycur.execute(sql)

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

try:
    config = yaml.safe_load(open('config.yml'))
except Exception as e:
    logging.exception('Exception occured')

    time.sleep(10)
    quit()

# ========== MAIN CODE HERE ===========
print("====== UPLOADER MCOLLECTION SLEMAN ======")
init()

# buat compile:
# pyinstaller --onefile --clean [nama_file]