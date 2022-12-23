import mysql.connector

mydb = mysql.connector.connect(
    host='localhost',
    user='root',
    password='',
    database='mcollectionsleman'
)

mycur = mydb.cursor()
sql = 'DROP TABLE saldotabungan'
mycur.execute(sql)

sql = 'CREATE TABLE '