import outlook
from getpass import getpass
import xlsxwriter

workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

mail = outlook.Outlook()
email_id = raw_input('Email : ') 
password = getpass() 
mail.login(email_id, password)
mail.readOnly('Inbox')
allIds = mail.allIds()
word = 'samvaad'
ctr = 0
for identifier in allIds:
    print identifier
    mail.getEmail(identifier)
    body = mail.mailbody()
    if(word in body):
        worksheet.write(ctr, 0, body)
workbook.close()
