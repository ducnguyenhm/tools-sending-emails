import openpyxl
import smtplib, ssl

mailbook = openpyxl.load_workbook(r"C:\Users\mrbin\OneDrive\Documents\Ams's Project\google_drive_download\mail.xlsx")
mailsheet = mailbook.active
contentbook = openpyxl.load_workbook(r"C:\Users\mrbin\OneDrive\Documents\Ams's Project\google_drive_download\content.xlsx")
contentsheet = contentbook.active
rowth = 0
colth = 0
mails = [] 
contents = []
for row in mailsheet.values:
    rowth += 1 
    if rowth == 1 : continue
    colth = 0
    for col in row: 
        colth += 1
        if colth == 1: continue
        mails.append(col)

rowth = 0
colth = 0
for row in contentsheet.values:
    rowth += 1 
    if rowth == 1 : continue
    colth = 0
    for col in row: 
        colth += 1
        if colth == 1: continue
        contents.append(col)

for i in mails:
    for j in contents: 
        print(i, end = " "), print(j)


# sending mail
port = 465  
smtp_server = "smtp.gmail.com"
sender_email = input("Type your email and press enter: ")  
password = input("Type your password and press enter: ")


context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    server.login(sender_email, password)
    
    for i in mails:
        for j in contents:
            server.sendmail(sender_email, i, j)


    
