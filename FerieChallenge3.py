import smtplib
import email
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header    import Header
from email.encoders import encode_base64
import ssl
import xlrd
import re
import os
from getpass import getpass

class Person:
    def __init__(self, email, name, surname):
        self.email = email
        self.name = name
        self.surname = surname


def mail(username,passwd,to, subject, text, attach, url):
    msg = MIMEMultipart()
    
    msg['From'] = username
    msg['To'] = to
    msg['Subject'] = subject
    
    msg.attach(MIMEText(text)) 
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(attach, 'rb').read())
    encode_base64(part)
    try:
        part.add_header('Content-Disposition',
                'attachment; filename="%s"' % os.path.basename(attach))
        msg.attach(part)
    except:
        print("Failed to load image, is it exist?")
    try:
        server = smtplib.SMTP(url, 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(username, passwd)
        server.sendmail(username, to, msg.as_string())
        server.close()
        time.sleep(5)
    except:
        print("failed to send mail")


address_file_path = input("Enter address file name: ")
user = input('Enter your e-mail address: ')
url = input('Enter your SMTP server url: ')
password = getpass('Enter your e-mail password: ')

col_names =[] 
person_array =[]
col_indexes=[] #1st index is email 2nd is for names
email_array=[]
names_array=[]


work_book = xlrd.open_workbook(address_file_path)
sheet = work_book.sheet_by_index(0)
sheet.cell_value(0, 0)
for i in range(sheet.ncols):
    col_names.append(sheet.cell_value(0, i))

for i in range(len(col_names)):
    if col_names[i] == 'E-mail':
        col_indexes.append(i)
    elif col_names[i] == 'Imię i nazwisko':
        col_indexes.append(i)
if len(col_indexes) != 2:
    raise Exception("Wrong data file format!")

for i in range(1,sheet.nrows):
    email_array.append(sheet.cell_value(i, col_indexes[0]))
    names_array.append(sheet.cell_value(i, col_indexes[1]))

if len(names_array) != len(email_array):
    raise Exception("Number of names is not same as emails number!")

for i in range(len(names_array)):
    temp = names_array[i].split(' ')
    person_array.append(Person(email_array[i], temp[0], temp[1]))

for i in range(len(person_array)):
    mail(user, password,person_array[i].email, "Your image", "Hi "+person_array[i].name+" it’s file generated for you",person_array[i].name+"_"+person_array[i].surname+"_image.png",url)
    


