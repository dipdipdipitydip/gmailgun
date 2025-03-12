
import pandas as pd
import smtplib
import datetime

import os

import time

from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PIL import Image                                                                                
from logo import logo
import colored
from colored import Fore, Back, Style
#from email_content import email_message


current_time = datetime.datetime.now()

sending_email = input("| ̵͇̿̿/’̿ ̿̿   💥 Enter sending email : ")

password = input('| ̵͇̿̿/’̿ ̿̿   🔒 Enter Gmail app password? : ')

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(sending_email, password)
list = True

send_list = input("| ̵͇̿̿/’̿ ̿̿   📃 Enter xlsx email list : " )
# check if file exits
if os.path.exists(send_list):
   
    print('''
---------------------------------------------📨
''')
else:
    print(f"| ̵͇̿̿/’̿ ̿̿   ❌ File {send_list} not found")
    send_list = input("| ̵͇̿̿/’̿ ̿̿   📃 Enter xlsx email list : " )
    print('''
---------------------------------------------📨
''')
    
def email_message():
    global email_msg

    email_msg = input("| ̵͇̿̿/’̿ ̿̿   📃 Enter email message file : " )
    # check if file exits
    if os.path.exists(email_msg):
   
        print('''
---------------------------------------------📨
''')
    else:
        print(f"| ̵͇̿̿/’̿ ̿̿   ❌ File {email_msg} not found")
        email_msg = input("| ̵͇̿̿/’̿ ̿̿   📃 Enter email message file : " )
        print('''
---------------------------------------------📨
''')
        

# reading the spreadsheet
email_list = pd.read_excel(f"{send_list}")

# getting the names and the emails
names = email_list['NAME']
emails = email_list['EMAIL']

subject = input("| ̵͇̿̿/’̿ ̿̿   📃 Enter email subject line : " )

email_message()

file_path = email_msg

with open(file_path, 'r') as file:
    file_content = file.read()

    
html = file_content

# for every record get the name and the email addresses
for i in range(len(emails)):
    name = names[i]
    email = emails[i]
 
# iterate through the records

    msg = MIMEMultipart()
    msg['From'] = sending_email
    msg['To'] = email
    msg['Subject'] = subject

    part2 = MIMEText(html, 'html')

    msg.attach(part2)


    # sending the email
    time.sleep(6)
    server.sendmail(sending_email, [email], msg.as_string())
    current_time = datetime.datetime.strptime(datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S"),"%Y-%m-%d %H:%M:%S")
   
  
    print(f"{Fore.rgb(0, 0, 255)}To: {Fore.rgb(255, 140, 0)}{email} | Subject: {subject} | {Fore.rgb(255, 140, 0)}Time Sent: | {current_time} | {Fore.rgb(0, 128, 0)} From: {sending_email}")
    time.sleep(6)
 
	# close the smtp server
server.close()

 
