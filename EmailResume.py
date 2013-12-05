#!/usr/bin/env python
import smtplib
import os
from email.mime.text import MIMEText
from xlutils.copy import copy
from xlwt import Workbook, easyxf
from xlrd import open_workbook, cellname
from time import gmtime, strftime
from datetime import datetime, timedelta

username = None
password = None

def gen_temp():
    book = Workbook(encoding="utf-8")
    sheet = book.add_sheet('Application Info')
    sheet.write(0, 0, "Company Name")
    sheet.write(0, 1, "Job Title")
    sheet.write(0, 2, "Contact Name")
    sheet.write(0, 3, "Recipient Email")
    sheet.write(0, 4, "Transcript")
    sheet.write(0, 5, "GRE")
    book.save("./Personal_Data/application_info.xls")

def extract_application():
    # Should return a list of dicts
    # 1. company_name
    # 2. job_title
    # 3. contact_name
    # 4. contact_address
    # 5. recip_email
    # 6. attach transcript
    # 7. attach GRE
    pass

# should accept 1. list from extract_application
#               2. email contend that rendered to CV

def gen_log():
    pass


def sendEmail():
    # Every email address should render the CV template and load info

    # Read Cover Letter
    fp = open("./Personal_Data/CV.html")
    msg = MIMEText(fp.read(), 'html')
    fp.close()

    msg['Subject'] = 'The contents of test'
    msg['From'] = username
    msg['To'] = 'qianyuzh@buffalo.edu'

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.ehlo()
    server.login(username,password)
    server.sendmail(msg['From'], [msg['To']], msg.as_string())
    server.quit()

def main():
    # Read Account Info
    f = open("./Personal_Data/gmail_account.txt")
    global username
    username = f.readline().split("=")[1].strip()
    global password
    password = f.readline().split("=")[1].strip()
    f.close()
    sendEmail()
    #gen_temp()

if __name__ == '__main__':
    main()
