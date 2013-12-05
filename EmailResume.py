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
    read_book = open_workbook('./Personal_Data/application_info.xls')
    r_sheet = read_book.sheet_by_index(0)
    info_list = []
    for row_index in range(1, r_sheet.nrows):
        info_list.append( dict(
            company_name = r_sheet.cell(row_index, 0).value,
            job_title    = r_sheet.cell(row_index, 1).value,
            contact_name = r_sheet.cell(row_index, 2).value,
            recip_email  = r_sheet.cell(row_index, 3).value,
            att_trans    = r_sheet.cell(row_index, 4).value,
            att_gre      = r_sheet.cell(row_index, 5).value,
            ))
    return info_list

# should accept 1. list from extract_application
#               2. email contend that rendered to CV

def render_CL(info):
    pass

def gen_log():
    pass


def sendEmail( recip_email, subject, msg):
    # Every email address should render the CV template and load info

    msg['To'] = recip_email
    msg['Subject'] = subject
    msg['From'] = username

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
    #gen_temp()
    info_list = extract_application()
    for info in info_list:
        # Render data to template cover letter
        # Render data to subject

        # Read Cover Letter
        fp = open("./Personal_Data/CV.html")
        msg = MIMEText(fp.read(), 'html')
        fp.close()
        subject = "This is a first test"
        sendEmail( info['recip_email'], subject, msg)
        # log the sending info
    #sendEmail()


if __name__ == '__main__':
    main()
