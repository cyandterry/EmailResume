#!/usr/bin/env python
import smtplib
import os
from email.mime.text import MIMEText
from xlutils.copy import copy
from xlwt import Workbook, easyxf
from xlrd import open_workbook, cellname
from time import gmtime, strftime
import datetime

username  = None
password  = None
real_name = None

def gen_temp():
    app_book = Workbook(encoding='utf-8')
    sheet1 = app_book.add_sheet('Application Info')
    sheet1.write(0, 0, 'Company Name')
    sheet1.write(0, 1, 'Job Title')
    sheet1.write(0, 2, 'Contact Name')
    sheet1.write(0, 3, 'Contact Address')
    sheet1.write(0, 4, 'Recipient Email')
    sheet1.write(0, 5, 'Transcript')
    sheet1.write(0, 6, 'GRE')
    app_book.save('./Personal_Data/application_info.xls')

    log_book = Workbook(encoding='utf-8')
    sheet2 = log_book.add_sheet('Application Info')
    sheet2.write(0, 0, 'Time')
    sheet2.write(0, 1, 'Company Name')
    sheet2.write(0, 2, 'Job Title')
    sheet2.write(0, 3, 'Contact Name')
    sheet2.write(0, 4, 'Contact Address')
    sheet2.write(0, 5, 'Recipient Email')
    sheet2.write(0, 6, 'Transcript')
    sheet2.write(0, 7, 'GRE')
    log_book.save('./Personal_Data/log.xls')

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
            company_name    = r_sheet.cell(row_index, 0).value,
            job_title       = r_sheet.cell(row_index, 1).value,
            contact_name    = r_sheet.cell(row_index, 2).value,
            contact_address = r_sheet.cell(row_index, 3).value,
            recip_email     = r_sheet.cell(row_index, 4).value,
            att_trans       = r_sheet.cell(row_index, 5).value,
            att_gre         = r_sheet.cell(row_index, 6).value,
            ))
    return info_list

def read_gmail_account():
    # Read Account Info
    f = open('./Personal_Data/gmail_account.txt')
    global username
    username = f.readline().split('=')[1].strip()
    global password
    password = f.readline().split('=')[1].strip()
    global real_name
    real_name = f.readline().split('=')[1].strip()
    f.close()

def render_CL(info):
    fp = open('./Personal_Data/CL.html')
    str_data = fp.read()
    fp.close()
    date = datetime.date.today()
    info['date'] = '%s %d, %s' %(date.strftime('%b'), int(date.strftime('%d')), date.strftime('%Y'))
    info['time'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    for key in info:
        str_data = str_data.replace('{%%%s}' % key, info[key])
    msg = MIMEText(str_data, 'html')
    return msg

def gen_log(info):
    read_book = open_workbook('./Personal_Data/log.xls')
    r_sheet = read_book.sheet_by_index(0)
    write_book = copy(read_book)
    w_sheet = write_book.get_sheet(0)

    # Copy read_book to write_book, which copys the existing logs
    for row_index in range(r_sheet.nrows):
        for col_index in range(r_sheet.ncols):
            w_sheet.write(row_index, col_index, r_sheet.cell(row_index, col_index).value)

    w_sheet.write(r_sheet.nrows, 0, info['time'])
    w_sheet.write(r_sheet.nrows, 1, info['company_name'])
    w_sheet.write(r_sheet.nrows, 2, info['job_title'])
    w_sheet.write(r_sheet.nrows, 3, info['contact_name'])
    w_sheet.write(r_sheet.nrows, 4, info['contact_address'])
    w_sheet.write(r_sheet.nrows, 5, info['recip_email'])
    w_sheet.write(r_sheet.nrows, 6, info['att_trans'])
    w_sheet.write(r_sheet.nrows, 7, info['att_gre'])
    write_book.save('./Personal_Data/log.xls')

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
    #gen_temp()
    read_gmail_account()
    info_list = extract_application()
    for info in info_list:
        msg = render_CL(info)
        subject = 'Application for %s from %s' % (info['job_title'], real_name)
        #sendEmail( info['recip_email'], subject, msg)
        gen_log(info)

if __name__ == '__main__':
    main()
