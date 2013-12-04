#!/usr/bin/env python
import smtplib

FROM = 'test@gmail.com'
TO  = ['test@hotmail.com']
SUBJECT = "Testing sending using gmail"
TEXT = 'Why,Oh why!'

message = """\From: %s\nTo: %s\nSubject: %s\n\n%s""" % (FROM, ", ".join(TO), SUBJECT, TEXT)

username = 'test@gmail.com'
password = 'test'
server = smtplib.SMTP('smtp.gmail.com:587')
server.starttls()
server.ehlo()
server.login(username,password)
server.sendmail(FROM, TO, message)
server.quit()
