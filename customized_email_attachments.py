# -*- coding: utf-8 -*-
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

import pandas
from pandas import ExcelWriter
import os
import glob

dfInvestors = pandas.read_csv("File.csv")
# Keep the Columns as 'Name', 'First Name' and 'Email'

nInvestors = len(dfInvestors.index)
countloop = 0

for i in dfInvestors["First Name"]:
	print(i) #Helps you check where the code broke
	fromaddr = "username@email.com"
	toaddr = dfInvestors.loc[countloop,'Email']
	print(toaddr)
 
	msg = MIMEMultipart()
 
	msg['From'] = fromaddr
	msg['To'] = toaddr
	msg['Subject'] = "My first customized email"
 
	body1 = " Dear "+i+","/n #you can add more body
	body = body1 

	msg.attach(MIMEText(body, 'plain'))
 
	filename = i+"_Unit statement_30.06.2019.pdf" #Keep a standard file name where you can add name and store it in a folder
	attachment = open("/Users/parasmalhotra/Desktop/Unit statement/Unit statement/"+filename, "rb")
 
	part = MIMEBase('application', 'octet-stream')
	part.set_payload((attachment).read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
 
	msg.attach(part)
 
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login(fromaddr, "password")
	text = msg.as_string()
	server.sendmail(fromaddr, toaddr, text)
	server.quit()
	countloop = countloop+1

print(countloop)
