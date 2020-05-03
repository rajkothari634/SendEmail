import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import xlrd

email_user = ''
email_password = ''
loc = ("emailids.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(email_user,email_password)
for i in range(sheet.nrows): 
    email_send = sheet.cell_value(i, 0)
    subject = 'Warmest congratulations on your achievement!'
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject

    body = """Slash-o-Code was organised on Hackerearth on 30th April, 2020 by Hackslash (Mozilla Campus Club), National Institute of Technology, Patna. It received a whopping response of 200+ participants from 40+ colleges throughout India. The 3 hours inter college-coding contest gained a lot of followers and recognition. This could have not been possible without the utmost dedication and involvement of all participants. Hackslash and HackerEarth extend their thankful wishes for your dedication and efforts. Please find the attached certificate of participation/achievement. For being a part of more of such events, please connect with us on:
    Facebook: https://facebook.com/hackslash.nitp
    Twitter: https://twitter.com/hackslash_nitp
    Thank you.
    Hackslash Club Mozilla Campus Club
    NIT Patna"""
    msg.attach(MIMEText(body,'plain'))

    filename=  email_send + '.pdf'
    attachment  =open(filename,'rb')

    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= "+filename)

    msg.attach(part)
    text = msg.as_string()
    
    server.sendmail(email_user,email_send,text)
server.quit()

