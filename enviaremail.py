import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

email = 'seuemail@gmail.com'
password = 'XXXXX'
send_to_email = 'destinatario@gmail.com'
subject = 'This is the subject' 
message = 'This is my message'

msg = MIMEMultipart()
msg['From'] = email
msg['To'] = send_to_email
msg['Subject'] = subject

filename = r'C:\Users\Pasta1.xlsx'
attachment = open(r'C:\Users\Pasta1.xlsx', 'rb')
xlsx = MIMEBase('application','vnd.openxmlformats-officedocument.spreadsheetml.sheet')
xlsx.set_payload(attachment.read())

encoders.encode_base64(xlsx)
xlsx.add_header('Content-Dispolsition', 'attachment', filename=filename)
msg.attach(xlsx)



msg.attach(MIMEText(message, 'plain'))

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email, password)
text = msg.as_string() 
server.sendmail(email, send_to_email, text)
server.quit()
