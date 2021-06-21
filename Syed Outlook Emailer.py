import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from getpass import getpass
import os.path

print("Syed's Outlook Emailer\n")


email_sender = input("Your Outlook.com Email: ")
email_pass = getpass("Your Outlook.com Password: ")
email_recipient = input("Recipient's Email Address: ")
email_subject = input("Email subject: ")
email_message = input("Email message: ")


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location = ''):

    


    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login(email_sender, email_pass)
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('Sent!')
        server.quit()
    except:
        print("SMPT server connection error")
    return True

send_email(email_recipient,
                email_subject,
                email_message)