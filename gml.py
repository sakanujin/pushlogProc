import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def send_email (sender_email, receiver_email, username, password, subject, body, file_path):
    # create e-mail
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To']   = receiver_email
    message['Subject']   = subject  

    # body sentence 
    message.attach(MIMEText(body, 'plain'))

    # set attach file
    with open(file_path, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name='sample.xlsx')
        part ['Content-Dispostion'] = f'attachment; filename=file.txt'
        message.attach(part)

    # connect mail server
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(username, password)
        text = message.as_string()

        # send mail
        server.sendmail(sender_email, receiver_email, text)
        print("send mail")
