import smtplib
import email

host_adress = "ssl.bildung-rp.de"
port = 25

mail_adress = "susannemarwitz@gs-sued.bildung.rp"
pw = "BobbyliXXX72"

s = smtplib.SMTP(host=host_adress, port=port)
s.starttls()
s.login(mail_adress, pw)


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# For each contact, send the email:
for name, email, link in zip(names, emails, links):
    msg = MIMEMultipart() # create a message

    # add in the actual person name to the message template
    message = message_template.substitute(PERSON_NAME=name.title()) # Python Template String!

    # setup the parameters of the message
    msg['From'] = mail_adress
    msg['To'] = email
    msg['Subject'] = betreff

    # add in the message body
    msg.attach(MIMEText(message, 'plain')) # "html?"

    # send the message via the server set up earlier.
    s.send_message(msg)

    del msg
