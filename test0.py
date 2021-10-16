import smtplib
from string import Template

host_adress = "mail.gmx.com"
port = 25

mail_adress = "thomas.marwitz@gmx.de"
pw = "JVDYFF63LFZPIIQT6GCL"

s = smtplib.SMTP(host=host_adress, port=port)
s.set_debuglevel(1)
s.starttls()
s.login(mail_adress, pw)


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# For each contact, send the email:

msg = MIMEMultipart() # create a message

# add in the actual person name to the message template

# setup the parameters of the message
msg['From'] = mail_adress
msg['To'] = "thomasmarwitz3@gmail.com"
msg['Subject'] = "Python SMTP Test"

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

# add in the message body

#t = read_template("msg.html")
#message = t.substitute(LINK="""<a href="https://www.google.com">Link</a>""")

message = open("Hallo.html").read()

msg.attach(MIMEText(message, 'html')) # "html?"

    ## Link:

# send the message via the server set up earlier.
s.send_message(msg)
