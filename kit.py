import smtplib
from string import Template

host_adress = "smtp.kit.edu"
port = 587

mail_adress = "uctwi@student.kit.edu"
pw = "AwiT,dnam,e-BdSa,swu-DLwea,udwws-Bms,dkvuhs2020"


# loading username, pw

s = smtplib.SMTP(host=host_adress, port=port)
s.set_debuglevel(1)
s.starttls()
s.login(mail_adress, pw)

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

msg = MIMEMultipart()
msg['From'] = mail_adress
msg['To'] = "thomasmarwitz3@gmail.com"  # alvindeguzman@hotmail.de
msg['Subject'] = "Python SMTP Test"     # Kennenlern Treffen blabla

def replaceValues(txt: str, data: dict[str, str]) -> str:
    for key, value in data.items():
        txt = txt.replace("{{" + key + "}}", value)
    return txt

def get_message(file_name: str) -> str:
    return replaceValues(
        txt=open(file_name).read(),
        data={"name" : "world!"},
    )

#msg.attach(MIMEText("Hallo!", 'plain')) # "html?"
msg.attach(MIMEText(get_message("message.html"), "html"))

s.send_message(msg)
