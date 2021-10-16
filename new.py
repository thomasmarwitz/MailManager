import smtplib
from string import Template

# load sensitive data
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()

# alvindeguzman@hotmail.de

class Message:

    def __init__(self, file_name: str) -> None:
        self.txt: str= open(file_name).read()

    def replaceValues(self, data: dict[str, str]) -> str:
        text: str = self.txt # don't modify original txt
        for key, value in data.items():
            text = text.replace("{{" + key + "}}", value)
        return text

class Manager:

    def __init__(self, excel_file: str):
        self.df: pd.DataFrame = pd.read_excel(excel_file)
        
    def process_rowwise(self):
        for (name1, email1, name2, email2) in zip(self.df["Name1"], self.df["Email1"], self.df["Name2"], self.df["Email2"]):
            pass


# email stuff
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

class Mailer:

    def __init__(self):
        try:
            self.email:             str = os.environ["EMAIL"]
            self.password:          str = os.environ["PASSWORD"]
            self.host_address:      str = os.environ["HOST_ADDRESS"]
            self.port:              str = os.environ["PORT"]
        except KeyError:
            print("Please specify EMAIL and PASSWORD in the .env file")
            raise SystemExit

        self.server = smtplib.SMTP(host=self.host_address, port=self.port)
        self.server.set_debuglevel(0)
        self.server.starttls()
        self.server.login(self.email, self.password)

        
    
    def send_test_message(self):
        msg = MIMEMultipart()
        
        msg['From'] =   formataddr((str(Header('Von Sender', 'utf-8')), self.email))
        msg['To'] =     formataddr((str(Header('An Empfaenger', "utf-8")), self.email))
        msg['Subject'] = "Automated Test E-Mail"
        
        msg.attach(MIMEText(
            Message("message.html").replaceValues({}),
            "html")
        )

        self.server.send_message(msg)



        


manager: Manager = Manager("Zuordnung.xlsx")



