import smtplib
from string import Template
from collections import namedtuple
import re

# load sensitive data
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()

# alvindeguzman@hotmail.de

class Message:

    def __init__(self, file_name: str) -> None:
        self.txt: str= open(file_name).read()
        self.subject: str = self.getSubject()

    def replaceValues(self, data: dict[str, str]) -> str:
        text: str = self.txt # don't modify original txt
        for key, value in data.items():
            text = text.replace("{{" + key + "}}", value)
        return text

    def getSubject(self) -> str:
        pattern: str = r"<!--\s+Betreff=\[([^\]]+)\]\s+-->"
        
        mo: re.Match = re.search(pattern, self.txt)
        if mo:
            return mo.group(1)
        else:
            return SystemExit("Couldn't extract Subject for Email")


Person = namedtuple('Person', ['name', 'email'])


# email stuff
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

class Mailer:

    def __init__(self, message_file): # required: message
        self.message: Message = Message(message_file)

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
        msg['Subject'] = self.message.subject #"Automated Test E-Mail"
        
        msg.attach(MIMEText(
            self.message.replaceValues({}), 
            "html")
        )

        self.server.send_message(msg)

    def _create_msg(self, toPerson: Person, attachedPerson: Person) -> MIMEMultipart:
        msg = MIMEMultipart()
        
        msg['From'] =   formataddr((str(Header('Von Sender', 'utf-8')), self.email))
        msg['To'] =     formataddr((str(Header(toPerson.name, "utf-8")), toPerson.email))
        msg['Subject'] = self.message.subject
        
        # body
        msg.attach(MIMEText(
            self.message.replaceValues(dict(
                NAME=toPerson.name,
                PARTNER=attachedPerson.name,
                EMAIL_PARTNER=attachedPerson.email,
            ))
        ))

    def send_invitation(self, toPerson: Person, attachedPerson: Person):
        self.server.send_message(self._create_msg(toPerson, attachedPerson))

class Manager:

    def __init__(self, excel_file: str):
        self.df: pd.DataFrame = pd.read_excel(excel_file)
        self.mailer: Mailer = Mailer()
        
    def process_rowwise(self):
        # iterate over matches
        for (name1, email1, name2, email2) in zip(self.df["Name1"], self.df["Email1"], self.df["Name2"], self.df["Email2"]):
            self._process_pairs(
                Person(name1, email1),
                Person(name2, email2)
            )

    def _process_pairs(self, person1: Person, person2: Person):
        # message person 1
        self.mailer.send_invitation(person1, person2)

        # message person 2
        self.mailer.send_invitation(person2, person1)
        


manager: Manager = Manager("Zuordnung.xlsx")
