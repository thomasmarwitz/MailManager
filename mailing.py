import smtplib
from string import Template
from collections import namedtuple
import re
import logging
import sys

# load sensitive data
from dotenv import load_dotenv
import os
import pandas as pd

logging.basicConfig(filename="mailing.log", format='%(levelname)s:%(asctime)s:%(message)s', level=logging.DEBUG)
load_dotenv()

# alvindeguzman@hotmail.de

class Message:

    def __init__(self, file_name: str) -> None:
        try:
            self.txt: str= open(file_name).read()
        except FileNotFoundError:
            logging.critical(f"couldn't find/load the file you specified: {file_name}")
            sys.exit(-1)
            
        self.subject: str = self.getSubject()

    def replaceValues(self, data: dict[str, str]) -> str:
        text: str = self.txt # don't modify original txt
        for key, value in data.items():
            text = text.replace("{{" + key + "}}", value)
        if "{{" in text:
            logging.warning("there may exist parameters in the message that haven't been replaced")

        return text

    def getSubject(self) -> str:
        pattern: str = r"<!--\s+Betreff=\[([^\]]+)\]\s+-->"
        
        mo: re.Match = re.search(pattern, self.txt)
        if mo:
            return mo.group(1)
        else:
            logging.critical("Couldn't extract Subject for Email")
            sys.exit(-1)


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
            self.sender:            str = os.environ["SENDER"]
        except KeyError:
            print("Please specify EMAIL and PASSWORD in the .env file")
            raise SystemExit

        self.server = smtplib.SMTP(host=self.host_address, port=self.port)
        self.server.set_debuglevel(0)
        self.server.starttls()
        logging.debug(f"login with {self.email}")
        self.server.login(self.email, self.password)
        logging.debug("logged in")

    
    def send_test_message(self):
        msg = MIMEMultipart()
        
        msg['From'] =   formataddr((str(Header('Von Sender', 'utf-8')), self.email))
        msg['To'] =     formataddr((str(Header('An Empfaenger', "utf-8")), self.email))
        msg['Subject'] = self.message.subject #"Automated Test E-Mail"
        
        msg.attach(MIMEText(
            self.message.replaceValues({}), 
            "html")
        )
        logging.debug("sending test message")
        self.server.send_message(msg)
        logging.debug("message sent")

    def _create_msg(self, toPerson: Person, attachedPerson: Person) -> MIMEMultipart:
        msg = MIMEMultipart()
        
        msg['From'] =   formataddr((str(Header(self.sender, 'utf-8')), self.email))
        msg['To'] =     formataddr((str(Header(toPerson.name, "utf-8")), toPerson.email))
        msg['Subject'] = self.message.subject
        
        # body
        msg.attach(MIMEText(
            self.message.replaceValues(dict(
                NAME=toPerson.name,
                PARTNER=attachedPerson.name,
                EMAIL_PARTNER=attachedPerson.email,
            )),
            "html"
        ))
        return msg

    def send_invitation(self, toPerson: Person, attachedPerson: Person):
        logging.debug(f"sending message to {toPerson.name}")
        self.server.send_message(self._create_msg(toPerson, attachedPerson))
        logging.debug(f"message sent")

class Manager:

    def __init__(self, excel_file: str, message_file: str):
        logging.debug("loading excel sheet")
        try:
            self.df: pd.DataFrame = pd.read_excel(excel_file)
        except FileNotFoundError:
            logging.critical(f"couldn't find/load the file you specified: {excel_file}")
            sys.exit(-1)

        logging.debug("finished loading excel sheet")
        self.mailer: Mailer = Mailer(message_file)

    def send_test(self) -> None:
        self.mailer.send_test_message()
        
    def process_rowwise(self):
        # iterate over matches
        logging.info("processing pairs:")
        for (name1, email1, name2, email2) in zip(self.df["Name1"], self.df["Email1"], self.df["Name2"], self.df["Email2"]):
            self._process_pairs(
                Person(name1, email1),
                Person(name2, email2)
            )
        logging.info("finished processing paris")

    def _process_pairs(self, person1: Person, person2: Person):
        logging.info(f"process pair: {person1.name} & {person2.name}")
        # message person 1
        self.mailer.send_invitation(person1, person2)

        # message person 2
        self.mailer.send_invitation(person2, person1)
        

class Question:

    def __init__(self, output_func, input_func, ignore_case: bool) -> None:
        self.output = output_func
        self.input = input_func
        self.ignore_case: bool = ignore_case
    
    def ask_user(self, question: str, accepted: list[str]) -> str:
        while True: # ask until valid input
            self.output(question)
            answer: str = self.input("> ")
            if self._is_accepted(answer, accepted):
                return answer
            else:
                self.output(f"wrong input, must be from: {accepted}\n")

    def _is_accepted(self, answer: str, accepted: list[str]):
        if self.ignore_case:
            return answer.lower() in [s.lower() for s in accepted]
        else:
            return answer in accepted
            



logging.info("STARTED PROGRAM")
manager: Manager = Manager(
    excel_file="Zuordnung.xlsx", 
    message_file="message.html"
)

question: Question = Question(print, input, ignore_case=True)

print("loaded data:\n", manager.df)
data_valid: str = question.ask_user("Is this data correct?", ["y", "n"])
if data_valid == "n":
    logging.critical("the process was cancelled by the user after data validation")
    sys.exit(-1)

send_test: str = question.ask_user("Send test message?", ["y", "n"])
if send_test == "y":
    manager.send_test()
    send_all: str = question.ask_user("Test message ok? Start sending messages?", ["y", "n"])
    if send_all == "n":
        logging.critical("the process was cancelled by the user after test message")
        sys.exit(-1)


manager.process_rowwise()
logging.info("FINISHED PROGRAM")