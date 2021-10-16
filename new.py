import smtplib
from string import Template

# load sensitive data
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()

try:
    email:      str = os.environ["EMAIL"]
    password:   str = os.environ["PASSWORD"]
except KeyError:
    print("Please specify EMAIL and PASSWORD in the .env file")
    raise SystemExit

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

manager: Manager = Manager("Zuordnung.xlsx")


# email stuff
from email.header import Header
from email.utils import formataddr

formataddr((str(Header('Someone Somewhere', 'utf-8')), 'xxxxx@gmail.com'))