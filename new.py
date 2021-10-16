import smtplib
from string import Template

# load sensitive data
from dotenv import load_dotenv
import os

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
