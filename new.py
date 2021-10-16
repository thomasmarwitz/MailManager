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


print(email, password)