import smtplib
from collections import namedtuple
import re
import logging
import sys
from typing import Callable
import click

# load sensitive data
from dotenv import load_dotenv
import os
import pandas as pd


def setup_logger(file, level=logging.INFO, log_to_stdout=True):
    logger = logging.getLogger()
    logger.setLevel(level)
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"
    )

    if log_to_stdout:
        stdout_handler = logging.StreamHandler(sys.stdout)
        stdout_handler.setFormatter(formatter)
        logger.addHandler(stdout_handler)

    file_handler = logging.FileHandler(file)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


logger = setup_logger("mailing.log", level=logging.DEBUG, log_to_stdout=True)

load_dotenv()

FIRST_PAIR_NAME = "Name1"
SECOND_PAIR_NAME = "Name2"
FIRST_PAIR_EMAIL = "Email1"
SECOND_PAIR_EMAIL = "Email2"


class Message:
    def __init__(self, file_name: str) -> None:
        try:
            self.txt: str = open(file_name).read()
        except FileNotFoundError:
            logger.critical(f"couldn't find/load the file you specified: {file_name}")
            sys.exit(-1)

        self.subject: str = self.get_subject()

    def replace_values(self, data: dict[str, str]) -> str:
        text: str = self.txt  # don't modify original txt
        for key, value in data.items():
            text = text.replace("{{" + key + "}}", str(value))
        if "{{" in text:
            logger.warning(
                "there may exist parameters in the message that haven't been replaced"
            )

        return text

    def get_subject(self) -> str:
        pattern: str = r"<!--\s+Betreff=\[([^\]]+)\]\s+-->"

        mo: re.Match = re.search(pattern, self.txt)
        if mo:
            return mo.group(1)
        else:
            logger.critical("Couldn't extract Subject for Email")
            sys.exit(-1)


Person = namedtuple(
    "Person",
    [
        "name",
        "email",
        "hynr",
        "studiengang",
        "semester",
        "stipstatus",
        "teilname",
        "präsenz",
        "interessen",
    ],
)


# email stuff
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr


class Mailer:
    def __init__(self, message_file):  # required: message
        self.message: Message = Message(message_file)

        try:
            self.email: str = os.environ["EMAIL"]
            self.password: str = os.environ["PASSWORD"]
            self.host_address: str = os.environ["HOST_ADDRESS"]
            self.port: str = os.environ["PORT"]
            self.sender: str = os.environ["SENDER"]
        except KeyError:
            print("Please specify EMAIL and PASSWORD in the .env file")
            raise SystemExit

        self.server = smtplib.SMTP(host=self.host_address, port=self.port)
        self.server.set_debuglevel(0)
        self.server.starttls()
        logger.debug(f"login with {self.email}")
        self.server.login(self.email, self.password)
        logger.debug("logged in")

    def send_test_message(self):
        logger.debug("generating test message")
        msg = MIMEMultipart()

        msg["From"] = formataddr((str(Header(self.sender, "utf-8")), self.email))
        msg["To"] = formataddr((str(Header("An Test Empfaenger", "utf-8")), self.email))
        msg["Subject"] = self.message.subject  # "Automated Test E-Mail"

        msg.attach(MIMEText(self.message.replace_values({}), "html"))
        logger.debug("sending test message")
        self.server.send_message(msg)
        logger.debug("message sent")

    def _create_msg(
        self, to_person: Person, attached_person: Person, is_test: bool = False
    ) -> MIMEMultipart:
        msg = MIMEMultipart()

        msg["From"] = formataddr((str(Header(self.sender, "utf-8")), self.email))
        msg["To"] = formataddr((str(Header(to_person.name, "utf-8")), to_person.email))
        msg["Subject"] = self.message.subject

        # body
        message_text: str = self.message.replace_values(
            dict(
                NAME=to_person.name,
                PARTNER=attached_person.name,
                EMAIL_PARTNER=attached_person.email,
                HY_NR_PARTNER=attached_person.hynr,
                STUDIENGANG=attached_person.studiengang,
                SEMESTER=attached_person.semester,
                STIPSTATUS=attached_person.stipstatus,
                INTERESSEN=attached_person.interessen,
                PRÄSENZ=attached_person.präsenz,
                TEILNAME=attached_person.teilname,
            )
        )
        msg.attach(MIMEText(message_text, "html"))

        if is_test:  # only message preview
            msg["Text"] = message_text
        return msg

    def send_invitation(self, to_person: Person, attached_person: Person):
        logger.debug(f"sending message to {to_person.name}")
        self.server.send_message(
            self._create_msg(
                to_person,
                attached_person,
            )
        )
        logger.debug("message sent")


def iter_row_wise(df: pd.DataFrame):
    n = len(df.index)
    EMPTY_LIST = ["" for _ in range(n)]
    for i in range(n):
        yield (
            Person(
                df[FIRST_PAIR_NAME][i],
                df[FIRST_PAIR_EMAIL][i],
                df.get("HYNR1", EMPTY_LIST)[i],
                df.get("Studiengang1", EMPTY_LIST)[i],
                df.get("Semester1", EMPTY_LIST)[i],
                df.get("Stipstatus1", EMPTY_LIST)[i],
                df.get("Teilname1", EMPTY_LIST)[i],
                df.get("Präsenz1", EMPTY_LIST)[i],
                df.get("Interessen1", EMPTY_LIST)[i],
            ),
            Person(
                df[SECOND_PAIR_NAME][i],
                df[SECOND_PAIR_EMAIL][i],
                df.get("HYNR2", EMPTY_LIST)[i],
                df.get("Studiengang2", EMPTY_LIST)[i],
                df.get("Semester2", EMPTY_LIST)[i],
                df.get("Stipstatus2", EMPTY_LIST)[i],
                df.get("Teilname2", EMPTY_LIST)[i],
                df.get("Präsenz2", EMPTY_LIST)[i],
                df.get("Interessen2", EMPTY_LIST)[i],
            ),
        )


class Manager:
    def __init__(self, excel_file: str, message_file: str):
        logger.debug("loading excel sheet")
        try:
            self.df: pd.DataFrame = pd.read_excel(excel_file)
        except FileNotFoundError:
            logger.critical(f"couldn't find/load the file you specified: {excel_file}")
            sys.exit(-1)

        logger.debug("finished loading excel sheet")
        self.mailer: Mailer = Mailer(message_file)

    def send_test(self) -> None:
        self.mailer.send_test_message()

    def process_rowwise(self):
        # iterate over matches
        logger.info("processing pairs:")
        for row_obj in iter_row_wise(self.df):
            self._process_pairs(
                person1=row_obj[0],
                person2=row_obj[1],
            )

        logger.info("finished processing pairs")

    def _process_pairs(self, person1: Person, person2: Person):
        logger.info(f"process pair: {person1.name} & {person2.name}")
        # message person 1
        self.mailer.send_invitation(person1, person2)

        # message person 2
        self.mailer.send_invitation(person2, person1)

    def get_pairs(self) -> list[str]:
        return [
            name1 + " - " + name2
            for name1, name2 in zip(self.df[FIRST_PAIR_NAME], self.df[SECOND_PAIR_NAME])
        ]

    def validate_message(self) -> str:
        self.mailer.message.get_subject()

        person_storage = next(iter_row_wise(self.df))
        to_person = person_storage[0]
        attached_person = person_storage[1]

        msg: MIMEMultipart = self.mailer._create_msg(
            to_person, attached_person, is_test=True
        )

        return f'From:    {msg["From"]}\nTo:      {msg["To"]}\nSubject: {msg["Subject"]}\nBody:\n{msg["Text"]}'


class Question:
    def __init__(self, output_func, input_func, ignore_case: bool) -> None:
        self.output: Callable = output_func
        self.input: Callable = input_func
        self.ignore_case: bool = ignore_case

    def ask_user(self, question: str, accepted: list[str]) -> str:
        while True:  # ask until valid input
            self.output(question)
            answer: str = self.input("> ")
            if self._is_accepted(answer, accepted):
                print()
                return answer
            else:
                self.output(f"wrong input, must be from: {accepted}\n")

    def _is_accepted(self, answer: str, accepted: list[str]):
        if self.ignore_case:
            return answer.lower() in [s.lower() for s in accepted]
        else:
            return answer in accepted


@click.command()
@click.option(
    "--matching_file",
    default="matching.xlsx",
    help="The file providing the matches (name, email, ...)",
)
@click.option("--message_template", default="message.html", help="The message to send")
def main(matching_file: str, message_template) -> str:
    manager: Manager = Manager(excel_file=matching_file, message_file=message_template)
    print(manager.df)
    print()

    question: Question = Question(print, input, ignore_case=True)

    ########## MESSAGE VALIDATION ##############
    print("Preview of first message:\n" + "=" * 80)
    print(manager.validate_message())
    print("=" * 80)

    message_valid: str = question.ask_user("Is this message correct?", ["y", "n"])
    if message_valid == "n":
        logger.critical(
            "the process was cancelled by the user after message validation"
        )
        return "sending cancelled (message probably invalid)"

    ########## DATA VALIDATION ##############
    print("loaded data:\n" + "\n".join(manager.get_pairs()))
    data_valid: str = question.ask_user("Is this data correct?", ["y", "n"])
    if data_valid == "n":
        logger.critical("the process was cancelled by the user after data validation")
        return "sending cancelled (data probably invalid)"

    ########## MAILING VALIDATION ##############
    send_test: str = question.ask_user(
        "Send test message? [to your own address]", ["y", "n"]
    )
    if send_test == "y":
        manager.send_test()
        send_all: str = question.ask_user(
            "Test message ok? Start sending messages?", ["y", "n"]
        )
        if send_all == "n":
            logger.critical("the process was cancelled by the user after test message")
            return "sending cancelled (test message probably invalid"

    manager.process_rowwise()
    return "SUCCESS"


if __name__ == "__main__":
    logger.info("STARTED PROGRAM")
    print("\n ==>", main())
    logger.info("FINISHED PROGRAM")
