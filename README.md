# MailManager

Sending Mails to notify pairs of people that they were matched with each other. Send a message to each person with the name of the other person and the contact details of the other person.

# Installation

Create a virtual environment with `python -m venv venv` and activate it with `source venv/bin/activate`. Install the requirements with `pip install -r requirements.txt`.

# Configuration

Fill in the `.env_TEMPLATE` with your own mail address and password. Rename it to `.env` and save it in the root directory of the project.

# Providing the matching data

The matching data is provided via an excel file. Take a look at the `matching_template.xlsx` file to see how the data should be structured. The first row is the header row specifying the column names. Each row then contains a pair of people that should meet each other. The first group of columns contains the data of the first person, the second group of columns contains the data of the second person. Save the file as `matching.xlsx` in the root directory of the project.

# Running the script

Run the script with `python mailing.py --matching_file 'matching.xlsx' --message_template 'message.html'`.
