#    Exchange email deadline Extraction
#    By      : Mohammed Al-jawaheri
#    Email   : Mobj@cmu.edu -- m_aljawaheri@outlook.com
#    Student -- CarnegieMellon University
from exchangelib import Credentials, Account
from re import*


#
## Handles all deadline collection. Will use RegEx through methods
## to extract information from from account.inbox item objects
class deadline_collector:
    def __init__(self, reg_ex):
        self.search_regex = reg_ex   # list of regexs specified by user (dates,etc)
    # Use methods to check / extract deadline
    def extract_deadline(self, email):
        # search for every regex
        for regex in self.search_regex:
            result = search(regex, email.text_body)
            if result and email.subject not in important_emails:
                # add to important emails
                important_emails[email.subject] = important_email_info(
                                                    email.sender, email.subject,
                                                    email.text_body, result)
            else:
                # if regex not in body search in subjectline
                result = search(regex, email.subject)
                if result and email.subject not in important_emails:
                    # add to important emails
                    important_emails[email.subject] = important_email_info(
                                                    email.sender, email.subject,
                                                    email.text_body, result)


#
## Class stores all important email-info per email. Subject line, dates and
## email body. Will store each object in an important emails dictionary
class important_email_info:
    def __init__(self, sender, subjectLine, email_body, deadline):
        self.sender = sender
        self.subject = subjectLine
        self.email_body = email_body
        self.deadline = deadline


# dictionary to check if email receieved hasnt already been processed
important_emails = {}  # reset every 24 hrs
# dictionary to run DLcollector filter over to filter important emails
all_emails = {}        # reset every 24 hrs
my_deadline_collector = deadline_collector(
                       ["January \d+",
                        "February \d+",
                        "March \d+",
                        "April \d+",
                        "May \d+",
                        "June \d+",
                        "July \d+",
                        "August \d+",
                        "September \d+",
                        "October \d+",
                        "November \d+",
                        "December \d+"])


def collect_important_emails():
    # Collecting info from exchange email:
    email = input("Enter your email: ")
    password = input("Enter your password: ")
    credentials = Credentials(email, password)   # Note: later change this to user input, possibly tkinter interface ?
    account = Account(email, credentials=credentials, autodiscover=True)

    # Get emails
    for item in account.inbox.all().order_by('-datetime_received')[:5]:
        if item.subject not in all_emails and item.subject != None:
            all_emails[item.subject] = item

    # Use deadline collector to filter all emails and append to important emails
    for email in all_emails:
        my_deadline_collector.extract_deadline(all_emails[email])
    return important_emails
