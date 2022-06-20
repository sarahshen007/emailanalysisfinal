# For storing data into local database
from os import remove
import sqlite3
import pandas as pd

# For analyzing data
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.corpus import words

# For analyzing emails from outlook
import datetime
import emails

import win32com
from win32com import client
from tkinter import filedialog
from tkinter import *
from bs4 import BeautifulSoup
from colorama import Fore

# VARIABLES
database = "example.db"                         # database file
product_correspondence = {}                     # issue - product relationship dict

correct_words = words.words()                   # spell check words
stop_words = set(stopwords.words('english'))    # stopwords set
relevant_words = set()
relevant_words.update(("account",
                        "order", 
                        "email", 
                        "store", 
                        "password", 
                        "website", 
                        "online", 
                        "reset", 
                        "parts", 
                        "site", 
                        "number", 
                        "address", 
                        "purchase", 
                        "card", 
                        "time", 
                        "log", 
                        "vehicle", 
                        "code", 
                        "sign",
                        "signed", 
                        "in", 
                        "app", 
                        "rewards", 
                        "paypal", 
                        "phone", 
                        "cart", 
                        "find", 
                        "rebate", 
                        "change", 
                        "stock",
                        "add", 
                        "item", 
                        "check",
                        "auto", 
                        "link", 
                        "items", 
                        "search", 
                        "credit", 
                        "ordered", 
                        "received", 
                        "info", 
                        "login", 
                        "access", 
                        "error", 
                        "place", 
                        "apply", 
                        "car", 
                        "coupon", 
                        "send", 
                        "message", 
                        "fit", 
                        "buy", 
                        "battery", 
                        "oil", 
                        "headlight", 
                        "day", 
                        "next", 
                        "money", 
                        "charged", 
                        "receive", 
                        "today", 
                        "complete", 
                        "purchase", 
                        "purchased", 
                        "shipping", 
                        "pay", 
                        "discount", 
                        "emails", 
                        "payment", 
                        "stores",
                        "delivery", 
                        "product", 
                        "checkout", 
                        "service", 
                        "part", 
                        "deals", 
                        "discover", 
                        "location", 
                        "store", 
                        "signed",
                        "signing"
                        "sign",
                        "submit",
                        "rebate",
                        "rebates",
                        "date",
                        "dates"
                        "products",
                        "searching",
                        "searched",
                        "searches",
                        "shelf",
                        "applied",
                        "apply",
                        "history",
                        "receipt",
                        "receipts",
                        "finds",
                        "credits",
                        "hub",
                        "info",
                        "information",
                        "vehicle",
                        "truck",
                        "motorcycle",
                        "motorcycles",
                        "gas",
                        "purchases",
                        "purchasing",
                        "buy",
                        "bought",
                        "Hyundai",
                        "Honda",
                        "stock",
                        "%"
                        ))


# FUNCTIONS FOR RETRIEVING OUTLOOK EMAILS

# generate list of emails from yesterday
def get_outlook_emails(feedback_folder):
    # Get the date
    d = (datetime.date.today() - datetime.timedelta (days=1)).strftime("%d-%m-%y")

    # Connect to outlook account
    app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    account = accounts[0]

    # Grab folder for CS EMAILS
    root_folder = app.Folders(account.DisplayName)
    emails_folder = emails.get_folder_by_name(feedback_folder, root_folder)

    # Get emails from folder and populate list of emails
    messages = emails_folder.Items

    # Get the email/s
    msg = messages.GetLast()

    # List of emails from day
    new_messages = []

    # Log keeping track of email objects
    emails_log = []

    # Parse spreadsheet to predict issue
    prev_data = storage.generate_issue_data()

    # Loop through emails
    while msg:
        # Get email date 
        date = msg.SentOn.strftime("%d-%m-%y")

        # Get Subject Line of email
        sjl = msg.Subject

        # Only add emails from yesterday           
        if d == date:
            new_messages.append(msg)

            # Dictionary to store message info
            info = {}

            # Date message was received
            date = str(msg.SentOn).split(' ')[0]
            date = datetime.datetime.strptime(date, '%Y-%m-%d').date()
            info['date'] = date

            # Remove unecessary characters from msg html
            regex = msg.HTMLBody.replace('\r', '').replace('\n', '') 

            # Parse into html using soup
            soup = BeautifulSoup(regex, "html.parser") 

            # Create list of category + values
            texts = str(soup.find_all('font')[0].encode_contents(encoding='utf-8')).strip('b').strip('\'').strip('\"').replace('<br/>', '\n')
            texts = emails.replaceCharacters(texts)
            texts = texts.strip().split('\n')
            texts = list(filter(None, texts))

            # Create list of pairs to populate info dictionary
            pairs = []
            
            # Edit list for unwanted extra elements caused by extra break elements
            lastKey = ""
            for data in texts:
                pair = data.split(':', 1)
                if len(pair) == 1:
                    info[lastKey] = info[lastKey] + pair[0]
                elif len(pair) == 2: 
                    lastKey = pair[0].strip()
                    info[lastKey] = pair[1].strip()

            # Generate summary of comment
            summary = info['Comment Value']

            predicted_issue = generate_issue(summary, prev_data)

            info['Issue Summary'] = predicted_issue[0]
            info['Product'] = predicted_issue[1]
            
            # Make new email object with info
            newEmail = emails.emailCreator(info)

            # Add email object to emails log
            emails_log.append(newEmail)
            
        msg = messages.GetPrevious()
    return emails_log


# FUNCTIONS FOR INTERACTING WITH DB
# retrieve list of all entries
def get_emails():
    emails = []
    with sqlite3.connect(database) as db:
        cursor = db.cursor()
        sql = "SELECT * FROM feedback"
        cursor.execute(sql)
        emails = cursor.fetchall()
    return emails

# retrieve list of (issue, product, comment)
def get_issues():
    issues = []
    with sqlite3.connect(database) as db:
        cursor = db.cursor()
        sql = "SELECT Issue, Product, Comment FROM feedback"
        cursor.execute(sql)
        issues = cursor.fetchall()
    return issues

# put spreadsheet data in db
def xl_db(excel):
    emails = pd.read_excel(
    excel, 
    sheet_name='CS Feedback',
    header=0)

    with sqlite3.connect(database) as db:
        cursor = db.cursor()

        cursor.execute(
        """
        CREATE TABLE if not exists feedback (
            Date DATE,
            Issue TEXT,
            Product TEXT,
            Name TEXT,
            Email TEXT,
            Comment TEXT,
            IP TEXT,
            Session TEXT,
            Followup BOOL
            );
        """)

        emails.to_sql('feedback', db, if_exists='append', index=False)

        cursor.close()

# add given list of emails to database
def add_emails(email_list):
    with sqlite3.connect(database) as db:
        cursor = db.cursor()
        for email in email_list:
            sql = f'INSERT INTO feedback (Date, Issue, Product, Name, Email, Comment, IP, Session, Followup) VALUES({email.date}, {email.issueSummary}, {email.product}, {email.name}, {email.customerEmail}, {email.comment}, {email.ipAddress}, {email.cookies}, {email.followup})'
            cursor.execute(sql)


# FUNCTIONS FOR TEXT MANIPULATION & ISSUE PREDICTION

# remove irrelevant words 
def remove_stop_words(text):
    text = re.sub(r'[^\w\s]','',text)
    word_tokens = word_tokenize(text)

    filtered_sentence = []

    for w in word_tokens:
        if w.lower() not in stop_words:
            filtered_sentence.append(w)

    return filtered_sentence

# generate issue prediction data
def generate_issue_data():
    global relevant_words
    global product_correspondence

    data = {}

    issues_comp = get_issues()
    prev_issue = 'General Inquiry'

    comments = [remove_stop_words(str(entry[2]).lower()) for entry in issues_comp]
    all_words = ''
    for comment in comments:
        all_words += (' ').join(comment) + " "
    all_words = word_tokenize(all_words)
    all_words_dist = nltk.FreqDist(w.lower() for w in all_words if w not in stop_words)

    relevant_words.update(all_words_dist.most_common(400))

    for i in range(len(issues_comp)-1, 0, -1):
        entry = issues_comp[i]
        issue = str(entry[0]).lower()
        product = str(entry[1]).lower()
        comment = remove_stop_words(str(entry[2]))

        if issue == 'none':
            issue = prev_issue
        else:
            prev_issue = issue
        
        if issue not in product_correspondence:
            product_correspondence[issue] = product

        issues_comp[i] = (issue, product_correspondence[issue], comment)
        relevant_words.update(issue.replace('-', '').split(' '))
        

    for i in range(len(issues_comp)):
        entry = issues_comp[i]
        issue = entry[0]
        product = entry[1]
        comment = entry[2]

        if not issue in data:
            data[issue] = {}

        num_words = 0

        for word in comment:
            if word in data[issue]:
                data[issue][word] += 1
                num_words += 1
            elif word in relevant_words:
                data[issue][word] = 1
                num_words += 1
        
        for word in data[issue]:
            if not num_words == 0:
                data[issue][word] /= num_words
        
    return data

# predict issue
def generate_issue(text, data):
    text = " ".join(remove_stop_words(text))
    text = word_tokenize(text)
    text_dist = nltk.FreqDist(word.lower() for word in text)

    issues_list = list(data.keys())
    words_list = text_dist.keys()
    weight_comp = []

    for i in range(len(issues_list)):
        issue = issues_list[i]
        weight = 0

        for word in data[issue]:
            if word in issues_list and word in words_list:
                weight += data[issue][word] * text_dist.freq(word) * 1.1
            elif word in words_list:
                weight += data[issue][word] * text_dist.freq(word)
        
        weight_comp.append(weight)

    max_index = 0
    max_val = 0

    for i in range(len(weight_comp)):
        if weight_comp[i] > max_val:
            max_val = weight_comp[i]
            max_index = i

    if max_val == 0:
        return 'General Inquiry'

    return [issues_list[max_index].title(), product_correspondence[issues_list[max_index].lower()]]

