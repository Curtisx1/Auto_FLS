# DESCRIPTION
""" Program to automate the creation of Jira tickets from an Outlook Inbox. Specifically voicemail messages. """
# NOTES
# Version:        1
# Author:         Curtis Lemke
# Creation Date:  12/9/2022
# Change Date :   3/17/2023
# Purpose/Change:

# Imported Libraries
# Some of these are unused/will be used in future revisions.
import io
import time
import datetime
import csv
from csv import writer, reader
import json
from pathlib import Path
import jira
from jira import JIRA
import requests
from itertools import cycle
import subprocess
import win32com.client
import re
import speech_recognition as sr
import soundfile
import numpy as np
import os
import dotenv
from dotenv import load_dotenv
from os import environ

class Auto_FLS():
    def __init__(self):
        # Loading ENV variables
        dotenv.load_dotenv(dotenv_path="C:/Users/clemke/Python/Auto_FLS-1/.venv/config.env")

    # Functions
    def jira_oauth(self):
        """Returns Oauth credentials for Jira. Variables retrieved from .env file"""
        #The first value will be your registered email in Jira,
        #The second value is your private API key.
 
        jira_connection = ('user@domain.com', 'APIKEY')
        return jira_connection


    def jira_connect(self):
        """Returns Jira connection. Requires OAUTH values and server address. Variables retrieved from .env file"""
        #The first value will be your registered email in Jira,
        #The second value is your private API key.
        email = os.getenv(self.JIRA_LOGIN)
        key = os.getenv(self.API_KEY)
        domain = os.getenv(self.DOMAIN)
        jira_connection = JIRA(
            basic_auth=(email, key),
            server=domain
        )
        return print(jira_connection)


    def auto_fls(self):
        """Creates Jira ticket from Outlook email"""
        # Count tracks the number of tickets created. Resets when closed/stopped.
        count = 0
        # The pause lenth can be changed in the .env file
        pause_timer = 60
        while True:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                # You can change the default folder in the .env file
                folder = outlook.Folders.Item("desired_inbox@domain.com")
                # You can change the default folder in the .env file
                inbox = folder.Folders.Item("Inbox")
                # You can change the default folder in the .env file
                archive = folder.Folders.Item("Archive")
                messages = inbox.Items
                # Using GetLast, we address the voicemails/emails that are sent in first as they have priority.
                message = messages.GetLast()
                # Captures the name of the user. This value is their display name set in Outlook. Defaults to email if not set.
                sender = message.SenderName
                # You can use the code below to capture emails sent with a specific subject line and change the sender. An example would be from a no-reply email with a subject of "NO REPLY"
                if sender == "Desired Sender":
                    sender = "New Sender"
                # Captures the time the email was sent
                creation_time = message.CreationTime.strftime(format="%H:%M-%b %d")
                # Captures the subject
                subject = message.Subject
                # Captures any attachments. Saves to a desired filepath
                attachment = message.Attachments
                attachment = attachment.Item(1)
                file_name = str(attachment).lower()
                path = Path("C:/desired_save_location")
                attachment.SaveASFile(f'{path}\{file_name}')
                # The following if/else statement will format the subject line depending on the start line of the subject. In our use case, messages from outside callers always started with "Message from".
                # This can be changed within the .env file or removed entirely if not required.
                if subject.startswith('Message from'):
                    subject_cleaned = re.split('\s+', subject)
                    branch = (f'Branch Number: {subject_cleaned[2]}')
                    call_back = (f'Call back number/EXT: {subject_cleaned[-1]}')
                    caller = (f'Caller Name:')
                else:
                    subject_cleaned = re.split('\s+', subject)
                    caller = (f'Caller Name: {subject_cleaned[2]}')
                    call_back = (f'Call back number/EXT: {subject_cleaned[-1]}')
                    branch = ("No branch given.")
                # Calls the Jira Connect function
                #jira_connection = jira_connect()
                # Dictionary of the values captured above
                # The 'project', 'issuetype', and 'labels' can be configured in the .env file
                issue_dict = {
                    'project': {'key': 'ITDESK'},
                    'summary': f'VM @ {creation_time} | From: {sender}',
                    'description': f'TIME CREATED:{creation_time}\nSENT FROM: {sender}\nCALLER INFO: {subject}\n{branch}\n{call_back}\n{caller}\nGoogle voice to text:\n{wav_text()}',
                    'issuetype': {'name': 'Service Request'},
                    'labels': ['Voicemail'],
                }
                new_issue = jira_connection.create_issue(fields=issue_dict)
                # The URL is set within the .env file. This URL should be retrieved from your Jira dashboard.
                url = f'https://mhcworkflow.atlassian.net/rest/api/3/issue/{new_issue}/attachments'
                headers = {
                    "X-Atlassian-Token": "no-check"
                }
                files = {
                    "file": ("voicemessage.wav", open("voicemessage.wav", "rb"))
                }
                # Sends the packaged data to the Jira API
                response = requests.post(url, headers=headers, files=files, auth=jira_oauth())
                # Prints the ticket number that has been created.
                print(f'Ticket {new_issue} has been created.')
                # Marks the email as read and moves it into the 'archive' folder. 
                # This folder location can be changed in the .env file
                if message.UnRead:
                    message.UnRead = False
                message.Move(archive)
                count += 1
            # Captures the attribute error that occurs when the inbox is empty. Pauses the script for 1 minutes before running again. 
            except AttributeError:
                print(f'Voicemail inbox cleared. Created {count} tickets.')
                print("---Pausing for 1 Minute---")
                time.sleep(pause_timer)
                continue


    def wav_text(self):
        """Takes the .wav file from the auto_fls function. Passes it to the Google voice to text API. Returns the translated text to be used in the description of the Jira ticket"""
        # The default filename is 'voicemessage.wav' this may be different depending on your Outlook setup.
        # This value can be set within the .env file
        data, samplerate = soundfile.read('voicemessage.wav')
        soundfile.write('new.wav', data, samplerate, subtype='PCM_16')
        r = sr.Recognizer()
        hellow = sr.AudioFile('new.wav')
        with hellow as source:
            audio = r.record(source)
        try:
            s = r.recognize_google(audio, show_all = True, )
            results = s['alternative'][0]
            return results['transcript']
        # This will capture the error that happens when the voicemail message is blank/empty
        except Exception:
            print("Error: Empty voicemail message.")

    def testing(self):
        key = os.getenv('KEY')
        print(key)

if __name__ == "__main__":
    