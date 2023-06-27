# Imported Libraries
import time
from pathlib import Path
from jira import JIRA
import requests
from itertools import cycle
import win32com.client
import re
import speech_recognition as sr
import soundfile
import os
from dotenv import load_dotenv
from os import environ
import glob
import logging

logging.basicConfig(filename='errors.log', level=logging.DEBUG)
# Setting up the enviroment variables. Looks for a .env file in the folder the .py file is saved in. Loads the first .env file found.
dir_path = os.path.dirname(os.path.realpath(__file__))
dotenv_files = glob.glob(os.path.join(dir_path, '*.env'))
if dotenv_files:
    load_dotenv(dotenv_files[0])


# Functions
def jira_connect():
    """Returns Jira connection. Requires OAUTH values and server address. Variables retrieved from .env file"""
    #The first value will be your registered email in Jira,
    #The second value is your private API key.
    email = os.getenv("JIRA_LOGIN")
    key = os.getenv("API_KEY")
    domain = os.getenv("DOMAIN")

    jira_connection = JIRA(basic_auth=(email, key), server=domain, options={"verify": False})
    return jira_connection

def jira_oauth():
    """Returns Oauth credentials for Jira. Variables retrieved from .env file"""
    #The first value will be your registered email in Jira,
    #The second value is your private API key.
    email = os.getenv("JIRA_LOGIN")
    key = os.getenv("API_KEY")
    jira_connection = (email, key)
    return jira_connection


def auto_fls():
    """Creates Jira ticket from Outlook email"""
    count = 0
    pause_timer = 60
    while True:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            folder = outlook.Folders.Item(os.getenv("DEFAULT_FOLDER"))
            inbox = folder.Folders.Item(os.getenv("DEFAULT_INBOX"))
            archive = folder.Folders.Item(os.getenv("DEFAULT_MOVE"))
            messages = inbox.Items
            message = messages.GetLast()
            sender = message.SenderName
            # This section can capture a specific sender and rename the variable. See below for an example.
            if sender == "Cisco Unity Connection Messaging System":
                sender = "FLS Voicemail Inbox"
            creation_time = message.CreationTime.strftime(format="%H:%M-%b %d")
            subject = message.Subject
            attachment = message.Attachments
            attachment = attachment.Item(1)
            file_name = str(attachment).lower()
            path = os.path.dirname(os.path.realpath(__file__))
            attachment.SaveASFile(f'{path}/{file_name}')
            if subject.startswith('Message from B'):
                subject_cleaned = re.split('\s+', subject)
                branch = (f'Branch Number: {subject_cleaned[2]}')
                call_back = (f'Call back number/EXT: {subject_cleaned[-1]}')
                caller = (f'Caller Name:')
            else:
                subject_cleaned = re.split('\s+', subject)
                caller = (f'Caller Name: {subject_cleaned[2]}')
                call_back = (f'Call back number/EXT: {subject_cleaned[-1]}')
                branch = ("No branch given.")
            issue_dict = {
                'project': {'key': 'ITDESK'},
                'summary': f'VM @ {creation_time} | From: {sender}',
                'description': f'TIME CREATED:{creation_time}\nSENT FROM: {sender}\nCALLER INFO: {subject}\n{branch}\n{call_back}\n{caller}\nGoogle voice to text:\n{wav_text()}',
                'issuetype': {'name': 'Service Request'},
                'labels': ['Voicemail'],
            }
            new_issue = jira_connect().create_issue(fields=issue_dict)
            url = f'{os.getenv("DOMAIN")}/rest/api/3/issue/{new_issue}/attachments'
            headers = {
                "X-Atlassian-Token": "no-check"
            }
            files = {
                "file": ("voicemessage.wav", open("voicemessage.wav", "rb"))
            }
            response = requests.post(url, headers=headers, files=files, auth=jira_oauth(), verify=False)
            print(f'Ticket {new_issue} has been created.')
            if message.UnRead:
                message.UnRead = False
            message.Move(archive)
            count += 1
        except AttributeError:
            print(f'Voicemail inbox cleared. Created {count} tickets.')
            print("---Pausing for 1 Minute---")
            time.sleep(pause_timer)
            continue


def wav_text():
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
    except Exception:
        print("If something went wrong here, its not the auto_fls, its the wav conversion/text to speech.")


if __name__ == "__main__":
    auto_fls()