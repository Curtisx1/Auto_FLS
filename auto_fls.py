# Imported Libraries
import time
from pathlib import Path
import jira
from jira import JIRA
import requests
from itertools import cycle
import win32com.client
import re
import speech_recognition as sr
import soundfile
import os
import dotenv
from dotenv import load_dotenv
from os import environ

class Auto_FLS():
    def __init__(self):
        # Loading ENV variables
        dotenv.load_dotenv(dotenv_path=".venv/config.env")
        self.jira_connection = self.jira_connect()
    # Functions

    def jira_connect(self):
        """Returns Jira connection. Requires OAUTH values and server address. Variables retrieved from .env file"""
        #The first value will be your registered email in Jira,
        #The second value is your private API key.
        email = os.getenv("JIRA_LOGIN")
        key = os.getenv("API_KEY")
        domain = os.getenv("DOMAIN")
        jira_connection = JIRA(
            basic_auth=(email, key),
            server=domain
        )
        return jira_connection
    
    def wav_text(self):
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

    def auto_fls(self, wav_text):
        """Creates Jira ticket from Outlook email"""
        count = 0
        pause_timer = 60
        while True:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                folder = outlook.Folders.Item("fls@mhc.com")
                inbox = folder.Folders.Item("Inbox")
                archive = folder.Folders.Item("Archive")
                messages = inbox.Items
                message = messages.GetLast()
                sender = message.SenderName
                if sender == "Cisco Unity Connection Messaging System":
                    sender = "FLS Voicemail Inbox"
                creation_time = message.CreationTime.strftime(format="%H:%M-%b %d")
                subject = message.Subject
                attachment = message.Attachments
                attachment = attachment.Item(1)
                file_name = str(attachment).lower()
                path = Path("C:/Users/clemke/Python/Jira_Automation/wav_files")
                attachment.SaveASFile(f'{path}\{file_name}')
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
                jira_connection = self.jira_connect()
                issue_dict = {
                    'project': {'key': 'ITDESK'},
                    'summary': f'VM @ {creation_time} | From: {sender}',
                    'description': f'TIME CREATED:{creation_time}\nSENT FROM: {sender}\nCALLER INFO: {subject}\n{branch}\n{call_back}\n{caller}\nGoogle voice to text:\n{wav_text()}',
                    'issuetype': {'name': 'Service Request'},
                    'labels': ['Voicemail'],
                }
                new_issue = jira_connection.create_issue(fields=issue_dict)
                url = f'https://mhcworkflow.atlassian.net/rest/api/3/issue/{new_issue}/attachments'
                headers = {
                    "X-Atlassian-Token": "no-check"
                }
                files = {
                    "file": ("voicemessage.wav", open("voicemessage.wav", "rb"))
                }
                response = requests.post(url, headers=headers, files=files, auth=self.jira_connect())
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


if __name__ == "__main__":
    Auto_FLS().auto_fls()