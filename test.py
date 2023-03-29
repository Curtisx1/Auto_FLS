
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

def auto_fls():
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
            path = Path("C:/Users/clemke/Python/Auto_FLS-1")
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
            issue_dict = {
                'project': {'key': 'ITDESK'},
                'summary': f'VM @ {creation_time} | From: {sender}',
                'description': f'TIME CREATED:{creation_time}\nSENT FROM: {sender}\nCALLER INFO: {subject}\n{branch}\n{call_back}\n{caller}\nGoogle voice to text:\n{wav_text()}',
                'issuetype': {'name': 'Service Request'},
                'labels': ['Voicemail'],
            }
            #print(f'Ticket {new_issue} has been created.')
            if message.UnRead:
                message.UnRead = False
            message.Move(archive)
            #count += 1
        except AttributeError:
            #print(f'Voicemail inbox cleared. Created {count} tickets.')
            print("---Pausing for 1 Minute---")
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