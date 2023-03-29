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

class Auto_FLS:
    def __init__(self):
        # Loading ENV variables
        load_dotenv()

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
        return print(jira_connection)
    
if __name__ == "__main__":
    obj = Auto_FLS()
    obj.jira_connect()