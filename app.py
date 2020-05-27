'''
Copyright 2020 Andre Telfer
'''

import os
import sys
import logging
import shutil
import json
import yaml
import pydash
import smtplib

import PySimpleGUI as sg

from pathlib import Path
from email.message import EmailMessage

#
# LOGGING
#

# Setup Logging          
logger = logging.getLogger('simple-mailinglist')

fh = logging.FileHandler(os.path.join(sys.path[0], 'simple-mailinglist.log'))
fh.setLevel(logging.INFO)

ch = logging.StreamHandler()
ch.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

fh.setFormatter(formatter)
ch.setFormatter(formatter)

logger.addHandler(fh)
logger.addHandler(ch)

# Change the working directory to the script's directory
os.chdir(sys.path[0])

#
# Load Config
#
if (not os.path.exists('config.yaml') and os.path.exists('config.yaml.template')):
    logger.info('Creating new config file from template')
    shutil.copyfile('config.yaml.template', 'config.yaml')

if (not os.path.exists('config.yaml')):
    logger.critical('Could not find config.yaml file.')
    sys.exit(0)

config_path = Path("config.yaml")
with open(config_path, 'r') as fp:
    config = yaml.load(fp)

#
# APP 
#
sg.theme('DarkBlue3') 

'''
Send button handler
'''
def handle_send():
    logger.info('got here!')
    event, values = window.read()
    smtp_conn = smtplib.SMTP_SSL(config.get('smtp-server'), int(config.get('smtp-port')))
    smtp_conn.login(config.get('username'), config.get('password'))

    msg = EmailMessage()
    msg["From"] = f"{config.get('display-name')} <{config.get('username')}>"
    msg["To"] = ""
    msg["Subject"] = values.get('email-subject')
    msg.add_header("List-Unsubscribe", f"<mailto: {config.get('username')}?subject=unsubscribe>")
    msg.set_content(values.get('email-content', ''))

    recipients = [config.get('username')]
    if values.get('mailinglist-file'):
        if os.path.exists(values.get('mailinglist-file')):
            with open(values.get('mailinglist-file'), 'r') as fp:
                logger.info("loading recipients from file")
                recipients += fp.readlines()
    
    logger.info(values.get('mailinglist-file'))

    smtp_conn.send_message(msg, to_addrs=recipients)
    smtp_conn.quit()

'''
Save button handler
'''
def handle_save():
    event, values = window.read()

    config['saved'] = {
        'mailinglist-file': values.get('mailinglist-file', ''),
        'template-file': values.get('template-file', ''),
        'email-subject': values.get('email-subject', ''),
        'email-content': values.get('email-content', '')
    }

    with open(config_path, 'w') as fp:
        yaml.dump(config, fp)
    logger.info(f'Saved data: {json.dumps(config)}')

# Map events to handlers
dispatch_dictionary = {
    'Send': handle_send,
    'Save': handle_save
}

# All the stuff inside your window.
gap = 1
long_input_length = 60 
short_secondary_input_length = 10
short_input_length = long_input_length - short_secondary_input_length - gap

label_size = (15, 1)
content_size = (long_input_length, 20)
long_input_size = (long_input_length + 3, 1)
short_input_size = (short_input_length, 1)
short_secondary_input_size = (short_secondary_input_length, 1)

layout = [  
    [sg.Text('For more settings use the config.yaml file.')],
    [sg.Text(f'Sending from', size=label_size), sg.Text(config["username"])],
    [sg.Text('Mailing List File', size=label_size), sg.Input(pydash.get(config, 'saved.mailinglist-file', ''), key='mailinglist-file', size=short_input_size), sg.FileBrowse(size=short_secondary_input_size)],
    [sg.Text('Template File (Opt.)', size=label_size), sg.Input(pydash.get(config, 'saved.template-file', ''), key='template-file', size=short_input_size), sg.FileBrowse(size=short_secondary_input_size)],
    [sg.Text('Subject', size=label_size), sg.InputText(pydash.get(config, 'saved.email-subject', ''), key='email-subject', size=long_input_size)],
    [sg.Text('Content', size=label_size), sg.Multiline(pydash.get(config, 'saved.email-content', ''), key='email-content', size=content_size)],
    [sg.Button('Send'), sg.Button('Save'), sg.Button('Cancel')]    
]

# Create the Window
window = sg.Window('Mailing List Tool', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()

    # Handle close window
    if event in ('Cancel', 'Quit'):   
        break

    # Handle buttons
    if event in dispatch_dictionary:
        func_to_call = dispatch_dictionary[event]   # get function from dispatch dictionary
        func_to_call()
    else:
        print('Event {} not in dispatch dictionary'.format(event))

window.close()