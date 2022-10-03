#!/usr/bin/env python


# Transaction Import Template Validation Tool Version 2.0
# Created by Josh Spradlin
# 2022-08-26

# library

import os
from termcolor import colored
import pwd # needed if the document is stored on computer instead of Google Drive
import tkinter as tk
from tkinter import filedialog
import pandas as pd
pd.set_option('display.max_columns', None) # keeps pandas from truncating columns
import numpy as np
import clipboard
pd.options.display.max_colwidth = 1000
import warnings
warnings.filterwarnings('ignore')
from tabulate import tabulate
from datetime import datetime
import pytz
import subprocess
import platform
# keeps pandas from truncating columns
pd.set_option('display.max_columns', None)
pd.options.display.max_colwidth = 1000
warnings.filterwarnings('ignore')

# Import File using dialog box

def raise_app(root: tk):
    root.attributes("-topmost", True)
    if platform.system() == 'Darwin':
        tmpl = 'tell application "System Events" to set frontmost of every process whose unix id is {} to true'
        script = tmpl.format(os.getpid())
        output = subprocess.check_call(['/usr/bin/osascript', '-e', script])
    root.after(0, lambda: root.attributes("-topmost", False))


check_for_google_drive = True
# you can change this to start at a different folder.
local_folder = "Downloads"

root = tk.Tk()
# os.system('''/usr/bin/osascript -e 'tell app "Finder" to set frontmost of process "Python" to true' ''')
# root.wm_attributes('-topmost', 1)
raise_app(root)
root.withdraw()
root.lift()

file_path = ''  # filedialog.askopenfilename()

if check_for_google_drive:
    if 'Google Drive.app' in os.listdir("/Applications/"):
        root.file_path = filedialog.askopenfilename(
            initialdir="/Volumes/GoogleDrive/My Drive/IMPORTS", title="SELECT the Task List File")
    else:
        root.file_path = filedialog.askopenfilename(initialdir="/Users/"+pwd.getpwuid(
            os.getuid()).pw_name+"/{local_folder}", title="SELECT the CSV file from Jira")
else:
    root.file_path = filedialog.askopenfilename(initialdir="/Users/"+pwd.getpwuid(
        os.getuid()).pw_name+"/{local_folder}", title="SELECT the CSV file from Jira")

file_type = root.file_path.split(".")[-1]

print(" ")
print("Importing a " + colored(f"{file_type}", 'cyan') + " file.")
print(" ")

if file_type == "xlsx":
    try:
        df = pd.read_excel(root.file_path, sheet_name='Transaction template', header=1)
    except:
        df = pd.read_excel(root.file_path)

elif file_type == "csv":
    df = pd.read_csv(root.file_path)

else:
    print("Not a compatible file type.")
    print(" ")

# Required Columns
current_required_columns = ['Agent first and lastname *', 'Transaction type (b,s) *', 'Client first name *',
       'Client last name *', 'Lead Source / type']

current_required_present = []
for i in df.columns:
    if i in current_required_columns:
        current_required_present.append(i)

legacy_required_columns = ['Agent first and lastname', 'Transaction type (b,s)', 'Client first name',
       'Client last name', 'Lead Source / type']

legacy_required_present = []
for i in df.columns:
    if i in legacy_required_columns:
        legacy_required_present.append(i)


# Current Columns
current_columns = ['Agent first and lastname *', 'Transaction type (b,s) *',
                   'Referral (Yes/No)', 'Rental (Yes/No)', 'Client first name *',
                   'Client last name *', 'Transaction amount *',
                   'GCI (Gross Commission Amount) *', 'Gross agent paid income *',
                   'Agent paid date', '1st Appointment date', 'Signed date',
                   'Under contract date', 'MLS live date (Listing date)',
                   'Closed (Settlement) date *', 'Client email', 'Property address_1',
                   'Property address_2', 'Property city', 'Property state',
                   'Property postal', 'Mobile phone', 'Home phone', 'Showings (# of)',
                   'Offer reference Date', 'Due diligence deadline', 'Lead Source / type',
                   'Lead date', 'Listing amount', 'Vendor attorney',
                   'Vendor closing gifts', 'Vendor electrician', 'Vendor escrow company',
                   'Vendor flooring', 'Vendor HVAC', 'Vendor handyman',
                   'Vendor landscaper', 'Vendor home inspection company',
                   'Vendor home warranty company', 'Vendor insurance company',
                   'Vendor mortgage company', 'Vendor moving company',
                   'Vendor pest terminator', 'Vendor photographer', 'Vendor plumber',
                   'Vendor relocation company', 'Vendor sign installer', 'Vendor solar',
                   'Vendor surveyor', 'Vendor title company', 'Transaction ID', 'Notes']


string_columns_current = ['Agent first and lastname *', 'Transaction type (b,s) *',
                          'Referral (Yes/No)', 'Rental (Yes/No)', 'Client first name *',
                          'Client last name *',
                           'Property city', 'Property state',
                          'Lead Source / type',
                          'Vendor attorney',
                          'Vendor closing gifts', 'Vendor electrician', 'Vendor escrow company',
                          'Vendor flooring', 'Vendor HVAC', 'Vendor handyman',
                          'Vendor landscaper', 'Vendor home inspection company',
                          'Vendor home warranty company', 'Vendor insurance company',
                          'Vendor mortgage company', 'Vendor moving company',
                          'Vendor pest terminator', 'Vendor photographer', 'Vendor plumber',
                          'Vendor relocation company', 'Vendor sign installer', 'Vendor solar',
                          'Vendor surveyor', 'Vendor title company']


mix_sting_columns_current = ['Property address_1', 'Property address_2', 'Notes']

email_string_columns_current = ['Client email']

phone_string_columns_current = ['Mobile phone', 'Home phone']

date_columns_current = [
    'Agent paid date', '1st Appointment date', 'Signed date',
    'Under contract date', 'MLS live date (Listing date)',
    'Closed (Settlement) date *', 
    'Offer reference Date', 'Due diligence deadline',
    'Lead date']

int_columns_current = ['Property postal', 'Showings (# of)', 'Transaction ID']

float_columns_current = ['Transaction amount *',
                         'GCI (Gross Commission Amount) *', 'Gross agent paid income *', 'Listing amount']




# Legacy Columns
legacy_columns = ['Agent first and lastname', 'Transaction type (b,s)', 'Referral (Yes/No)', 'Rental (Yes/No)', 'Client first name',
                  'Client last name', 'Transaction amount', 'GCI (Gross Commission Amount)', 'Gross agent paid income', 'Agent paid date',
                  '1st Appointment date', 'Signed date', 'Under contract date', 'MLS live date (Listing date)', 'Closed (Settlement) date',
                  'Client email', 'Property address_1', 'Property address_2', 'Property city', 'Property state', 'Property postal', 'Mobile phone',
                  'Home phone', 'Showings (# of)', 'Offer reference Date', 'Due diligence deadline', 'Lead Source / type', 'Lead date', 'Listing amount',
                  'Vendor attorney', 'Vendor closing gifts', 'Vendor electrician', 'Vendor escrow company', 'Vendor flooring', 'Vendor HVAC',
                  'Vendor handyman', 'Vendor landscaper', 'Vendor home inspection company', 'Vendor home warranty company', 'Vendor insurance company',
                  'Vendor mortgage company', 'Vendor moving company', 'Vendor pest terminator', 'Vendor photographer', 'Vendor plumber',
                  'Vendor relocation company', 'Vendor sign installer', 'Vendor solar', 'Vendor surveyor', 'Vendor title company', 'Transaction ID', 'Notes']


string_columns_legacy = ['Agent first and lastname', 'Transaction type (b,s)', 'Referral (Yes/No)', 'Rental (Yes/No)', 'Client first name',
                         'Client last name', 'Property city', 'Property state', 'Lead Source / type',
                         'Vendor attorney', 'Vendor closing gifts', 'Vendor electrician', 'Vendor escrow company', 'Vendor flooring', 'Vendor HVAC',
                         'Vendor handyman', 'Vendor landscaper', 'Vendor home inspection company', 'Vendor home warranty company', 'Vendor insurance company',
                         'Vendor mortgage company', 'Vendor moving company', 'Vendor pest terminator', 'Vendor photographer', 'Vendor plumber',
                         'Vendor relocation company', 'Vendor sign installer', 'Vendor solar', 'Vendor surveyor', 'Vendor title company']

mix_sting_columns_legacy = ['Property address_1', 'Property address_2', 'Notes']

date_columns_legacy = ['Agent paid date',
                       '1st Appointment date', 'Signed date', 'Under contract date', 'MLS live date (Listing date)',
                       'Closed (Settlement) date''Offer reference Date', 'Due diligence deadline',  'Lead date', ]

email_string_columns_legacy = ['Client email']

phone_string_columns_legacy = ['Mobile phone', 'Home phone']

int_columns_legacy = ['Property postal', 'Showings (# of)', 'Transaction ID']

float_columns_legacy = ['Transaction amount', 'GCI (Gross Commission Amount)', 'Gross agent paid income',
                        'Listing amount']




print(" ")
if len(df.columns[df.columns.str.contains("\*")]) > 0:

    print(colored('Using Current Template', 'green'))
    # df.columns = df.columns.str.replace("\*","").str.strip()
    string_columns = string_columns_current
    date_columns = date_columns_current
    int_columns = int_columns_current
    float_columns = float_columns_current
    template_version = 'current'
else:
    print(colored("Not Current Temaplate.", "yellow"))
    string_columns = string_columns_legacy
    date_columns = date_columns_legacy
    int_columns = int_columns_legacy
    float_columns = float_columns_legacy
    template_version = 'legacy'

print(" ")

print(f"File path: {root.file_path}")

print(" ")

df = df.drop_duplicates().reset_index(drop=True) # remove duplicate rows

def validate_transaction_data(df, template_version = template_version):

    if template_version == 'current':
        required = current_required_columns
        present = current_required_present
        string_columns = string_columns_current
        date_columns = date_columns_current
        phone_columns = phone_string_columns_current
        mix_columns = mix_sting_columns_current
        email_columns = email_string_columns_current
        float_columns = float_columns_current
        integer_columns = int_columns_current

    elif template_version == 'legacy':
        required = legacy_required_columns
        present = legacy_required_present
        string_columns = string_columns_legacy
        date_columns = date_columns_legacy
        phone_columns = phone_string_columns_legacy
        mix_columns = mix_sting_columns_legacy
        email_columns = email_string_columns_legacy
        float_columns = float_columns_legacy
        integer_columns = int_columns_legacy

    print(colored('Checking present and required', 'yellow'))

    print("COLUMN CHECK:")

    if len(present)/len(required) != 1:
        text_color = 'red'
        missing_required_columns = []
        for i in required:
            if i not in present:
                missing_required_columns.append(i)
        required_text = colored(f'Missing Required Column(s): {missing_required_columns} \n', 'red')
    else:
        text_color = 'green'
        required_text = " "

    print(f"Number of columns: {len(df.columns)}")
    print(f"Number of not NA columns: {len(df.dropna(axis = 1, how = 'all').columns)}")
    print("Number of required columns: " + colored(f"{len(present)}/{len(required)}", text_color))
    print(required_text)

    # df = df.dropna(axis = 1, how = 'all')
    # print("Empty columns dropped.")
    # print(' ')


    print('CHECKING COLUMN FORMAT')
    print(" ")
    for i in df.columns:
        if i in string_columns:
            try:
                if df[i].dtype != "O":
                    print(colored("Format Error: ", 'red') + colored("String ", 'magenta') + f"{i}")
                    
            except:
                print(colored("Format Error: ", 'red') + colored("String ", 'magenta') + f"{i}. Not Object.")

        elif i in date_columns:
            try:
                pd.to_datetime(df[i])
            except:
                print(colored("Format Error: ", 'red') + colored("Date ", 'cyan') + f"{i}")

        # elif i in phone_columns and len(df) - df[i].str.contains('^\d{10}$').sum() >= 0:
        #     print(colored("Format Error: ", 'red') + colored("Phone Number ", 'green') + f"{i}")

        elif i in mix_columns and df[i].dtype != "O":
            print(colored("Format Error: ", 'red') + colored("String ", 'magenta') + f"{i}")

        elif i in integer_columns:
            try:
                df[i].astype(int)
            except:
                print(colored("Format Error: ", 'red') + colored("Integer ", 'blue') + f"{i}")
        
        elif i in float_columns:
            try:
                df[i].astype(object)
            except:
                print(colored("Format Error: ", 'red') + colored("Float ", 'blue') + f"{i}")

        elif i in email_columns:
            try:
                df[i].astype(str)
                if len(df) -  df[i].str.contains("@").sum() != 0:
                    print(colored("Format Error: ", 'red') + colored("Email ", 'green') + f"{i}")
            except:
                print(colored("Format Error: ", 'red') + colored("Email ", 'green') + f"{i}")
                

validate_transaction_data(df)