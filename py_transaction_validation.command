#!/usr/bin/env python

# library

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
    df = pd.read_excel(root.file_path, sheet_name='Transaction template', header=1)

elif file_type == "csv":
    df = pd.read_csv(root.file_path)

else:
    print("Not a compatible file type.")
    print(" ")


required_columns = ['Agent first and lastname *', 'Transaction type (b,s) *', 'Client first name *',
       'Client last name *', 'Lead Source / type']

required_present = []
for i in df.columns:
    if i in required_columns:
        required_present.append(i)

print("COLUMN CHECK:")

if len(required_present)/len(required_columns) != 1:
    text_color = 'red'
    missing_required_columns = []
    for i in required_columns:
        if i not in required_present:
            missing_required_columns.append(i)
    required_text = colored(f'Missing Required Column(s): {missing_required_columns} \n', 'red')
else:
    text_color = 'green'
    required_text = " "

print(f"Number of columns: {len(df.columns)}")
print(f"Number of not NA columns: {len(df.dropna(axis = 1, how = 'all').columns)}")
print("Number of required columns: " + colored(f"{len(required_present)}/{len(required_columns)}", text_color))
print(required_text)




# Validate Transaction Imports'

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
                         'Client last name', 'Client email', 'Property address_1', 'Property address_2', 'Property city', 'Property state', 'Mobile phone',
                         'Home phone', 'Lead Source / type',
                         'Vendor attorney', 'Vendor closing gifts', 'Vendor electrician', 'Vendor escrow company', 'Vendor flooring', 'Vendor HVAC',
                         'Vendor handyman', 'Vendor landscaper', 'Vendor home inspection company', 'Vendor home warranty company', 'Vendor insurance company',
                         'Vendor mortgage company', 'Vendor moving company', 'Vendor pest terminator', 'Vendor photographer', 'Vendor plumber',
                         'Vendor relocation company', 'Vendor sign installer', 'Vendor solar', 'Vendor surveyor', 'Vendor title company', 'Transaction ID', 'Notes']

date_columns_legacy = ['Agent paid date',
                       '1st Appointment date', 'Signed date', 'Under contract date', 'MLS live date (Listing date)',
                       'Closed (Settlement) date''Offer reference Date', 'Due diligence deadline',  'Lead date', ]

int_columns_legacy = ['Property postal', 'Showings (# of)']

float_columns_legacy = ['Transaction amount', 'GCI (Gross Commission Amount)', 'Gross agent paid income',
                        'Listing amount']

string_columns_current = ['Agent first and lastname *', 'Transaction type (b,s) *',
                          'Referral (Yes/No)', 'Rental (Yes/No)', 'Client first name *',
                          'Client last name *', 'Client email', 'Property address_1',
                          'Property address_2', 'Property city', 'Property state',
                          'Mobile phone', 'Home phone',
                          'Lead Source / type',
                          'Vendor attorney',
                          'Vendor closing gifts', 'Vendor electrician', 'Vendor escrow company',
                          'Vendor flooring', 'Vendor HVAC', 'Vendor handyman',
                          'Vendor landscaper', 'Vendor home inspection company',
                          'Vendor home warranty company', 'Vendor insurance company',
                          'Vendor mortgage company', 'Vendor moving company',
                          'Vendor pest terminator', 'Vendor photographer', 'Vendor plumber',
                          'Vendor relocation company', 'Vendor sign installer', 'Vendor solar',
                          'Vendor surveyor', 'Vendor title company', 'Transaction ID']

date_columns_current = [
    'Agent paid date', '1st Appointment date', 'Signed date',
    'Under contract date', 'MLS live date (Listing date)',
    'Closed (Settlement) date *', 
    'Offer reference Date', 'Due diligence deadline',
    'Lead date']

int_columns_current = ['Property postal', 'Showings (# of)']

float_columns_current = ['Transaction amount *',
                         'GCI (Gross Commission Amount) *', 'Gross agent paid income *', 'Listing amount']


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
    print(colored("Using legacy temaplate.", "yellow"))
    string_columns = string_columns_legacy
    date_columns = date_columns_legacy
    int_columns = int_columns_legacy
    float_columns = float_columns_legacy
    template_version = 'legacy'

print(" ")

print(f"File path: {root.file_path}")

print(" ")


def validate_transaction_data(df, string_columns=string_columns, date_columns=date_columns, int_columns=int_columns, float_columns=float_columns, template_version = template_version):

    for i in string_columns:
        if i in df.columns:
            df[i] = df[i].astype(str)
            if df[i].dtype != "O":
                print(colored("ERROR: ", 'red'), colored('String Column\t', 'yellow'),
                    "Some data in column", colored(f"{i}", 'yellow'),  "is formatted incorrectly")
            else:
                pass
        else:
            print(colored("String", 'yellow', attrs = ['bold']) + " column " + colored(f"{i}", 'magenta', attrs = ['bold']) + " not included in data.")

    for i in date_columns:
        # df[i] = pd.to_datetime(df[i])
        if i in df.columns:
            if df[i].dtype != "datetime64[ns]":
                print(colored("FORMAT ERROR: ", 'red'), colored('Date Column\t', 'green'),
                      "Some data in column", colored(f"{i}", 'green'),  "is formatted incorrectly")
            else:
                pass
        else:
            print(colored("Date", 'green', attrs = ['bold']) + " column " + colored(f"{i}", 'magenta', attrs = ['bold']) + " not included in data.")

    for i in int_columns:
        if i in df.columns:
            if df[i].dtype != "int64" and df[i].dtype != "int32":
                print(colored("FORMAT ERROR: ", 'red'), colored('Integer Column\t', 'cyan'),
                      "Some data in column", colored(f"{i}", 'cyan'), "is formatted incorrectly")
            else:
                pass
        else:
            print(colored("Integer", 'cyan', attrs = ['bold']) + " column " + colored(f"{i}", 'magenta', attrs = ['bold']) + " not included in data.")

    for i in float_columns:
        if i in df.columns:
            if df[i].dtype != "float64" and df[i].dtype != "float32":
                print(colored("FORMAT ERROR: ", 'red'), colored('Float Column\t', 'blue'),
                      "Some data in column", colored(f"{i}", 'blue'), "is formatted incorrectly")
            else:
                pass
        else:
            print(colored("Float", 'blue', attrs = ['bold']) + " column " + colored(f"{i}", 'magenta', attrs = ['bold']) + " not included in data.")

    # check for correct email format
    if len(df[(df['Client email'].notna()) & (df['Client email'] != 'nan')][~(df['Client email'].str.contains('@'))]) > 0:
        print(colored('Some emails are not formatted correctly.',
              'yellow', attrs=['bold']))
        print(tabulate(df[(df['Client email'].notna()) & (df['Client email'] != 'nan')][~(
            df['Client email'].str.contains('@'))][['Agent first and lastname *', 'Client email']]))
    else:
        pass
    

    if template_version == 'current':
        # check for duplicate rows
        if len(df.duplicated()) > 0:
            print(colored('Duplicate rows found:', 'yellow', attrs=['bold']))
            print(tabulate(df[df.duplicated()]['Agent first and lastname *']))

        if len(df[((df['Transaction type (b,s) *'].isna()) | (df['Transaction type (b,s) *'] == 'nan')) & (df['Agent first and lastname *'].notna())]) > 0:
            len_trans_missing = len(df[((df['Transaction type (b,s) *'].isna()) | (
                df['Transaction type (b,s) *'] == 'nan')) & (df['Agent first and lastname *'].notna())])
            print(colored(
                f"{len_trans_missing} Transaction type(s) missing.", 'red', attrs=['bold']))

        if len(df[(df['Transaction type (b,s) *'] != "s") & (df['Transaction type (b,s) *'] != "b")]):
            wrong_trans = len(df[(df['Transaction type (b,s) *'] != "s")
                            & (df['Transaction type (b,s) *'] != "b")])
            print(colored(
                f"{wrong_trans} Transaction Type(s) incorrectly formatted or missing", "red", attrs=['bold']))
            print(tabulate(df[(df['Transaction type (b,s) *'] != "s") & (df['Transaction type (b,s) *']
                != "b")][['Agent first and lastname *', 'Transaction type (b,s) *']]))

    elif template_version == 'legacy':
            # check for duplicate rows
        if len(df.duplicated()) > 0:
            print(colored('Duplicate rows found:', 'yellow', attrs=['bold']))
            print(tabulate(df[df.duplicated()]['Agent first and lastname *']))

        if len(df[((df['Transaction type (b,s) *'].isna()) | (df['Transaction type (b,s) *'] == 'nan')) & (df['Agent first and lastname *'].notna())]) > 0:
            len_trans_missing = len(df[((df['Transaction type (b,s) *'].isna()) | (
                df['Transaction type (b,s) *'] == 'nan')) & (df['Agent first and lastname *'].notna())])
            print(colored(
                f"{len_trans_missing} Transaction type(s) missing.", 'red', attrs=['bold']))

        if len(df[(df['Transaction type (b,s) *'] != "s") & (df['Transaction type (b,s) *'] != "b")]):
            wrong_trans = len(df[(df['Transaction type (b,s) *'] != "s")
                            & (df['Transaction type (b,s) *'] != "b")])
            print(colored(
                f"{wrong_trans} Transaction Type(s) incorrectly formatted or missing", "red", attrs=['bold']))
            print(tabulate(df[(df['Transaction type (b,s) *'] != "s") & (df['Transaction type (b,s) *']
                != "b")][['Agent first and lastname *', 'Transaction type (b,s) *']]))
    
    else:
        print(colored("Neither current nor legacy template.", 'orange'))



validate_transaction_data(df)