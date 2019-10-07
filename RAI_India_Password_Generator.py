# -*- coding: utf-8 -*-
"""
Created on Friday January 4, 2019

@author: jhweaver

Generates new passwords for RAIOAM India Users
Sets new AD passwords for users
Emails manager with their new passwords

Requires: Python 3.x, 7zip (depending on architecture, check 'create_csv' function for path)

"""

from time import sleep
import random
from getpass import getpass
import subprocess
import json
import csv
import os
import smtplib
from os.path import basename
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email.mime.base import MIMEBase
from email import encoders
from pprint import pprint as p


def get_aduser():
    '''
    takes no input, runs powershell command to return RAI AD users with email addresses
    '''
    subprocess.call('powershell.exe -executionpolicy Bypass -file Get-ADUser.ps1', shell=True)
    

def import_employee_info():
    '''
    Takes powershell 'Get-ADUser' results, filters out just RAI India workers
    and stores in a list
    
    Each worker is an item in the list, comprised of three dicts:
        Username: fmlast
        Email Address: first.last@domain.com
        Password = null
    '''
    india_users = []
    with open('AD_User_info.json') as f:        
        data = json.load(f)
        for item in data:
            if 'India' in item['Distinguished Name'] and item['Enabled']:
                india_users.append({'Username': item['UserName'], 'Email': item['Email Address'], 'Password': ""})
    return india_users

def pass_gen():
    '''
    Generates and returns a password.
    Called by gen_account_creds, and is run on each user.
    '''
    alphabet = "abcdefghijklmnopqrstuvwxyz"
    symbols = "!@#$%^&*?"
    pw_length = 12
    mypw = ""

    
    for i in range(pw_length):
        next_index = random.randrange(len(alphabet))
        mypw = mypw + alphabet[next_index]
    
    # replace 1 or 2 characters with a number
    for i in range(random.randrange(1,3)):
        replace_index = random.randrange(len(mypw)//2)
        mypw = mypw[0:replace_index] + str(random.randrange(10)) + mypw[replace_index+1:]
    
    # replace 1 or 2 letters with an uppercase letter
    for i in range(random.randrange(1,3)):
        replace_index = random.randrange(len(mypw)//2,len(mypw))
        mypw = mypw[0:replace_index] + mypw[replace_index].upper() + mypw[replace_index+1:]

    # replace 1 or 2 letters with a special character
    for i in range(random.randrange(1,3)):
        replace_index = random.randrange(len(mypw)//2,len(mypw))
        mypw = mypw[0:replace_index] + str(symbols[random.randrange(9)]) + mypw[replace_index+1:]
    
    return mypw


def gen_account_creds(user_list):
    '''
    Takes user list as input, and replaces the null password dict with a 
    generated one from pass_gen()
    '''
    for user in user_list:
        user['Password'] = pass_gen()
    return user_list


def update_AD_creds(accounts):
    ''' 
    Takes user account as input, and updates user account with their new credentials in AD
    '''
    for user in accounts:
        print("Changing password for %s" % user['Username'])
        subprocess.call("powershell.exe Set-ADAccountPassword -identity %s -reset -newpassword (ConvertTo-SecureString -asplaintext '%s' -Force)" % (user['Username'], user['Password']))

        
def create_csv(account_creds):
    '''
    Creates csv file for sharing passwords to administrator
    Once csv is created, creates zip with password and destroys original
    '''
    while True:
        try:
            keys = account_creds[0].keys()
            with open('RAI_India_Credentials.csv','w') as f:
                dict_writer = csv.DictWriter(f, keys)
                dict_writer.writeheader()
                dict_writer.writerows(account_creds)
            print("\n=== Creating password protected file ===")
            with open(os.devnull) as DEVNULL:
                subprocess.call(['powershell.exe', '&', '"C:/Program Files/7-Zip/7z.exe"', 'a', 'password', 'RAI_India_Credentials.zip'] + ['RAI_India_Credentials.csv'], stdout=DEVNULL)
            sleep(2)
            print("File created\n")
            os.remove('RAI_India_Credentials.csv') 
            break
        except PermissionError:
            print("Permission denied - please confirm that the file is not in use.")
            input("Press any key to retry...")
            
            
def email_creds():
    '''
    Emails manager with password protected zip containing credentials
    
    Reference on sending emails: https://stackoverflow.com/questions/25346001/add-excel-file-attachment-when-sending-python-email
    
    '''
    print("=== Email Authentication ===")
    send_from = input("Please enter your Office 365 email address: ")
    send_to = 'receiveing_email_address'
##    send_to = 'jhweaver@intinc.com'    
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = 'Updated Jabber Credentials'
    msg.attach(MIMEText("Please see attachment for updated credentials.\n\nIf you have any questions, please email helpdesk@intinc.com."))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("RAI_India_Credentials.zip", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="RAI_India_Credentials.zip"')
    msg.attach(part)

    #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
    #SSL connection only working on Python 3+
    smtp = smtplib.SMTP(host='smtp.office365.com', port=587, timeout=120)
    smtp.starttls()
    while True:
        try:
            username = send_from
            password = getpass("Please enter password for %s: " % send_from)
            smtp.login(send_from, password)
            break
        except:
            print("\nAuthentication failed. Try again...\n")
    smtp.sendmail(send_from, send_to, msg.as_string())
    print("\nEmail successfully sent")
    smtp.quit()
    

print("You are about to change all passwords for RAI India users.")
print("Are you sure you want to continue?")
confirm_run = input("(y/n) > ")
if confirm_run == 'y':
    # Runs PS command to output AD User info to 'AD_User_info.json'
    get_aduser()   
    # Gets list of RAI India employees 
    user_list = import_employee_info()
    print("\nThe following users will have their passwords changed:\n")
    for user in user_list:
        print(user['Email'])
    print("")
    print(len(user_list), "users will be effected.")
    confirm = input("\nWould you like to continue? (y/n): ")
    if confirm == "y":
        # Returns list of employees with newly generated passwords
        print("\n=== Generating new passwords ===")
        account_creds = gen_account_creds(user_list)
##        account_creds = [{'Username':'test','Email':'jhw@intinc.com','Password':'gsdr2356sd!'}]
        print("Completed.\n")
        p(account_creds)
        sleep(2)
        # Users powershell to update AD credentials for users
        print("=== Updating passwords ===")
        update_AD_creds(account_creds)
        # Creates csv with credentials, zips it up and password protects it
        create_csv(account_creds)
        # Emails credentials to users
        email_creds()
        # Quit
        print("\nPlease upload the output 'RAI_India_Credentials.zip' to the following Sharepoint folder:\nlink-to-site")
        input("\nPress any key to quit")
    else:
        print("Cancelling...")
else:
    print("Cancelling...")
