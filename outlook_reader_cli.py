import win32com.client
import datetime
import argparse
import re
import itertools
import pandas as pd
from dataclasses import dataclass
import sys
from striprtf.striprtf import rtf_to_text
import importlib

# Create an argument parser
parser = argparse.ArgumentParser(description="Access outlook email items from today")
# Add an argument for the subject, to or from
parser.add_argument("-s", "--subject", help="The subject of the email")
parser.add_argument("-t", "--to", help="The recipient of the email")
parser.add_argument("-f", "--sender", help="The sender of the email")
parser.add_argument("-fold", "--folder", help="folder to search")
parser.add_argument("-list", "--listfolders", help="list folders", action="store_true")
parser.add_argument("-r", "--raw", help="raw rtf output", action="store_true")
parser.add_argument("-rtf", "--rtfparse", help="rtf parse output - helpful for findstr/grep", action="store_true")
# Parse the arguments
args = parser.parse_args()

def get_outlook_emails_from_today(subject=None, to=None, from_=None):
    # Create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Get the default namespace
    namespace = outlook.GetNamespace("MAPI")
    # Get the default inbox folder
    inbox = namespace.GetDefaultFolder(6) # 6 is the index for inbox
    # Get the items in the inbox folder
    items = inbox.Items
    # Restrict the items to only those received today
    today = datetime.date.today()
    start = today.strftime("%m/%d/%Y")
    start_time = (today + datetime.timedelta(days=-1)).strftime("%m/%d/%Y")
    end = (today + datetime.timedelta(days=1)).strftime("%m/%d/%Y")
    restriction = "[ReceivedTime] >= '" + start_time + "' AND [ReceivedTime] <= '" + end + "'"
    items = items.Restrict(restriction)
    # Loop through the items and print some information
    result = []
      # Loop through the items and append a tuple with some information to the list$
    for item in items:
        result.append((item.Subject, item.SenderName, item.ReceivedTime, item.Body, ))
      # Return the list of tuples$
    return result



def read_outlook_subfolder_items(subfolder_name):
    # Create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Get the default namespace
    namespace = outlook.GetNamespace("MAPI")
    # Get the default inbox folder
    inbox = namespace.GetDefaultFolder(6) # 6 is the index for inbox
    # Get the subfolder by name
    subfolder = inbox.Folders(subfolder_name)
    # Get the items in the subfolder
    items = subfolder.Items
    today = datetime.date.today()
    start = today.strftime("%m/%d/%Y")
    start_time = (today + datetime.timedelta(days=-1)).strftime("%m/%d/%Y")
    end = (today + datetime.timedelta(days=1)).strftime("%m/%d/%Y")
    restriction = "[ReceivedTime] >= '" + start_time + "' AND [ReceivedTime] <= '" + end + "'"
    items = items.Restrict(restriction)
    # Create an empty list to store the tuples
    result = []
    # Loop through the items and append a tuple with some information to the list
    for item in items:
        result.append((item.Subject, item.SenderName, item.ReceivedTime, item.Body, ))
    # Return the list of tuples
    return result

def list_outlook_subfolders():
    # Create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Get the default namespace
    namespace = outlook.GetNamespace("MAPI")
    # Get the default inbox folder
    inbox = namespace.GetDefaultFolder(6) # 6 is the index for inbox
    # Get the items in the subfolder
    items = inbox.Folders
    # Create an empty list to store the tuples
    result = []
    # Loop through the items and append a tuple with some information to the list
    for item in items:
        print(item)
    # Return the list of tuples
    return result

def internal_compare_route(item_tuple, compare, pattern):
    # Check the compare keyword and get the corresponding member of the tuple
    if compare == "subject":
        item = item_tuple[0]
    elif compare == "SenderName":
        item = item_tuple[1]
    elif compare == "ReceivedTime":
        item = item_tuple[2]
    elif compare == "Body":
        item = item_tuple[3]
    else:
        return False # Invalid compare keyword
    # Use re.search to check if the pattern matches the item
    return bool(re.search(pattern, item))



@dataclass
class options:
    subject: str
    sender: str
    to: str
    folder: str
    listfolder: bool
    raw : bool
    rtf : bool

def print_body_data( my_options, item ):
    # item is a tuple of the email.
    if my_options.raw:
        print(item[3])
        return
    if my_options.rtf:
        my_text = rtf_to_text( item[3])
        print(my_text)


        


def extract_table_ews(get_sublist, create_df_from_list, item):
    print( item )
    list_body = item[3].split("\r\n")
    print( create_df_from_list( get_sublist( list_body)))

if __name__ == "__main__":

    my_options = options(subject = "", sender ="", to="", folder ="", listfolder = False, raw = False, rtf = True)    
#    email_list = get_outlook_emails_from_today(subject = args.subject, to=args.to, from_ = args.sender)
    plug = importlib.import_module("plugin_ews")

    cls = getattr( plug, "EWS")

    print( cls )
    # Check the arguments and perform different actions
    if args.subject:
    # Do something with the subject argument
        print("Subject:", args.subject)
        my_options.subject = args.subject
    if args.to:
    # Do something with the to argument
        print("To:", args.to)
        my_options.to = args.to
    if args.sender:
    # Do something with the sender argument
        print("Sender:", args.sender)
        my_options.sender = args.sender
    if args.folder:
    # Do something with the folder argument
        print("Folder:", args.folder)
        my_options.folder = args.folder
    if args.listfolders:
    # Do something with the listfolders argument
        print("List folders:", args.listfolders)
        my_options.listfolder = True
    if args.raw:
        my_options.raw = True
    if args.rtfparse:
        my_options.rtf = True
    #else:
    # No arguments were given
    if len(sys.argv) == 1:
        print("No arguments were given. Use -h or --help for more information.")
    
    if my_options.listfolder == True:
        list_outlook_subfolders()

    if my_options.folder != "":
        email_list = read_outlook_subfolder_items( my_options.folder)
    else:
        email_list = get_outlook_emails_from_today(subject = my_options.subject, to=my_options.to, from_ = my_options.sender)

    if my_options.subject != "":
        subject_list = []
        for item in email_list:
            if internal_compare_route( item, "subject", my_options.subject):
                subject_list.append(item)
                #extract_table_ews(get_sublist, create_df_from_list, item)
        email_list = subject_list
    if my_options.sender != "":
        sender_list = []
        for item in email_list:
            if internal_compare_route( item, "SenderName", my_options.sender ):
                sender_list.append(item)
        email_list = sender_list
    my_plug = cls("EWS")
    for item in email_list:
        print_body_data(my_options, item)
        my_plug.load(item[3])
        my_plug.process()
        

                #extract_table_ews(get_sublist, create_df_from_list, item)
        

