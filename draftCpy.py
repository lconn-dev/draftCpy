# DraftCpy.py - Export Drafts and Import drafts across G-Mail accounts
import httplib2
import os
import random
import sys
import time
import string
import random
import base64
from pathlib import Path

#from email.mime.audio import MIMEAudio
#from email.mime.base import MIMEBase
#from email.mime.image import MIMEImage
#from email.mime.multipart import MIMEMultipart
#from email.mime.text import MIMEText
#import mimetypes

from termcolor import colored, cprint


from tqdm import tqdm #progress bar lib

import pickle # for credential caching for OAUTH
import os.path

# Google Libs:
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from apiclient import errors

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

def random_generator(size=6, chars=string.ascii_uppercase + string.digits):
    return ''.join(random.choice(chars) for x in range(size))

def getTargetLabelId(service, target): # find the ID of our target label
  try:
    result = None
    response = service.users().labels().list(userId="me").execute()
    labels = response['labels']
    for label in labels:
      if label["name"] == target:
        result = label["id"]
        break
    return result
  except errors.HttpError as error:
    print('An error occurred pulling labels:', error)


def ListDrafts(service, user_id):
  try:
    skip = True
    #ids = []
    maxRes = 25
    response = service.users().drafts().list(userId=user_id, maxResults=maxRes).execute()
    drafts = response['drafts']
    draftsProcd = 0
    #totalDrafts = len(drafts)
    iters = 1
    unreadable = 0
    saved = 0
    estimate = response['resultSizeEstimate']
    
    fileName = str(Path.home()) + "/Documents/DraftCpy Exports/archive_" + random_generator() + ".drafts"
    print("Preparing export file:", fileName)
    Path(fileName).touch()

    draftLabel = "<DRAFTEXPORT>"
    targetLabelId = getTargetLabelId(service, draftLabel)
    if(targetLabelId == None):
        print("Failure detecting label ID's. Make sure the drafts are in the <DRAFTEXPORT> label.")
        exit()

    #while os.path.exists(fileName):
        #fileName = "archive_" + random_generator() + ".drafts"

    running = True
    with tqdm(total = estimate + 1) as pbar:
        while running:
            iters += 1
            if skip: # skip the first result
                skip = False
                continue
            else:
                if "nextPageToken" in response:
                    npt = response["nextPageToken"]
                    response = service.users().drafts().list(userId=user_id, pageToken=npt, maxResults=maxRes).execute()
                else:
                    response = service.users().drafts().list(userId=user_id, maxResults=maxRes).execute() 
                    running = False

                drafts = response['drafts']
                time.sleep(0.05) # give the servers sime time to breathe
            
            if running:
                estimate = draftsProcd + response['resultSizeEstimate']
            else:
                estimate = draftsProcd + len(response['drafts'])
            
            pbar.total = estimate
            pbar.refresh()
            # Collect draft IDs from this pass and display progress information
            #print(draftsProcd)
            for draft in drafts:
                draftsProcd += 1
                #ids += [draft["id"]]
                pbar.update(1)
                pbar.set_description(" > Processing ID# %s" % draft["id"])

                # Pull draft
                try:
                    # pull label metadata
                    data = service.users().drafts().get(userId=user_id, id=draft["id"], format="metadata").execute()
                    try:
                        if targetLabelId in data['message']['labelIds']: # check label
                            # pull full message in RAW format
                            data = service.users().drafts().get(userId=user_id, id=draft["id"], format="raw").execute()
                            # open file for export and write picked raw data
                            with open(fileName, 'ab') as exportFile:
                                pickle.dump(data['message']['raw'], exportFile)
                                saved += 1
                        else:
                            unreadable += 1
                    except:
                        unreadable += 1

                except errors.HttpError as error:
                    print ("An error occurred while pulling a draft", error)

                time.sleep(0.05) # give the servers sime time to breathe
    
    pbar.close()
    
    print(" > Processed", draftsProcd, "drafts.", "Consumed", iters, "API Calls.")
    print(" > Unreadable Drafts:", unreadable)
    print(" > Exported Drafts:", saved)
    print( " ")
    print(" > Export Completed. Saved to file:", fileName, "\n")
    input(" >>> PRESS ENTER TO EXIT <<< ")
    #return drafts

  except errors.HttpError as err:
    print( '[ERROR] An error occurred: ', err)

def export(service):
    ListDrafts(service, "me")

def authorize(useCache):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.bin') and not useCache:
        with open('token.bin', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.bin', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)


def pickleLoader(pklFile):
    try:
        while True:
            yield pickle.load(pklFile)
    except EOFError:
        pass


def CreateDraft(service, user_id, message_body):
  """Create and insert a draft email. Print the returned draft's message and id.
  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message_body: The body of the email message, including headers.

  Returns:
    Draft object, including draft id and message meta data.
  """
  try:
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()

    #print 'Draft id: %s\nDraft message: %s' % (draft['id'], draft['message'])
    print("Creaded draft with Id# :", draft['id'])
    return draft
  except errors.HttpError as error:
    #print 'An error occurred: %s' % error
    print("Error adding draft: ", error)
    return None

def MakeLabel(label_name, mlv='show', llv='labelShow'):
  label = {'messageListVisibility': mlv,
           'name': label_name,
           'labelListVisibility': llv}
  return label

def ModifyMessage(service, user_id, msg_id, msg_labels):
    try:
            
        message = service.users().messages().modify(
            userId=user_id,
            id=msg_id,
            body=msg_labels
        ).execute()

        #label_ids = message['labelIds']
        return message

    except errors.HttpError as error:
        print("An Error occured while labeling message.", error)
        input("PRESS ENTER TO EXIT")
        exit()

def importDrafts(fpath, service):
    print("Loading Archive...")
    lbl = input("Enter Label for imported drafts: ")
    try:
        label = service.users().labels().create(
            userId="me",
            body=MakeLabel(lbl)
        ).execute()

        print("Created Label: ", label['id'])
        
        
        amnt = 0
        with open(fpath, 'rb') as f:
            for event in pickleLoader(f):
                # Create The draft
                cdraft = CreateDraft(service, "me", {'raw': event})
                labels = {'addLabelIds': [label['id']]}
                ModifyMessage(service, "me", cdraft['message']['id'], labels)
                amnt += 1
            print(" > Imported", amnt, "drafts.")
            input(" >>> PRESS ENTER TO EXIT <<< ")

    except errors.HttpError as error:
        #print 'An error occurred: %s' % error
        print("An error occurred while creating the label\"", lbl, "\"")
        print(error)
        input("PRESS ENTER TO EXIT")
        exit()

def main():
    print("\n##########################################")
    print( "             DraftCpy v1.12b"               )
    print( "          Luke Connor (c) 2020"             )
    print("##########################################\n")
    print(" Enter \'QUIT\' at any prompt to exit the program\n")
    print(" > Connecting...\n" )

    inv = ""
    useCachedCreds = True
    
    if os.path.exists('token.bin'):
        while(inv != "Y" and inv != "N" and inv != "QUIT"):
            inv = input("Invalidate cached credentials? (Y/N): ")

            if(inv == "N"):
                useCachedCreds = False

            if(inv == "QUIT"): quit()

    emSvc = authorize(useCachedCreds)
    print(" > Successfully Connected & Authenticated.")

    ieOpt = ""
    while ieOpt != "I" and ieOpt != "E" and ieOpt != "QUIT":
        ieOpt = input("- (I)mport or (E)xport? (Type \'I\' or \'E\' and press \'Enter\'): ")

    if (ieOpt == "QUIT"):
        print(" > Goodbye!")
    if (ieOpt == "E"):
        print(" > Starting Export...")
        export(emSvc)
    if(ieOpt == "I"):
        print(" > Starting Import...")
        invalid = True
        while(invalid):
            fpath = input(" Enter File Path: ")
            if os.path.exists(fpath):
                importDrafts(fpath, emSvc)
                invalid = False
            else:
                print("ERROR! File does not exist. Check your file path and try again")


if __name__ == "__main__":
    main()