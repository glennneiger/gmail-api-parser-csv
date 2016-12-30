# Based on and using code from examples at: https://developers.google.com/gmail/api/
# Google API test harness: https://developers.google.com/apis-explorer/?hl=en_GB#p/gmail/v1/

from __future__ import print_function
import email.mime.text
import httplib2
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from googleapiclient.discovery import build
from apiclient import errors
import base64
import sys, os, re, StringIO
import email, mimetypes
from alchemyapi import AlchemyAPI
import csv
import sys #maybe used in the future


# --------------------------------------------------------------------------------------
# This is the name of the secret file you download from https://console.developers.google.com/iam-admin/projects
# Give it a name that is unique to this project
CLIENT_SECRET_FILE = 'python_gmail_api_client_secret.json'
# This is the file that will be created in ~/.credentials holding your credentials. It will be created automatically
# the first time you authenticate and will mean you don't have to re-authenticate each time you connect to the API.
# Give it a name that is unique to this project
CREDENTIAL_FILE = 'python_gmail_api_credentials.json'

APPLICATION_NAME = 'python-gmail-api'
# Set to True if you want to authenticate manually by visiting a given URL and supplying the returned code
# instead of being redirected to a browser. Useful if you're working on a server with no browser.
# Set to False if you want to authenticate via browser redirect.
MANUAL_AUTH = True
# --------------------------------------------------------------------------------------

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    if MANUAL_AUTH:
        flags.noauth_local_webserver=True
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials at ~/.credentials/gmail-python-quickstart.json
SCOPES = ['https://mail.google.com/',
          'https://www.googleapis.com/auth/gmail.modify',
          'https://www.googleapis.com/auth/gmail.readonly']

class PythonGmailAPI:
    def __init__(self):
        pass

    def gmail_list(self):
        credentials = self.__get_credentials()
        service = self.__build_service(credentials)
        emailIds = self.ListMessagesMatchingQuery(service, 'me', '')
        fullEmailList = []
        for eId in emailIds:
            fullEmailList.append(self.GetMessage(service, 'me', eId['id']))

        subjectField = ["" for x in range(len(fullEmailList))]
        fromField = [["", ""] for x in range(len(fullEmailList))] #name and email in arr
        toField = [["", ""] for x in range(len(fullEmailList))] #name and email in arr
        filenameField = ["" for x in range(len(fullEmailList))] #attachment? will have to add this later
        is_bodyField = ["" for x in range(len(fullEmailList))]
        typeField = ["" for x in range(len(fullEmailList))] #same as above?
        charsetField = ["" for x in range(len(fullEmailList))] #ascii etc
        descField = ["" for x in range(len(fullEmailList))] #what is this?
        sizeField = ["" for x in range(len(fullEmailList))] #count characters
        bodyField = ["" for x in range(len(fullEmailList))]
        sentimentTypeField = ["" for x in range(len(fullEmailList))]
        sentimentScoreField = ["" for x in range(len(fullEmailList))]

        ctr = 0 #keep track of which email we're at

        # write to file specified in arguments list
        g = open("1.csv", "w")
        w = csv.writer(g, lineterminator='\n')
        w.writerow(('subject', 'fromName','fromMail','toName','toMail','filename','is_body','type','charset','desc',
                    'size','body','sentimentType','sentimentScore'))

        for x in fullEmailList:
            for h in x['payload']['headers']:
                if h['name'] == 'Subject':
                    subjectField[ctr] = h['value']
            for f in x['payload']['headers']:
                if f['name'] == 'From':
                    if (len(f['value'].split(" ")) == 1): fromField[ctr][1] = \
                        f['value'].split(" ")[0].replace("<","").replace(">", "") #if only email (no name)
                    else:
                        fromName = str(f['value']).split(" ")[0].strip() + " " + str(f['value']).split(" ")[1].strip()
                        if (len(str(f['value']).split(" ")) > 2): fromMail = \
                            str(f['value']).split(" ")[2].replace("<","").replace(">", "")
                        fromField[ctr][0] = fromName
                        if (fromMail) : fromField[ctr][1] = fromMail
            for t in x['payload']['headers']:
                if t['name'] == 'To':
                    if (len(t['value'].split(" ")) == 1): toField[ctr][1] = \
                        t['value'].split(" ")[0].replace("\<", "").replace(">","") #if only email (no name)
                    else:
                        toName = str(t['value']).split(" ")[0].strip() + " " + str(t['value']).split(" ")[1].strip()
                        if (len(str(t['value']).split(" ")) > 2): toMail = str(t['value']).split(" ")[2].replace("<", "").replace(">","")
                        toField[ctr][0] = toName
                        if (toMail): toField[ctr][1] = toMail

            filenameField[ctr] = "None" #by default, no attachment
            if ('parts' in x['payload']) :
                for p in x['payload']['parts']:
                    if p['filename'] != "": filenameField[ctr] = p['filename'] #add attachment name if applicable


            if 'parts' in x['payload']: is_bodyField[ctr] = x['payload']['parts'][0]['mimeType']
            else: is_bodyField[ctr] = "text/plain"

            typeField[ctr] = is_bodyField[ctr]  # same thing as is_body field anyway?
            charsetField[ctr] = 'UTF-8'  # hardcoded always UTF-8
            descField[ctr] = 'None'  # there is no description field on Gmail emails

            if ('parts' in x['payload']): sizeField[ctr] = x['payload']['parts'][1]['body']['size']
            else: sizeField[ctr] = x['payload']['body']['size']

            if ('parts' in x['payload'] and 'data' in x['payload']['parts'][0]['body']): bodyField[ctr] = \
                base64.b64decode(x['payload']['parts'][0]['body']['data']) #base64.base64decode
            elif ('parts' in x['payload']): bodyField[ctr] = \
                base64.b64decode(x['payload']['parts'][0]['parts'][0]['body']['data'])
            else: bodyField[ctr] = base64.b64decode(x['payload']['body']['data'])

            #get sentiment for each email
            alchemyapi = AlchemyAPI()
            response = alchemyapi.sentiment('text', bodyField[ctr])
            if response['status'] == 'OK':
                sentimentTypeField[ctr] = response['docSentiment']['type']
                if 'score' in response['docSentiment']:
                    sentimentScoreField[ctr] = response['docSentiment']['score']
                else: sentimentScoreField[ctr] = ""
            else:
                print('Error in sentiment analysis call: ', response['statusInfo'])

            w.writerow(("'" + str(subjectField[ctr]) + "'",
                        fromField[ctr][0],
                        fromField[ctr][1],
                        toField[ctr][0],
                        toField[ctr][1],
                        filenameField[ctr],
                        is_bodyField[ctr],
                        typeField[ctr],
                        charsetField[ctr],
                        descField[ctr],
                        sizeField[ctr],
                        "'" + bodyField[ctr] + "'",
                        sentimentTypeField[ctr],
                        sentimentScoreField[ctr]))

            #DELETE PRINT STATEMENT ONCE CSV GENERATOR IS WORKING
            print("\n\nSubject: " + subjectField[ctr])
            print("From: " + str(fromField[ctr]))
            print("To: " + str(toField[ctr]))
            print("filename: " + filenameField[ctr])
            print("is_body: " + is_bodyField[ctr])
            print("type: " + typeField[ctr])
            print("charset: " + charsetField[ctr])
            print("desc: " + descField[ctr])
            print("size: " + str(sizeField[ctr]))
            print("Body: \n" + bodyField[ctr])
            print("\nSentiment Analysis:")
            print("type: " + sentimentTypeField[ctr])
            print("score; " + sentimentScoreField[ctr])

            ctr = ctr + 1 #proceed to next email



    #from Gmail API documentation - get a list of every email in inbox
    def ListMessagesMatchingQuery(self, service, user_id, query=''):
        """List all Messages of the user's mailbox matching the query.

        Args:
            service: Authorized Gmail API service instance.
            user_id: User's email address. The special value "me"
            can be used to indicate the authenticated user.
            query: String used to filter messages returned.
            Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

        Returns:
            List of Messages that match the criteria of the query. Note that the
            returned list contains Message IDs, you must use get with the
            appropriate ID to get the details of a Message.
        """
        try:
            response = service.users().messages().list(userId=user_id,
                                                    q=query).execute()
            messages = []
            if 'messages' in response:
                messages.extend(response['messages'])

            while 'nextPageToken' in response:
                page_token = response['nextPageToken']
                response = service.users().messages().list(userId=user_id, q=query,
                                                    pageToken=page_token).execute()
                messages.extend(response['messages'])
            return messages
        except errors.HttpError, error:
            print('An error occurred: %s' % error)

    # retrieve message from Gmail (given the id from the listMessages method)
    def GetMessage(self, service, user_id, msg_id):
        """Get a Message with given ID.

        Args:
          service: Authorized Gmail API service instance.
          user_id: User's email address. The special value "me"
          can be used to indicate the authenticated user.
          msg_id: The ID of the Message required.

        Returns:
          A Message.
        """
        try:
            message = service.users().messages().get(userId=user_id, id=msg_id).execute()

            #print('Message snippet: %s' % message['snippet'])

            return message
        except errors.HttpError, error:
            print('An error occurred: %s' % error)


    #isnt used for now, but might need to be for MIME (rich text?) messages
    def GetMimeMessage(self, service, user_id, msg_id):
        """Get a Message and use it to create a MIME Message.

        Args:
          service: Authorized Gmail API service instance.
          user_id: User's email address. The special value "me"
          can be used to indicate the authenticated user.
          msg_id: The ID of the Message required.

        Returns:
          A MIME Message, consisting of data from Message.
        """
        try:
            message = service.users().messages().get(userId=user_id, id=msg_id,
                                                     format='raw').execute()

            msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))

            mime_msg = email.message_from_string(msg_str)

            return mime_msg
        except errors.HttpError, error:
            print('An error occurred: %s' % error)

    # def gmail_send(self, sender_address, to_address, subject, body):
    #     print('Sending message, please wait...')
    #     message = self.__create_message(sender_address, to_address, subject, body)
    #     credentials = self.__get_credentials()
    #     service = self.__build_service(credentials)
    #     raw = message['raw']
    #     raw_decoded = raw.decode("utf-8")
    #     message = {'raw': raw_decoded}
    #     message_id = self.__send_message(service, 'me', message)
    #     print('Message sent. Message ID: ' + message_id)


    def __get_credentials(self):
        """Gets valid user credentials from storage.
        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.
        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir, CREDENTIAL_FILE)
        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            credentials = tools.run_flow(flow, store, flags)
            print('Storing credentials to ' + credential_path)
        return credentials

    def __create_message(self, sender, to, subject, message_text):
      """Create a message for an email.
      Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.
      Returns:
        An object containing a base64url encoded email object.
      """
      message = email.mime.text.MIMEText(message_text, 'plain', 'utf-8')
      message['to'] = to
      message['from'] = sender
      message['subject'] = subject
      encoded_message = {'raw': base64.urlsafe_b64encode(message.as_bytes())}
      return encoded_message


    def __send_message(self, service, user_id, message):
      """Send an email message.
      Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message: Message to be sent.
      Returns:
        Sent Message ID.
      """
      message = (service.users().messages().send(userId=user_id, body=message)
                .execute())
      return message['id']

    def __build_service(self, credentials):
        """Build a Gmail service object.
        Args:
            credentials: OAuth 2.0 credentials.
        Returns:
            Gmail service object.
        """
        http = httplib2.Http()
        http = credentials.authorize(http)
        return build('gmail', 'v1', http=http)


def main():
    PythonGmailAPI().gmail_list()

if __name__ == '__main__':
    main()

