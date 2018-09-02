from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools


class EmailFinder(object):
    # If modifying these scopes, delete the file token.json.
    SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'

    def __init__(self):
        store = file.Storage('token.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', self.SCOPES)
            creds = tools.run_flow(flow, store)
        self.service = build('gmail', 'v1', http=creds.authorize(Http()))

    def __filter_on_content(self, content):
        return None

    def find_emails(self, search_str, has_attachment):
        return None

    def find_commute_receipts(self, home_addr, office_addr):
        # from:("uber receipts") has:attachment after:2018-03-31 before:2018-05-01
        # from:("ola") has:attachment after:2018-03-31 before:2018-05-01
        q = 'from:("uber receipts") has:attachment after:2018-03-31 before:2018-05-01'
        uber_mails = self.service.users().messages().list(userId='me', q=q).execute()
        uber_mails = service.users().messages().list(userId='me', q=q).execute()
        sample_mail = service.users().messages().get(userId='me', id='162bf46aa0b0a4e3', format='full').execute()
        self.find_emails("uber", True)
        self.find_emails("ola", True)
        self.__filter_on_content([home_addr, office_addr])
        return None

    def save_commute_receipts(self, dest_dir):
        # self.find_commute_receipts
        # save attachments
        return None