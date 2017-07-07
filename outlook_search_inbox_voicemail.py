import win32com.client
import re
import pyperclip
__version__ = '1.2'


class VoiceMailOutlook:
    """
    Will traverse a shared mailbox ("Voicemail") for undelievered emails, extract these email addresses and
    copy to the clipboard
    """
    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.inbox = None
        self.email_pattern = r"[a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*@(?:[a-z0-9]" \
                             r"(?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])"
        self.text_description_pattern = "The email address that you entered couldn't be found"

    def process_emails(self):
        print('Running: ' + __version__)
        mapi = self.outlook.GetNamespace('MAPI')
        # inbox = mapi.GetDefaultFolder(6)  # 6=olFolderInbox=my own inbox
        recipient = mapi.CreateRecipient('Voicemail')  # RM1048='Alias' of IT Third Party Response
        recipient.Resolve()
        if recipient.Resolved:
            print('Mailbox was successfully resolved.')
            self.inbox = mapi.GetSharedDefaultFolder(recipient, 6)
            messages = self.inbox.Items
            result = []
            print('Found ' + str(len(messages)) + ' messages to process:\nProcessing', end='')
            for no, message in enumerate(messages, start=1):
                if no % 10 == 0:
                    print('Processing number: ' + str(no))
                # print(message.Subject)
                if re.search('Undeliverable', message.Subject) \
                        and re.search(self.text_description_pattern, message.Body)\
                        and re.search(self.email_pattern, message.Body.lower()):
                            email = re.search(self.email_pattern, message.Body.lower()).group()
                            # print(email)
                            result.append(email)
            return result
        else:
            print('Mailbox was NOT successfully resolved')

v = VoiceMailOutlook()
email_addresses = v.process_emails()
email_addresses = set(email_addresses)
print('\n'.join(email_addresses))
pyperclip.copy('\n'.join(email_addresses))
print(str(len(email_addresses)) + ' email addresses have been copied to the clipboard.')