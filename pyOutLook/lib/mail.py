import os
import datetime
from pyOutLook.lib.outlook import Outlook


class Mail(object):
    def __init__(self, mail=None):
        self.mail = mail
        self.__recipients = []
        if not self.mail:
            raise Exception('Mail object cannot be empty')

    @property
    def subject(self):
        return self.mail.Subject

    @property
    def topic(self):
        return self.mail.ConversationTopic

    @property
    def body(self):
        return self.mail.Body

    @property
    def time(self):
        timestamp = datetime.datetime
        timestamp = self.mail.ReceivedTime
        return timestamp.strftime('%d/%m/%Y %I:%M:%S %p')

    @property
    def to(self):
        return self.mail.To

    @property
    def cc(self):
        return self.mail.CC

    @property
    def sender(self):
        return self.mail.Sender.Name

    @property
    def recipients(self):
        self.__recipients = self.to.split(';')
        self.__recipients.extend(self.cc.split(';'))
        return self.__recipients

    @property
    def has_attachments(self):
        if self.mail.Attachments.Count < 1:
            return False
        else:
            return True

    def download_attachments(self, download_path='C:\\Downloads\\'):
        if self.has_attachments is False:
            return False
        else:
            if not os.path.exists(download_path):
                os.makedirs(download_path)
            attachments = self.mail.Attachments
            for attachment in attachments:
                file = attachment.FileName
                attachment.SaveAsFile("".join((download_path, file)))
            return True

    @property
    def read(self):
        display_string = ''.join(('From: ', self.sender, '\n', 'To: ', self.to, '\n', 'CC: ',
                                  self.cc, 'Subject: ', self.subject, '\n', 'Sent: ',
                                  self.time, '\n\n', '-----------------------------------------------',
                                  self.body, '------------------------------------------------'))
        return display_string
