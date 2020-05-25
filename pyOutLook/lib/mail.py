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

    @subject.setter
    def subject(self, subject):
        self.mail.Subject = subject

    @property
    def topic(self):
        return self.mail.ConversationTopic

    @property
    def body(self):
        return self.mail.Body

    @body.setter
    def body(self, body):
        self.mail.Body = body

    @property
    def time(self):
        timestamp = datetime.datetime
        timestamp = self.mail.ReceivedTime
        return timestamp.strftime('%d/%m/%Y %I:%M:%S %p')

    @property
    def to(self):
        return self.mail.To

    @to.setter
    def to(self, to):
        self.mail.To = to

    @property
    def cc(self):
        return self.mail.CC

    @cc.setter
    def cc(self, cc):
        self.mail.CC = cc

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

    @property
    def attachments(self):
        attachments = []
        for attachment in self.mail.Attachments:
            attachments.append(attachment)
        return attachments

    @attachments.setter
    def attachments(self, attachments):
        if type(attachments) is not list:
            raise Exception('Attachments must be added in a list')
        for attachment in attachments:
            if not os.path.exists(attachment):
                raise Exception('Attachment path incorrect or not found')
            self.mail.Attachments.Add(attachment)

    def download_attachments(self, download_path='C:\\Downloads\\'):
        if self.has_attachments is False:
            return False
        else:
            if not os.path.exists(download_path):
                os.makedirs(download_path)
            attachments = self.mail.Attachments
            for attachment in attachments:
                file = attachment.FileName
                attachment.SaveAsFile(''.join((download_path, file)))
            return True

    @property
    def read(self):
        display_string = ''.join(('From: ', self.sender, '\n', 'To: ', self.to, '\n', 'CC: ',
                                  self.cc, 'Subject: ', self.subject, '\n', 'Sent: ',
                                  self.time, '\n\n', '-----------------------------------------------',
                                  self.body, '------------------------------------------------'))
        return display_string

    @property
    def preview_mail(self):
        display_string = ''.join(('To: ', self.to, '\n', 'CC: ',
                                  self.cc, '\n', 'Subject: ', self.subject, '\n', 'Attachments: ',
                                  str([attachment.FileName for attachment in self.attachments]), '\n\n',
                                  '-----------------------------------------------', '\n\n', self.body, '\n\n',
                                  '-----------------------------------------------'))
        return display_string

    @property
    def verify(self):
        if not self.mail.To or not self.mail.Subject:
            return False
        else:
            return True
