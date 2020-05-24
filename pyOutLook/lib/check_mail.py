import datetime
from pyOutLook.lib.outlook import Outlook
from pyOutLook.lib.mail import Mail


class Check(Outlook):
    def __init__(self):
        super().__init__()

    def set_outlook_folder(self, folder):
        super()._select_folder(folder_type=folder)

    def get_latest_mail(self, folder=''):
        if not folder:
            if not self._folder:
                raise Exception('Folder must be specified')
        else:
            super()._select_folder(folder_type=folder)
        all_mails = super()._get_mails()
        return Mail(all_mails.GetLast())

    def get_oldest_mail(self, folder=''):
        if not folder:
            if not self._folder:
                raise Exception('Folder must be specified')
        else:
            super()._select_folder(folder_type=folder)
        all_mails = super()._get_mails()
        return Mail(all_mails.GetFirst())

    def get_all_mails(self, folder=''):
        if not folder:
            if not self._folder:
                raise Exception('Folder must be specified')
        else:
            super()._select_folder(folder_type=folder)
        all_mails = super()._get_mails()
        all_mail_objs = [Mail(mail) for mail in all_mails]
        return all_mail_objs

    def get_specific_mail(self, folder='', **kwargs):
        if not folder:
            if not self._folder:
                raise Exception('Folder must be specified')
        else:
            super()._select_folder(folder_type=folder)
        all_mails = super()._get_mails()
        all_mail_objs = self.get_all_mails()
        if len(kwargs) < 1:
            raise Exception('kwargs must be supplied')
        found_mail = []
        found = False
        for mail in all_mail_objs:
            for key in kwargs.keys():
                if key:
                    if not hasattr(mail, key):
                        raise Exception(''.join(('Unsupported kwarg passed: ', key)))
                    if getattr(mail, key) != kwargs[key]:
                        found = False
                        break
                    else:
                        found = True
                        continue
            if not found:
                continue
            else:
                found_mail.append(mail)
        return found_mail