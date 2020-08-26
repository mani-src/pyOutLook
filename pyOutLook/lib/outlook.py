from pyOutLook.lib.com import Com


class Outlook(object):
    def __init__(self):
        self._outlook = None
        self._folder = None
        self._mail_item = None

    def _init_com_obj(self):
        com_obj = Com('Outlook.Application', 'MAPI')
        self._outlook = com_obj.com_obj

    def _init_send_com(self):
        com_obj = Com('Outlook.Application')
        self._outlook = com_obj.com_obj

    def _select_folder(self, folder_type):
        if not type(folder_type):
            raise Exception('folder_type is of incorrect format')
        folder_num = 0
        if folder_type.upper() == 'INBOX':
            folder_num = 6
        elif folder_type.upper() == 'SENT ITEMS':
            folder_num = 5
        elif folder_type.upper() == 'DELETED ITEMS':
            folder_num = 3
        elif folder_type.upper() == 'OUTBOX':
            folder_num = 4
        elif folder_type.upper() == 'DRAFTS':
            folder_num = 14
        elif folder_type.upper() == 'CALENDAR':
            folder_num = 9
        elif folder_type.upper() == 'CONTACTS':
            folder_num = 10
        elif folder_type.upper() == 'TASKS':
            folder_num = 13
        elif folder_type.upper() == 'NOTES':
            folder_num = 12
        elif folder_type.upper() == 'JOURNAL':
            folder_num = 11
        else:
            raise Exception('Unsupported outlook folder type requested\n '
                            'Please check the docstring to get the supported outlook folder types')
        if folder_num != 0:
            self._folder = self._outlook.GetDefaultFolder(folder_num)

    def _get_mails(self):
        if not self._folder:
            raise Exception('Folder not selected')
        mail_list = self._folder.Items
        return mail_list

    def _create_mail_item(self):
        self._mail_item = self._outlook.CreateItem(0)    # 0 is for mail item. For the other types, read the references
        return self._mail_item
