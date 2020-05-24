"""
Folder mappings to access any specific folders from the Outlook account
3  Deleted Items
4  Outbox
5  Sent Items
6  Inbox
9  Calendar
10 Contacts
11 Journal
12 Notes
13 Tasks
14 Drafts
"""

import winsound
import win32com.client as win32


class Com(object):
    def __init__(self, application_name, name_space=None):
        self.application_name = application_name
        self.name_space = name_space
        self.__com_obj = None
        self.__com_obj_init()

    @property
    def com_obj(self):
        return self.__com_obj

    def __com_obj_init(self):
        if self.name_space is not None:
            self.__com_obj = win32.Dispatch(self.application_name).GetNamespace(self.name_space)
        else:
            self.__com_obj = win32.Dispatch(self.application_name)

