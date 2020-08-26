from pyOutLook.lib.outlook import Outlook


class Send(Outlook):
    def __init__(self):
        super().__init__()
        super()._init_send_com()
        self.__mail_item = super()._create_mail_item()

    def get_mail_obj(self):
        return self.__mail_item

    def send(self, mail):
        if str(type(mail)) != 'pyOutLook.lib.mail.Mail':
            return
        if not mail.verify:
            raise Exception('To and Subject must be specified')
        self.__mail_item.Send()

    def get_forward_mail(self, recv_mail, keep_attachments=True):
        send_mail = recv_mail.mail.Forward()
        if not keep_attachments:
            for attachment in recv_mail.attachments:
                attachment.Remove(1)
        return send_mail


