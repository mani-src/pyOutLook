import time
from pyOutLook.lib.check_mail import Check
from pyOutLook.lib.send_mail import Send
from pyOutLook.lib.mail import Mail

check_email = Check()
check_email.set_outlook_folder('Inbox')
# mail = Mail()
mail = check_email.get_latest_mail()
print("Subject: " + mail.subject)
send = Send()
send_mail = Mail(send.get_mail_obj())
send_mail.to = 'ma185307@ncr.com'
send_mail.subject = 'Test1'
send_mail.body = 'This is a test email'
send_mail.attachments = ['C:\\MyProjects\\image001.jpg']
print(send_mail.preview_mail)
# send.send(send_mail)
while check_email.get_latest_mail().subject != 'Test1':
    time.sleep(5)
mail = check_email.get_latest_mail()
send_mail = send.get_forward_mail(mail)
# send_mail.subject = forward_mail.Subject
# for attachment in forward_mail.Attachments:
#    send_mail.mail.Attachments.Add(attachment)
send_mail.to = 'Manikanta, Ambadipudi'
print(send_mail.subject)
send.send(send_mail)
# mail_list = check_email.get_all_mails()
# found_mails = check_email.get_specific_mail(subject='Another test email')
# print(found_mails)

