from pyOutLook.lib.check_mail import Check
from pyOutLook.lib.mail import Mail

check_email = Check()
check_email.set_outlook_folder('Inbox')
# mail = Mail()
mail = check_email.get_latest_mail()
print("Subject: " + mail.subject)
# mail_list = check_email.get_all_mails()
found_mails = check_email.get_specific_mail(subject='Another test email')
print(found_mails)
