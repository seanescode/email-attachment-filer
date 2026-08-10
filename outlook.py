import win32com.client
import datetime


def check_email():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Inbox folder
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items

    found = False

    date_today = datetime.date.today()
    date_checking_from = date_today - datetime.timedelta(weeks=4)
    print(date_checking_from)

    nominated_senders = ["xyz123automation@gmail.com", "account-security-noreply@accountprotection.microsoft.com", "j@gmail.com", "k@gmail.com"]

    for email_sender in nominated_senders:
        for email in messages:
            try:
                print(email_sender)
                sender = email.SenderEmailAddress

                if sender.lower() == email_sender.lower():
                    if  email.ReceivedTime.date() > date_checking_from:
                        # if email.Attachments.Count > 0:
                        print("Email with attachment(s) found!")
                        print("Subject:", email.Subject)
                        print("Received:", email.ReceivedTime.date())
                        print("Message:", email.Body[:200])

                        found = True

            except Exception:
                print("Exception! Email not found!")

    if not found:
        print("No emails found from nominated senders")


if __name__ == "__main__":
    check_email()

