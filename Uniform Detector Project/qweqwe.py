import imapclient
import email
import smtplib

def send_email(sender_email, receiver_email, subject, body):
    """
    Sends an email using Gmail.

    Args:
        sender_email: The sender's email address.
        receiver_email: The receiver's email address.
        subject: The subject of the email.
        body: The body of the email.
    """

    # # Connect to Gmail IMAP server
    # server = imapclient.IMAPClient('imap.gmail.com', ssl=True)
    # server.login(sender_email, 'eris nyat ehxd gydk')

    # Compose the email message
    msg = email.message.EmailMessage()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.set_content(body)

    # Send the email using SMTP
    smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_server.starttls()
    smtp_server.login(sender_email, 'ewua lhvi hoqq slot')
    smtp_server.send_message(msg)
    smtp_server.quit()

# Example usage
sender_email = 'djtiglao@cca.edu.ph'
receiver_email = 'jerrymarsantos16@gmail.com'
subject = 'Test Email'
body = 'This is a test email sent using Python.'

send_email(sender_email, receiver_email, subject, body)