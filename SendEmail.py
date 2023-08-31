import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from BizProject import file_name

def sendEmail(email_list, file_name):
    # Send CSV File
    port = 465  # For non-SSL
    smtp_server = "mail.autofirstparts.com"
    password = "190Silver01!@"

    # Create a multipart message and set headers
    message = MIMEMultipart()
    sender_email = "it@autofirstparts.com"
    receiver_email = 'andrew@autofirstparts.com'  # ", ".join(email_list)
    body = "Please find the latest inventory update in the attached file."
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Biz Inventory Update"

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Open attachment file in binary mode for reading
    with open(file_name, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_name}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)

    print("Email sent successfully")


# Call the sendEmail function
sendEmail([], file_name)  # Pass the file_name as an argument