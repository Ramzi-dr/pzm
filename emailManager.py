

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from payloadCollection import PayloadCollection


def send_email(subject="", message="", attachment_path=""):
    # Email configuration
    from_email = PayloadCollection.email_user
    password = PayloadCollection.email_pass
    username = PayloadCollection.email_user
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    sven = "sb@einbruchschutz.ch"
    to_pzm_email = [
        "ramzi.d@outlook.com",
        "rdr@einbruchschutz.ch",
        "ramzi.techdesign@gmail.com",
        "pcmedianetwork@gmail.com"
    ]
    to_bst_email = ["ramzi.d@outlook.com", "rdr@einbruchschutz.ch",sven]
    to_email = to_bst_email if subject == "Error in Pzm EventServer" else to_pzm_email
    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = ", ".join(to_email)
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))
    if attachment_path:
        
        with open(attachment_path, "rb") as attachment:
            # Extracting the filename from the path
            attachment_filename = attachment_path.split("/")[-1]
            part = MIMEApplication(attachment.read(), Name=attachment_filename)
            part.add_header("Content-Disposition", f'attachment; filename="{attachment_filename}"')
            msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Explicitly handle the TLS connection
        server.login(username, password)

        for recipient in to_email:
            server.sendmail(from_email, recipient.strip(), msg.as_string())

        server.quit()

       # print("Email sent successfully.")
    except Exception as e:
        print("Error sending email:", str(e))
