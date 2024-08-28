import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import logging
from config import EMAIL_ACCOUNT, RECEIVER_EMAIL, SMTP_SERVER, SMTP_PORT, PASSWORD

def send_reports(filepath):
    try:
        subject = "Processed and Converted Excel File"
        body = "Please find the attached processed and converted Excel file."

        msg = MIMEMultipart()
        msg['From'] = EMAIL_ACCOUNT
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'octet-stream')
        with open(filepath, 'rb') as f:
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(filepath)}')
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ACCOUNT, PASSWORD)
            server.sendmail(EMAIL_ACCOUNT, RECEIVER_EMAIL, msg.as_string())
            logging.info("Email sent successfully.")
    except Exception as e:
        logging.error(f"Failed to send email: {str(e)}")
