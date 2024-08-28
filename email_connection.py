import imaplib
import logging
from config import IMAP_SERVER, EMAIL_ACCOUNT, PASSWORD

def connect_to_email():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ACCOUNT, PASSWORD)
        logging.info("Connected to email server.")
        return mail
    except Exception as e:
        logging.error(f"Failed to connect to email server: {str(e)}")
        return None
