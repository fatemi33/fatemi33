import os
import email
import logging
from logging_setup import setup_logging
from email_connection import connect_to_email
from file_processor import extract_info_from_filename, process_and_save_excel
from send_email import send_reports
from log_summary import log_summary


def download_and_process_data():
    mail = connect_to_email()
    if not mail:
        return

    mail.select('inbox')
    status, response = mail.search(None, '(UNSEEN)')
    email_ids = response[0].split()

    if not email_ids:
        logging.info("No Excel files were found.")
        return

    directory = "C:/Users/Hosse/Desktop/RFID"
    if not os.path.exists(directory):
        os.makedirs(directory)

    for e_id in email_ids:
        status, msg_data = mail.fetch(e_id, '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])
        email_subject = msg['subject']
        logging.info(f"Processing email: {email_subject}")

        for part in msg.walk():
            if part.get_content_disposition() == 'attachment':
                filename = part.get_filename()

                if filename and filename.endswith('.xls'):
                    filepath = os.path.join(directory, filename)
                    logging.info(f"Saving file to: {filepath}")
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    logging.info(f"Downloaded file: {filename}")

                    extracted_info = extract_info_from_filename(filename)

                    # پردازش و ذخیره فایل
                    e200_filepath, e280_filepath = process_and_save_excel(filepath, directory, extracted_info)

                    # ارسال فایل‌ها
                    send_reports(e200_filepath)
                    send_reports(e280_filepath)

    mail.logout()


if __name__ == "__main__":
    setup_logging()
    download_and_process_data()
    log_summary()  # ثبت خلاصه پروژه در لاگ
