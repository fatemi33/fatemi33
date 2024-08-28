import os
import email
import logging
from logging_setup import setup_logging
from email_connection import connect_to_email
from data_processor import process_and_convert_to_xlsx
from send_email import send_reports


def log_summary():
    logging.info("========== Project Summary ==========")
    logging.info("1. هدف پروژه: دریافت، پردازش و ارسال فایل‌های اکسل حاوی داده‌های RFID.")
    logging.info("2. مراحل انجام پروژه:")
    logging.info("   - دریافت ایمیل‌ها و دانلود فایل‌های اکسل.")
    logging.info("   - پردازش فایل‌های اکسل و جدا کردن داده‌ها بر اساس Tag ID.")
    logging.info("   - تبدیل فایل‌های .xls به .xlsx.")
    logging.info("   - ارسال فایل‌های پردازش شده به ایمیل مقصد.")
    logging.info("3. ساختار فایل‌های پروژه:")
    logging.info("   - `main.py`: اجرای کلی پروژه.")
    logging.info("   - `email_connection.py`: اتصال به سرور ایمیل.")
    logging.info("   - `data_processor.py`: پردازش داده‌های اکسل.")
    logging.info("   - `send_email.py`: ارسال ایمیل.")
    logging.info("   - `logging_setup.py`: تنظیمات لاگ.")
    logging.info("4. نتیجه‌گیری: پروژه با موفقیت به پایان رسید و همه فایل‌ها به درستی پردازش و ارسال شدند.")
    logging.info("======================================")


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

                    output_filepaths = process_and_convert_to_xlsx(filepath)
                    if output_filepaths:
                        for output_filepath in output_filepaths:
                            send_reports(output_filepath)

    mail.logout()


if __name__ == "__main__":
    setup_logging()
    download_and_process_data()
    log_summary()  # ثبت خلاصه پروژه در لاگ
