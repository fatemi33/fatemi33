import logging

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s:%(message)s',
        handlers=[
            logging.FileHandler("rfid_processor.log", encoding='utf-8'),  # تنظیم انکودینگ به UTF-8
            logging.StreamHandler()
        ]
    )
