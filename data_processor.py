import pandas as pd
import logging


def process_and_convert_to_xlsx(filepath):
    try:
        # خواندن فایل اکسل اصلی
        df = pd.read_excel(filepath)

        # جدا کردن داده‌های E200 و E280
        df_e200 = df[df['Tag ID'].str.startswith('E200')]
        df_e280 = df[df['Tag ID'].str.startswith('E280')]

        # ذخیره فایل‌های جداگانه به فرمت .xlsx
        output_filepath_e200 = filepath.replace('.xls', '_E200.xlsx')
        output_filepath_e280 = filepath.replace('.xls', '_E280.xlsx')

        df_e200.to_excel(output_filepath_e200, index=False)
        df_e280.to_excel(output_filepath_e280, index=False)

        logging.info(f"File converted and saved as: {output_filepath_e200} and {output_filepath_e280}")

        # بازگشت مسیر فایل‌های جدید برای ارسال از طریق ایمیل
        return [output_filepath_e200, output_filepath_e280]

    except Exception as e:
        logging.error(f"Error processing file {filepath}: {str(e)}")
        return None
