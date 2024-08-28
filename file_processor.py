import re
import pandas as pd
from datetime import datetime

def extract_info_from_filename(filename):
    # حذف پسوند فایل
    base_name = filename.split('.')[0]

    # الگوی شناسایی نام فایل: فرض بر این است که ساختار فایل ثابت است
    pattern = r"(\d{1,3})([A-Z]\d)(\d+)"
    match = re.match(pattern, base_name)

    if not match:
        raise ValueError("Filename structure doesn't match the expected pattern.")

    customer_code = match.group(1)
    exit_code = match.group(2)
    barcode = match.group(3)

    # گرفتن تاریخ و زمان کنونی
    current_date = datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M:%S")

    return {
        "customer_code": customer_code,
        "exit_code": exit_code,
        "barcode": barcode,
        "date": current_date,
        "time": current_time
    }

def process_and_save_excel(filepath, output_dir, extracted_info):
    df = pd.read_excel(filepath)

    # جداسازی بر اساس E200 و E280
    e200_df = df[df['Tag ID'].str.startswith('E200')]
    e280_df = df[df['Tag ID'].str.startswith('E280')]

    # اضافه کردن ستون‌های جدید با نام‌های آلمانی
    for data in [e200_df, e280_df]:
        data.insert(0, "Nr", range(1, len(data) + 1))
        data.insert(1, "Kundencode", extracted_info['customer_code'])
        data.insert(2, "Ausgangscode", extracted_info['exit_code'])
        data.insert(3, "Datum", extracted_info['date'])
        data.insert(4, "Uhrzeit", extracted_info['time'])
        data.insert(5, "Barcode", extracted_info['barcode'])

    # ذخیره فایل‌های جداگانه
    e200_filename = f"{extracted_info['customer_code']}_{extracted_info['exit_code']}_E200.xlsx"
    e280_filename = f"{extracted_info['customer_code']}_{extracted_info['exit_code']}_E280.xlsx"

    e200_filepath = f"{output_dir}/{e200_filename}"
    e280_filepath = f"{output_dir}/{e280_filename}"

    e200_df.to_excel(e200_filepath, index=False)
    e280_df.to_excel(e280_filepath, index=False)

    return e200_filepath, e280_filepath
