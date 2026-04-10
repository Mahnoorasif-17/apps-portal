import openpyxl
from datetime import datetime
import re
from .utils import *


def process_step_1(filepath):
    wb = openpyxl.load_workbook(filepath)
    original_sheet = wb.sheetnames[0]
    step1_sheet = copy_sheet(wb, original_sheet, "Step 1")

    # if not verify_date_range(step1_sheet):
    #     raise ValidationError("Date range in A2 is not for the current month")

    format_header(step1_sheet)
    highlight_rows(step1_sheet)
    return wb


def verify_date_range(sheet):
    date_text = sheet['A2'].value
    date_text = date_text.strip()
    match = re.match(
        r"(\d{2}/\d{2}/\d{2}) to (\d{2}/\d{2}/\d{2})", str(date_text))
    if not match:
        raise ValidationError(f"Invalid date format in cell A2: {date_text}")
    start_str, _ = match.groups()
    start_date = datetime.strptime(start_str, "%m/%d/%y")
    now = datetime.now()
    return start_date.month == now.month and start_date.year == now.year


def delete_above_header(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=20):
        if row[0].value == "RegID":
            header_row = row[0].row
            for _ in range(header_row - 1):
                sheet.delete_rows(1)
            return
    raise ValidationError("Header row not found")
