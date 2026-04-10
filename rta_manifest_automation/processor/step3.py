from datetime import datetime
import re
from .utils import *


def process_step_3(workbook):
    step3_sheet = copy_sheet(workbook, "Step 2", "Step 3")

    remove_empty_row_below_header(step3_sheet)
    format_header(step3_sheet, header_row=1)
    insert_time_column(step3_sheet)
    clear_all_highlighting(step3_sheet)
    highlight_rows(step3_sheet, header_row=1)

    extract_time_from_item(step3_sheet)
    forward_fill_columns(step3_sheet)
    fill_customer_column_by_regid(step3_sheet)
    autofit_columns(step3_sheet)


def insert_time_column(sheet):
    date_col = get_column_index_by_header(sheet, "Date/Time", header_row=1)
    time_col = date_col + 1

    sheet.insert_cols(time_col)
    sheet.cell(row=1, column=date_col).value = "Date"
    sheet.cell(row=1, column=time_col).value = "Time"


def extract_time_from_item(sheet):
    header_row = 1
    time_col = get_column_index_by_header(sheet, "Time", header_row)
    item_col = get_column_index_by_header(sheet, "Item", header_row)
    last_row = sheet.max_row

    time_pattern = re.compile(r"\b(\d{1,2}:\d{2}\s?(?:AM|PM|am|pm))\b")

    for row in range(2, last_row + 1):
        item_cell = sheet.cell(row=row, column=item_col)
        if item_cell.value and isinstance(item_cell.value, str):
            match = time_pattern.search(item_cell.value)
            if match:
                time_str = match.group(1)
                sheet.cell(row=row, column=time_col).value = time_str.strip()
                item_cell.value = item_cell.value.replace(time_str, "").strip()


def forward_fill_columns(sheet):
    header_row = 1
    mech_row = get_mechanical_totals_row(sheet)
    data_end_row = mech_row - 2

    columns_to_fill = ["RegID", "Date", "Time", "Tender"]
    for col_name in columns_to_fill:
        col_index = get_column_index_by_header(sheet, col_name, header_row)
        last_value = None
        for row in range(2, data_end_row + 1):
            cell = sheet.cell(row=row, column=col_index)
            if cell.value in (None, ""):
                if last_value is not None:
                    if col_name.lower() == "date":
                        # Ensure correct date format without time
                        if isinstance(last_value, datetime):
                            cell.value = last_value.date()
                        else:
                            cell.value = last_value
                        cell.number_format = 'm/d/yyyy'
                    else:
                        cell.value = last_value
            else:
                last_value = cell.value


def get_mechanical_totals_row(sheet):
    for row in range(1, sheet.max_row + 1):
        val = sheet.cell(row=row, column=get_column_index_by_header(
            sheet, "Item", 1)).value
        if isinstance(val, str) and "mechanical totals" in val.lower():
            return row
    raise ValidationError("'Mechanical Totals' row not found.")


def remove_empty_row_below_header(sheet):
    row_idx = 2
    if all(sheet.cell(row=row_idx, column=col).value in (None, "") for col in range(1, sheet.max_column + 1)):
        sheet.delete_rows(row_idx)


def fill_customer_column_by_regid(sheet):
    header_row = 1
    mech_row = get_mechanical_totals_row(sheet)
    last_data_row = mech_row - 2

    regid_col = get_column_index_by_header(sheet, "RegID", header_row)
    customer_col = get_column_index_by_header(sheet, "Customer", header_row)

    # Step 1: Sort by RegID (ascending)
    sort_sheet_by_column(sheet, regid_col, header_row, last_data_row)

    # Step 2: Group-wise fill
    current_regid = None
    current_customer = None

    for row in range(2, last_data_row + 1):
        regid = sheet.cell(row=row, column=regid_col).value
        customer_cell = sheet.cell(row=row, column=customer_col)

        if regid != current_regid:
            current_regid = regid
            if customer_cell.value and str(customer_cell.value).strip():
                current_customer = customer_cell.value
            else:
                current_customer = "ZNoFirstName ZNoLastName"
                customer_cell.value = current_customer
        else:
            if customer_cell.value in (None, ""):
                customer_cell.value = current_customer
