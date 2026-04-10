from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

import os


ORANGE_FILL = PatternFill(start_color="FFFFD580",
                          end_color="FFFFD580", fill_type="solid")
GRAY_FILL = PatternFill(start_color='DDDDDD',
                        end_color='DDDDDD', fill_type='solid')


class ValidationError(Exception):
    def __init__(self, message, workbook=None):
        super().__init__(message)
        self.workbook = workbook


def generate_new_filename(filepath):
    base, ext = os.path.splitext(filepath)
    return f"{base} - Processed{ext}"


def copy_sheet(workbook, source_name, target_name):
    source = workbook[source_name]
    target = workbook.copy_worksheet(source)
    target.title = target_name
    return target


def highlight_row(sheet, row, max_col, fill):
    for col in range(1, max_col + 1):
        sheet.cell(row=row, column=col).fill = fill


def get_footer_row(sheet):
    for row in reversed(range(1, sheet.max_row + 1)):
        if any(sheet.cell(row=row, column=col).value for col in range(1, sheet.max_column + 1)):
            return row
    raise ValidationError("Footer row not found")


def format_header(sheet, header_row=7):
    if sheet[f"A{header_row}"].value != "RegID":
        raise ValidationError(f"Expected 'RegID' in cell A{header_row}")

    sheet.freeze_panes = f"A{header_row + 1}"

    max_col = sheet.max_column
    max_row = sheet.max_row
    for col in range(1, max_col + 1):
        cell = sheet.cell(row=header_row, column=col)
        cell.alignment = Alignment(horizontal='left')
    last_col_letter = get_column_letter(max_col)
    sheet.auto_filter.ref = f"A{header_row}:{last_col_letter}{max_row}"


def highlight_rows(sheet, header_row=7):
    max_row = sheet.max_row
    max_col = sheet.max_column
    for col in range(1, max_col + 1):
        sheet.cell(row=header_row, column=col).fill = GRAY_FILL
    for row in reversed(range(1, max_row + 1)):
        if any(sheet.cell(row=row, column=col).value for col in range(1, max_col + 1)):
            for col in range(1, max_col + 1):
                sheet.cell(row=row, column=col).fill = GRAY_FILL
            break


def highlight_header_row(sheet, header_row=7):
    max_col = sheet.max_column
    for col in range(1, max_col + 1):
        sheet.cell(row=header_row, column=col).fill = GRAY_FILL


def get_column_index_by_header(sheet, header_name, header_row=1):
    for col in range(1, sheet.max_column + 1):
        cell = sheet.cell(row=header_row, column=col)
        if str(cell.value).strip().lower() == header_name.strip().lower():
            return col
    raise ValidationError(
        f"Header '{header_name}' not found in row {header_row}")


def freeze_top_and_filter(sheet):
    sheet.freeze_panes = "A2"
    header_row = 1
    last_col = get_column_letter(sheet.max_column)
    last_row = sheet.max_row
    sheet.auto_filter.ref = f"A{header_row}:{last_col}{last_row}"


def sort_sheet_by_column(sheet, col_index, header_row, last_row):
    data = []
    for row in range(header_row + 1, last_row + 1):
        row_values = [sheet.cell(row=row, column=col).value for col in range(
            1, sheet.max_column + 1)]
        data.append((sheet.cell(row=row, column=col_index).value, row_values))

    data.sort(key=lambda x: (x[0] if x[0] is not None else float('inf')))

    for i, (_, row_values) in enumerate(data, start=header_row + 1):
        for col_idx, value in enumerate(row_values, start=1):
            sheet.cell(row=i, column=col_idx).value = value


def remove_empty_columns(sheet):
    for col in range(sheet.max_column, 0, -1):
        if all(sheet.cell(row=row, column=col).value in (None, "") for row in range(1, sheet.max_row+1)):
            sheet.delete_cols(col)


def remove_columns_by_header(sheet, headers):
    header_row = 1
    for header in headers:
        try:
            col = get_column_index_by_header(sheet, header, header_row)
            sheet.delete_cols(col)
        except ValueError:
            continue


def drop_rows_with_empty_item(sheet):
    item_col = get_column_index_by_header(sheet, "Item", 1)
    for row in range(sheet.max_row, 1, -1):
        if not sheet.cell(row=row, column=item_col).value:
            sheet.delete_rows(row)


def clear_all_highlighting(sheet):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            sheet.cell(row=row, column=col).fill = PatternFill()


def apply_filter_top(sheet):
    last_col = get_column_letter(sheet.max_column)
    sheet.auto_filter.ref = f"A1:{last_col}{sheet.max_row}"


def remove_footer_and_mech_rows(sheet):
    footer_start = get_footer_row(sheet)
    sheet.delete_rows(footer_start, sheet.max_row - footer_start + 1)


def color_row(sheet, row, fill):
    for col in range(1, sheet.max_column + 1):
        sheet.cell(row=row, column=col).fill = fill


def delete_above_header(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=20):
        if row[0].value == "RegID":
            header_row = row[0].row
            for _ in range(header_row - 1):
                sheet.delete_rows(1)
            return
    raise ValidationError("Header row not found")


def autofit_columns(sheet):
    col_widths = {}

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            if cell.value is not None:
                col_letter = get_column_letter(cell.column)
                val_len = len(str(cell.value))
                if col_letter not in col_widths or val_len > col_widths[col_letter]:
                    col_widths[col_letter] = val_len

    for col_letter, width in col_widths.items():
        sheet.column_dimensions[col_letter].width = width + 2
