import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font


# =====================================================
# CLEAN TEXT
# =====================================================
def clean_text(val):
    if val is None:
        return ""
    return re.sub(r"[^a-zA-Z0-9 ]", " ", str(val)).lower().strip()


# =====================================================
# PARSE AMOUNT
# =====================================================
def parse_amount(val):
    if val is None or val == "":
        return 0.0
    val = str(val).replace("$", "").replace(",", "").strip()
    try:
        return float(val)
    except:
        return 0.0


# =====================================================
# SAFE CELL READ
# =====================================================
def safe_cell(sheet, row, col):
    try:
        return sheet.cell(row=row, column=col).value
    except:
        return None


# =====================================================
# STEP 6
# =====================================================
def process_step_6(workbook):

    step5 = workbook["Step 5"]

    if "Step 6" in workbook.sheetnames:
        del workbook["Step 6"]

    step6 = workbook.create_sheet("Step 6")

    # HEADERS
    headers = [
        "UID", "RegID", "Date", "Time", "Item",
        "Tender", "Customer", "Amount", "Tax", "Total Amount", "Taxable"
    ]
    step6.append(headers)

    item_col = 5
    amount_col = 8

    TAX_RATE = 0.08875

    EXCLUDE_KEYWORDS = [
        "coupon", "discount", "void", "term",
        "late fee", "mailbox", "setup fee", "renew"
    ]

    VALID_ITEMS = [
        "copies", "fax", "lamination", "passport", "postcard",
        "printing", "scan", "box", "envelope", "tape",
        "bubble", "paper", "stamp", "tube", "crate",
        "packing", "rental", "post office"
    ]

    TAXABLE_KEYWORDS = [
        "copies", "fax", "lamination",
        "passport", "postcard",
        "printing", "scan"
    ]

    def is_no_fill(cell):
        fill = cell.fill
        if fill is None:
            return True
        if fill.fill_type is None:
            return True
        if getattr(fill, "patternType", None) is None:
            return True
        return False

    # =====================================================
    # COLLECT ROWS FIRST (so we can SORT)
    # =====================================================
    processed_rows = []

    total_amount = 0.0
    total_tax    = 0.0
    total_total  = 0.0

    for row in range(2, step5.max_row + 1):

        item_raw   = safe_cell(step5, row, item_col)
        item_clean = clean_text(item_raw)
        amount     = parse_amount(safe_cell(step5, row, amount_col))

        if item_clean == "" and amount == 0:
            continue

        if not is_no_fill(step5.cell(row=row, column=item_col)):
            continue

        if any(k in item_clean for k in EXCLUDE_KEYWORDS):
            continue

        if not any(k in item_clean for k in VALID_ITEMS):
            continue

        # TAX
        is_taxable   = any(k in item_clean for k in TAXABLE_KEYWORDS)
        tax          = round(amount * TAX_RATE, 2) if is_taxable else 0.0
        taxable_flag = "y" if is_taxable else "n"
        total        = round(amount + tax, 2)

        # BUILD ROW
        row_data = []
        for c in range(1, 11):
            row_data.append(safe_cell(step5, row, c))

        row_data[7] = amount
        row_data[8] = tax
        row_data[9] = total

        processed_rows.append((item_clean, row_data, taxable_flag))

        total_amount += amount
        total_tax    += tax
        total_total  += total

    # =====================================================
    # SORT A → Z BY ITEM
    # =====================================================
    processed_rows.sort(key=lambda x: x[0])

    # =====================================================
    # WRITE DATA
    # =====================================================
    for _, row_data, taxable_flag in processed_rows:
        step6.append(row_data + [taxable_flag])

    # =====================================================
    # APPLY $ FORMAT
    # =====================================================
    for row in range(2, step6.max_row + 1):
        for col in [8, 9, 10]:
            cell = step6.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'

    # =====================================================
    # TOTALS ROW
    # =====================================================
    step6.append([])
    step6.append([
        "", "", "", "", "", "", "TOTALS",
        total_amount,
        total_tax,
        total_total,
        ""
    ])

    last_row = step6.max_row
    for col in [8, 9, 10]:
        step6.cell(row=last_row, column=col).number_format = '$#,##0.00'
        step6.cell(row=last_row, column=col).font          = Font(bold=True)
    step6.cell(row=last_row, column=7).font = Font(bold=True)

    # =====================================================
    # FORMAT + FILTER + FREEZE
    # =====================================================
    for col in range(1, 12):
        step6.column_dimensions[get_column_letter(col)].width = 20

    step6.freeze_panes    = "A2"
    step6.auto_filter.ref = f"A1:K{step6.max_row}"

    # =====================================================
    # BUILD RETAIL TAB — exact copy of Step 6
    # =====================================================
    _build_retail_tab(workbook, step6, headers, processed_rows, total_amount, total_tax, total_total)

    return step6


# =====================================================
# RETAIL TAB
# =====================================================
def _build_retail_tab(workbook, step6, headers, processed_rows, total_amount, total_tax, total_total):
    """
    Creates a 'Retail' tab that is an exact copy of Step 6:
    same headers, same sorted data rows, same totals row,
    same column widths, freeze panes, and autofilter.
    """
    TAB_NAME = "Retail"

    if TAB_NAME in workbook.sheetnames:
        del workbook[TAB_NAME]

    retail = workbook.create_sheet(TAB_NAME)

    # --- Headers ---
    retail.append(headers)

    # --- Data rows (same sorted order as Step 6) ---
    for _, row_data, taxable_flag in processed_rows:
        retail.append(row_data + [taxable_flag])

    # --- Apply $ format ---
    for row in range(2, retail.max_row + 1):
        for col in [8, 9, 10]:
            cell = retail.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'

    # --- Totals row ---
    retail.append([])
    retail.append([
        "", "", "", "", "", "", "TOTALS",
        total_amount,
        total_tax,
        total_total,
        ""
    ])

    last_row = retail.max_row
    for col in [8, 9, 10]:
        retail.cell(row=last_row, column=col).number_format = '$#,##0.00'
        retail.cell(row=last_row, column=col).font          = Font(bold=True)
    retail.cell(row=last_row, column=7).font = Font(bold=True)

    # --- Column widths, freeze, filter ---
    for col in range(1, 12):
        retail.column_dimensions[get_column_letter(col)].width = 20

    retail.freeze_panes    = "A2"
    retail.auto_filter.ref = f"A1:K{retail.max_row}"