from copy import copy
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import FilterColumn, Filters
from .utils import *
 
# Colors
FILL_PURPLE = PatternFill(start_color="B44CB5", end_color="B44CB5", fill_type="solid")  # dark purple
FILL_BLUE   = PatternFill(start_color="FF0099FF", end_color="FF0099FF", fill_type="solid")  # dark blue
FILL_GREEN  = PatternFill(start_color="FF00CC00", end_color="FF00CC00", fill_type="solid")  # dark green
FILL_GRAY   = PatternFill(start_color="FF808080", end_color="FF808080", fill_type="solid")  # gray for discount/coupon header
 
 
def process_step_5(workbook):
    # Copy Step 4 to Step 5
    step5 = copy_sheet(workbook, "Step 4", "Step 5")
 
    item_col     = get_column_index_by_header(step5, "Item", 1)
    customer_col = get_column_index_by_header(step5, "Customer", 1)
    uid_col      = get_column_index_by_header(step5, "UID", 1)
    amount_col   = get_column_index_by_header(step5, "Amount", 1)
    regid_col    = get_column_index_by_header(step5, "RegID", 1)
 
    # --- Fix UID column number format so it shows full number not 3E+09 ---
    for row in range(2, step5.max_row + 1):
        cell = step5.cell(row=row, column=uid_col)
        if cell.value:
            cell.number_format = '0'
 
    # --- Keywords that are NOT purple even if amount is 0 ---
    # Term rows belong to Mailbox only, not Step 5 purple
    MAILBOX_ONLY_KEYWORDS = ["term:", "term "]
 
    # --- Color purple rows: Amount == 0, petty cash, void, regular:saved, tip ---
    # Excludes mailbox-only keywords like Term
    purple_rows = set()
    for row in range(2, step5.max_row + 1):
        item_val   = str(step5.cell(row=row, column=item_col).value or "").strip().lower()
        amount_val = step5.cell(row=row, column=amount_col).value
 
        # Skip term rows — they belong to Mailbox only
        if any(kw in item_val for kw in MAILBOX_ONLY_KEYWORDS):
            continue
 
        # Amount is zero
        is_zero_amount = (amount_val == 0 or amount_val == 0.0)
 
        # Petty cash (negative values) — also catches typo "petty cahs"
        is_petty_cash = "petty cash" in item_val or "petty cahs" in item_val
 
        # Void transactions
        is_void = "void" in item_val
 
        # Regular : Saved
        is_regular_saved = "regular" in item_val and "saved" in item_val
 
        # Tip transactions
        is_tip = "tip" in item_val
 
        if is_zero_amount or is_petty_cash or is_void or is_regular_saved or is_tip:
            color_row(step5, row, FILL_PURPLE)
            purple_rows.add(row)
 
    # --- Color E-Scribers rows blue ONLY if not already purple ---
    for row in range(2, step5.max_row + 1):
        if row in purple_rows:
            continue  # don't overwrite purple with blue
        customer_val = str(step5.cell(row=row, column=customer_col).value or "").strip().lower()
        if any(kw in customer_val for kw in ["e-scriber", "escriber"]):
            color_row(step5, row, FILL_BLUE)
 
    # --- Find mailbox RegIDs first ---
    MAILBOX_KEYWORDS  = ["mailbox"]
    RENEW_KEYWORDS    = ["renew"]
    TERM_KEYWORDS     = ["term"]
    SETUP_KEYWORDS    = ["setup fee", "set up fee"]
    INCLUDES_KEYWORDS = ["includes", "free month"]
 
    # Keywords that qualify a row for mailbox tab (coupon handled separately)
    MAILBOX_ROW_KEYWORDS = [
        "mailbox", "renew", "term", "setup fee", "set up fee",
        "includes", "free month", "late fee"
    ]
 
    # Explicit exclusions — never include even if RegID matches
    MAILBOX_EXCLUSION_KEYWORDS = ["manila", "envelope", "bubble"]
 
    # Step 1: collect all RegIDs that have a mailbox-related item
    mailbox_regids = set()
    for row in range(2, step5.max_row + 1):
        item_val = str(step5.cell(row=row, column=item_col).value or "").strip().lower()
        is_mailbox_item = (
            any(kw in item_val for kw in MAILBOX_KEYWORDS) or
            any(kw in item_val for kw in RENEW_KEYWORDS) or
            any(kw in item_val for kw in SETUP_KEYWORDS) or
            any(kw in item_val for kw in INCLUDES_KEYWORDS)
        )
        if is_mailbox_item:
            regid = step5.cell(row=row, column=regid_col).value
            if regid:
                mailbox_regids.add(regid)
 
    # Step 2: for each mailbox RegID, collect rows STARTING from
    # the first mailbox/renew/setup/includes row
    mailbox_rows = []
    mailbox_row_set = set()
 
    from collections import defaultdict
    regid_rows = defaultdict(list)
    for row in range(2, step5.max_row + 1):
        regid = step5.cell(row=row, column=regid_col).value
        if regid in mailbox_regids:
            regid_rows[regid].append(row)
 
    for regid, rows in regid_rows.items():
        first_mailbox_idx = None
        for i, row in enumerate(rows):
            item_val = str(step5.cell(row=row, column=item_col).value or "").strip().lower()
            if any(kw in item_val for kw in MAILBOX_ROW_KEYWORDS):
                first_mailbox_idx = i
                break
 
        if first_mailbox_idx is None:
            continue
 
        for row in rows[first_mailbox_idx:]:
            if row in mailbox_row_set:
                continue
 
            item_val = str(step5.cell(row=row, column=item_col).value or "").strip().lower()
 
            if any(kw in item_val for kw in MAILBOX_EXCLUSION_KEYWORDS):
                continue
 
            is_mailbox_row = any(kw in item_val for kw in MAILBOX_ROW_KEYWORDS)
            is_coupon = "coupon" in item_val
 
            if not is_mailbox_row and not is_coupon:
                continue
 
            mailbox_rows.append(row)
            mailbox_row_set.add(row)
 
            # Color green in Step 5 only if NOT a term row
            if not any(kw in item_val for kw in TERM_KEYWORDS):
                color_row(step5, row, FILL_GREEN)
 
    # Sort mailbox_rows by row number to preserve original order
    mailbox_rows.sort()
 
    # -----------------------------------------------------------------------
    # ADD HELPER COLUMN for filter-based default view (purple rows only)
    # -----------------------------------------------------------------------
    helper_col = step5.max_column + 1
    helper_col_letter = get_column_letter(helper_col)
 
    step5.cell(row=1, column=helper_col).value = "_filter"
 
    for row in range(2, step5.max_row + 1):
        step5.cell(row=row, column=helper_col).value = "PURPLE" if row in purple_rows else "OTHER"
 
    step5.freeze_panes = "A2"
    last_col = get_column_letter(step5.max_column)
    step5.auto_filter.ref = f"A1:{last_col}{step5.max_row}"
 
   # filter_col_index = helper_col - 1  # 0-based
    #fc = FilterColumn(colId=filter_col_index)
    #fc.filters = Filters(filter=["PURPLE"])
    #step5.auto_filter.filterColumn.append(fc)
 
    step5.column_dimensions[helper_col_letter].hidden = True
 
    # --- Copy mailbox rows to Mailbox tab ---
    # Pass helper_col so _copy_rows_to_tab knows to exclude it
    _copy_rows_to_tab(step5, workbook, "Mailbox", mailbox_rows, exclude_col=helper_col)
 
    # --- Build Mailbox Working tab ---
    build_mailbox_working(step5, workbook, mailbox_rows)
 
    # --- Build Void-Discount-Coupons tab ---
    build_void_discount_coupons(step5, workbook, purple_rows)
 
    # --- Autofit columns ---
    autofit_columns(step5)
 
 
def build_void_discount_coupons(source, workbook, purple_rows):
    """
    Build 'Void-Discount-Coupons' tab.
 
    Layout:
    -------
    TOP SECTION — Purple rows (Void, Petty Cash, Regular:Saved, $0 amount, Tip)
      - Header row: styled with purple fill + filter + freeze
      - Data rows: purple fill (same FILL_PURPLE as Step 5)
 
    BLANK ROW (separator)
 
    BOTTOM SECTION — Discounts & Coupons
      - Sub-header row: gray fill, bold, with label "Retail Discount for the [period]"
        and column headers: UID, RegID, Date, Time, Item, Tender, Customer, Amount
      - Data rows: white/no fill
 
    Rules:
    - Discount/Coupon rows: item contains 'coupon', 'discount'
    - These are sourced fresh from step5 (not filtered by purple_rows)
    - Both sections share same columns: UID, RegID, Date, Time, Item, Tender, Customer, Amount
    - Top header has freeze panes at A2 and autofilter
    - Amount column formatted as currency
    - UID formatted as full number (no scientific notation)
    """
    TAB_NAME = "Void-Discount-Coupons"
 
    if TAB_NAME in workbook.sheetnames:
        del workbook[TAB_NAME]
    ws = workbook.create_sheet(TAB_NAME)
 
    uid_col      = get_column_index_by_header(source, "UID", 1)
    regid_col    = get_column_index_by_header(source, "RegID", 1)
    date_col     = get_column_index_by_header(source, "Date", 1)
    time_col     = get_column_index_by_header(source, "Time", 1)
    item_col     = get_column_index_by_header(source, "Item", 1)
    tender_col   = get_column_index_by_header(source, "Tender", 1)
    customer_col = get_column_index_by_header(source, "Customer", 1)
    amount_col   = get_column_index_by_header(source, "Amount", 1)
 
    COL_UID      = 1
    COL_REGID    = 2
    COL_DATE     = 3
    COL_TIME     = 4
    COL_ITEM     = 5
    COL_TENDER   = 6
    COL_CUSTOMER = 7
    COL_AMOUNT   = 8
    TOTAL_COLS   = 8
 
    COLUMN_HEADERS = ["UID", "RegID", "Date", "Time", "Item", "Tender", "Customer", "Amount"]
 
    for col_idx, header in enumerate(COLUMN_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.fill  = FILL_PURPLE
        cell.font  = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
 
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(TOTAL_COLS)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"
 
    write_row = 2
    for src_row in sorted(purple_rows):
        _write_row_to_sheet(
            ws, source, write_row, src_row,
            uid_col, regid_col, date_col, time_col,
            item_col, tender_col, customer_col, amount_col,
            COL_UID, COL_REGID, COL_DATE, COL_TIME,
            COL_ITEM, COL_TENDER, COL_CUSTOMER, COL_AMOUNT,
            TOTAL_COLS, fill=FILL_PURPLE
        )
        write_row += 1
 
    write_row += 1  # blank separator row
 
    period_label = _get_period_label(source, date_col)
    gray_header_row = write_row
 
    label_cell = ws.cell(row=gray_header_row, column=1)
    label_cell.value = f"Retail Discount for the {period_label}"
    label_cell.fill  = FILL_GRAY
    label_cell.font  = Font(bold=True, color="FFFFFF")
    label_cell.alignment = Alignment(horizontal="left", vertical="center")
 
    for col_idx, header in enumerate(COLUMN_HEADERS, start=1):
        cell = ws.cell(row=gray_header_row, column=col_idx)
        if col_idx == 1:
            pass
        else:
            cell.value = header
            cell.fill  = FILL_GRAY
            cell.font  = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
 
    write_row += 1
 
    DISCOUNT_KEYWORDS = ["coupon", "discount"]
 
    for src_row in range(2, source.max_row + 1):
        item_val = str(source.cell(row=src_row, column=item_col).value or "").strip().lower()
        if any(kw in item_val for kw in DISCOUNT_KEYWORDS):
            _write_row_to_sheet(
                ws, source, write_row, src_row,
                uid_col, regid_col, date_col, time_col,
                item_col, tender_col, customer_col, amount_col,
                COL_UID, COL_REGID, COL_DATE, COL_TIME,
                COL_ITEM, COL_TENDER, COL_CUSTOMER, COL_AMOUNT,
                TOTAL_COLS, fill=None
            )
            write_row += 1
 
    autofit_columns(ws)
 
 
def _write_row_to_sheet(
    ws, source, write_row, src_row,
    uid_col, regid_col, date_col, time_col,
    item_col, tender_col, customer_col, amount_col,
    COL_UID, COL_REGID, COL_DATE, COL_TIME,
    COL_ITEM, COL_TENDER, COL_CUSTOMER, COL_AMOUNT,
    total_cols, fill
):
    """Helper: write a single source row into target ws row with optional fill."""
    col_map = {
        COL_UID:      uid_col,
        COL_REGID:    regid_col,
        COL_DATE:     date_col,
        COL_TIME:     time_col,
        COL_ITEM:     item_col,
        COL_TENDER:   tender_col,
        COL_CUSTOMER: customer_col,
        COL_AMOUNT:   amount_col,
    }
    for out_col, src_col in col_map.items():
        src_cell = source.cell(row=src_row, column=src_col)
        tgt_cell = ws.cell(row=write_row, column=out_col)
        tgt_cell.value = src_cell.value
        if fill:
            tgt_cell.fill = fill
        if out_col == COL_UID:
            tgt_cell.number_format = '0'
        elif out_col == COL_AMOUNT:
            tgt_cell.number_format = '$#,##0.00'
            amount_val = src_cell.value
            if amount_val is not None and isinstance(amount_val, (int, float)) and amount_val < 0:
                tgt_cell.font = Font(color="FF0000")
        elif out_col == COL_DATE:
            tgt_cell.number_format = 'mm/dd/yyyy'
 
 
def _get_period_label(source, date_col):
    """Extract a period label like 'March 2026' from the first non-empty date."""
    import datetime
    for row in range(2, source.max_row + 1):
        val = source.cell(row=row, column=date_col).value
        if val:
            try:
                if isinstance(val, (datetime.datetime, datetime.date)):
                    return val.strftime("%B %Y")
                from dateutil import parser as dateparser
                parsed = dateparser.parse(str(val))
                return parsed.strftime("%B %Y")
            except Exception:
                pass
    return "the period"
 
 
def _copy_rows_to_tab(source, workbook, tab_name, row_indices, exclude_col=None):
    """
    Copy specified rows from source sheet into a named tab.
    Always deletes and recreates tab clean. All rows forced to dark green fill.
    exclude_col: 1-based column index to skip (e.g. the hidden _filter helper column).
    """
    if tab_name in workbook.sheetnames:
        del workbook[tab_name]
 
    ws = workbook.create_sheet(tab_name)

    # Determine which columns to copy — skip exclude_col
    cols_to_copy = [
        c for c in range(1, source.max_column + 1)
        if c != exclude_col
    ]

    # Write header row using only the included columns
    for out_col, src_col in enumerate(cols_to_copy, start=1):
        ws.cell(row=1, column=out_col).value = source.cell(row=1, column=src_col).value

    freeze_top_and_filter(ws)
    highlight_rows(ws, header_row=1)
    autofit_columns(ws)
 
    # Write data rows
    write_row = 2
    for src_row in row_indices:
        for out_col, src_col in enumerate(cols_to_copy, start=1):
            src_cell = source.cell(row=src_row, column=src_col)
            tgt_cell = ws.cell(row=write_row, column=out_col)
            tgt_cell.value         = src_cell.value
            tgt_cell.fill          = FILL_GREEN
            tgt_cell.number_format = src_cell.number_format
        write_row += 1
 
 
def build_mailbox_working(source, workbook, mailbox_rows):
    """
    Build 'Mailbox Working' sheet from mailbox rows.
 
    Columns: UID, RegID, Date, Time, Item, Mailbox #, Mailbox Type, [blank],
             Tender, Customer, Amount, Tax, Total Amount
 
    Tax logic (UPDATED):
    - ONLY if a RegID group contains a COUPON row → ALL rows in that group get 0 tax
    - Late fees and discounts alone do NOT zero out tax anymore
    - All other rows: Tax = Amount * 8.875%
    """
    TAB_NAME = "Mailbox Working"
 
    if TAB_NAME in workbook.sheetnames:
        del workbook[TAB_NAME]
    ws = workbook.create_sheet(TAB_NAME)
 
    src_uid_col      = get_column_index_by_header(source, "UID", 1)
    src_regid_col    = get_column_index_by_header(source, "RegID", 1)
    src_date_col     = get_column_index_by_header(source, "Date", 1)
    src_time_col     = get_column_index_by_header(source, "Time", 1)
    src_item_col     = get_column_index_by_header(source, "Item", 1)
    src_tender_col   = get_column_index_by_header(source, "Tender", 1)
    src_customer_col = get_column_index_by_header(source, "Customer", 1)
    src_amount_col   = get_column_index_by_header(source, "Amount", 1)
 
    COL_UID      = 1
    COL_REGID    = 2
    COL_DATE     = 3
    COL_TIME     = 4
    COL_ITEM     = 5
    COL_MBOX_NUM = 6
    COL_MBOX_TYP = 7
    # col 8 = blank
    COL_TENDER   = 9
    COL_CUSTOMER = 10
    COL_AMOUNT   = 11
    COL_TAX      = 12
    COL_TOTAL    = 13
 
    headers = {
        COL_UID: "UID", COL_REGID: "RegID", COL_DATE: "Date", COL_TIME: "Time",
        COL_ITEM: "Item", COL_MBOX_NUM: "Mailbox #", COL_MBOX_TYP: "Mailbox Type",
        COL_TENDER: "Tender", COL_CUSTOMER: "Customer",
        COL_AMOUNT: "Amount", COL_TAX: "Tax", COL_TOTAL: "Total Amount"
    }
    for col, header in headers.items():
        ws.cell(row=1, column=col).value = header
 
    format_header(ws, header_row=1)
    freeze_top_and_filter(ws)
    highlight_rows(ws, header_row=1)
 
    import re
 
    def extract_mailbox_number(item_text):
        if not item_text:
            return None
        match = re.search(r'mailbox\s*#(\d+)', item_text, re.IGNORECASE)
        if match:
            return int(match.group(1))
        match = re.search(r'mailbox\s+(\d+)', item_text, re.IGNORECASE)
        if match:
            return int(match.group(1))
        return None
 
    def extract_mailbox_type(item_text):
        if item_text and "business" in item_text.lower():
            return " BUSINESS"
        return None
 
    def is_term_or_child_row(item_text):
        item_lower = (item_text or "").strip().lower()
        return (
            item_lower.startswith("term") or
            "  term" in item_lower or
            item_lower.startswith("term:")
        )
 
    # Only COUPON rows trigger zero-tax for the whole group
    zero_tax_regids = set()
    for src_row in mailbox_rows:
        item_val_scan = str(source.cell(src_row, column=src_item_col).value or "").lower()
        if "coupon" in item_val_scan:
            regid_scan = source.cell(src_row, column=src_regid_col).value
            if regid_scan:
                zero_tax_regids.add(regid_scan)
 
    write_row = 2
    last_mbox_row_for_regid = {}
 
    for src_row in mailbox_rows:
        item_val   = str(source.cell(src_row, column=src_item_col).value or "")
        amount_val = source.cell(src_row, column=src_amount_col).value or 0
        regid_val  = source.cell(src_row, column=src_regid_col).value
 
        # --- Mailbox # ---
        mbox_num = extract_mailbox_number(item_val)
        if mbox_num is not None:
            ws.cell(row=write_row, column=COL_MBOX_NUM).value = mbox_num
            last_mbox_row_for_regid[regid_val] = write_row
        else:
            parent_row = last_mbox_row_for_regid.get(regid_val)
            if parent_row:
                ws.cell(row=write_row, column=COL_MBOX_NUM).value = f"=F{parent_row}"
            else:
                ws.cell(row=write_row, column=COL_MBOX_NUM).value = None
 
        # --- Mailbox Type ---
        mbox_type = extract_mailbox_type(item_val)
        if mbox_type:
            ws.cell(row=write_row, column=COL_MBOX_TYP).value = mbox_type
            last_mbox_row_for_regid[str(regid_val) + "_type"] = mbox_type
        elif is_term_or_child_row(item_val):
            inherited_type = last_mbox_row_for_regid.get(str(regid_val) + "_type")
            ws.cell(row=write_row, column=COL_MBOX_TYP).value = inherited_type
        else:
            ws.cell(row=write_row, column=COL_MBOX_TYP).value = None
 
        # --- Core fields ---
        ws.cell(row=write_row, column=COL_UID).value      = source.cell(src_row, column=src_uid_col).value
        ws.cell(row=write_row, column=COL_REGID).value    = source.cell(src_row, column=src_regid_col).value
        ws.cell(row=write_row, column=COL_DATE).value     = source.cell(src_row, column=src_date_col).value
        ws.cell(row=write_row, column=COL_TIME).value     = source.cell(src_row, column=src_time_col).value
        ws.cell(row=write_row, column=COL_ITEM).value     = item_val
        ws.cell(row=write_row, column=COL_TENDER).value   = source.cell(src_row, column=src_tender_col).value
        ws.cell(row=write_row, column=COL_CUSTOMER).value = source.cell(src_row, column=src_customer_col).value
        ws.cell(row=write_row, column=COL_AMOUNT).value   = amount_val
 
        # --- Tax ---
        if regid_val in zero_tax_regids:
            ws.cell(row=write_row, column=COL_TAX).value = 0
        else:
            ws.cell(row=write_row, column=COL_TAX).value = f"=K{write_row}*8.875%"
 
        # --- Total = Amount + Tax ---
        ws.cell(row=write_row, column=COL_TOTAL).value = f"=K{write_row}+L{write_row}"
 
        # --- Number formats ---
        ws.cell(row=write_row, column=COL_UID).number_format      = '0'
        ws.cell(row=write_row, column=COL_AMOUNT).number_format   = '$#,##0.00'
        ws.cell(row=write_row, column=COL_TAX).number_format      = '$#,##0.00'
        ws.cell(row=write_row, column=COL_TOTAL).number_format    = '$#,##0.00'
        ws.cell(row=write_row, column=COL_DATE).number_format     = 'mm/dd/yyyy'
 
        # --- Row color: green ---
        for col in range(1, 14):
            ws.cell(row=write_row, column=col).fill = FILL_GREEN
 
        write_row += 1
 
    # --- Totals row ---
    last_data_row = write_row - 1
    totals_row = write_row + 1  # one blank row gap
 
    ws.cell(row=totals_row, column=COL_AMOUNT).value = f"=SUM(K2:K{last_data_row})"
    ws.cell(row=totals_row, column=COL_TAX).value    = f"=SUM(L2:L{last_data_row})"
    ws.cell(row=totals_row, column=COL_TOTAL).value  = f"=SUM(M2:M{last_data_row})"
 
    ws.cell(row=totals_row, column=COL_AMOUNT).number_format = '$#,##0.00'
    ws.cell(row=totals_row, column=COL_TAX).number_format    = '$#,##0.00'
    ws.cell(row=totals_row, column=COL_TOTAL).number_format  = '$#,##0.00'
 
    ws.cell(row=totals_row, column=COL_CUSTOMER).value = "Total:"
    ws.cell(row=totals_row, column=COL_CUSTOMER).font  = Font(bold=True)
    ws.cell(row=totals_row, column=COL_AMOUNT).font    = Font(bold=True)
    ws.cell(row=totals_row, column=COL_TAX).font       = Font(bold=True)
    ws.cell(row=totals_row, column=COL_TOTAL).font     = Font(bold=True)
 
    autofit_columns(ws)