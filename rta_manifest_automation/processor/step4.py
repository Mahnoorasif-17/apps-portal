from copy import copy
from openpyxl.styles import PatternFill, Font
from .utils import *
from datetime import date

FILL_LIGHT_ORANGE = PatternFill(
    start_color="FFFFD580", end_color="FFFFD580", fill_type="solid")
FILL_LIGHT_PURPLE = PatternFill(
    start_color="FFE5CCFF", end_color="FFE5CCFF", fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(
    start_color="FFCCEEFF", end_color="FFCCEEFF", fill_type="solid")
FILL_LIGHT_GREEN = PatternFill(
    start_color="FFCCFFCC", end_color="FFCCFFCC", fill_type="solid")


def process_step_4(workbook):
    step4 = copy_sheet(workbook, "Step 3", "Step 4")

    remove_empty_columns(step4)
    remove_columns_by_header(step4, ["SubTotal", "Tax", "Total", "User"])
    drop_rows_with_empty_item(step4)
    remove_footer_and_mech_rows(step4)
    clear_all_highlighting(step4)

    format_header(step4, header_row=1)
    highlight_header_row(step4, header_row=1)
    autofit_columns(step4)

    # --- Add UID column AFTER all cleanup steps ---
    add_uid_column(step4)
    # ----------------------------------------------

    distribute_items_to_sheets(step4, workbook)


def add_uid_column(sheet):
    # Insert blank column at position A (shift everything right)
    sheet.insert_cols(1)
    # Set header
    sheet.cell(row=1, column=1).value = "UID"

    # Get date prefix from the last data row's Date column (yymmdd)
    date_col = get_column_index_by_header(sheet, "Date", 1)
    last_date = None
    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=date_col).value
        if val:
            last_date = val
    if last_date is None:
        from datetime import date
        last_date = date.today()
    if hasattr(last_date, 'strftime'):
        prefix = last_date.strftime("%y%m%d")
    else:
        from datetime import datetime
        prefix = datetime.strptime(str(last_date), "%Y-%m-%d").strftime("%y%m%d")

    # Fill UID for every data row
    counter = 1
    for row in range(2, sheet.max_row + 1):
        # Only assign UID if the row has any data
        if any(sheet.cell(row=row, column=col).value for col in range(2, sheet.max_column + 1)):
            sheet.cell(row=row, column=1).value = int(f"{prefix}{counter:04d}")
            counter += 1


def distribute_items_to_sheets(source, workbook):
    mapping = [
        ("DHL", "dhl", FILL_LIGHT_ORANGE, ["dhl drop off"]),
        ("USPS", "usps", FILL_LIGHT_PURPLE, ["void"]),
        ("FedEx", "fedex", FILL_LIGHT_BLUE,   ["void"]),
        ("UPS", "ups", FILL_LIGHT_GREEN,      ["void"])
    ]
    TAB_COLORS = {
        "DHL": "FFD580", "USPS": "E5CCFF", "FedEx": "CCEEFF", "UPS": "CCFFCC"
    }

    item_col = get_column_index_by_header(source, "Item", 1)
    regid_col = get_column_index_by_header(source, "RegID", 1)

    service_sheets = {}
    regid_service_map = {}

    # 1st pass: copy regular service rows
    for sheet_name, keyword, fill, excludes in mapping:
        target = workbook.create_sheet(sheet_name)
        target.sheet_properties.tabColor = TAB_COLORS[sheet_name]
        copy_headers(source, target)
        format_header(target, header_row=1)
        freeze_top_and_filter(target)
        highlight_rows(target, header_row=1)
        service_sheets[sheet_name] = {"sheet": target, "row": 2, "fill": fill}
        autofit_columns(target)

    for row in range(2, source.max_row + 1):
        val = str(source.cell(row=row, column=item_col).value or "")
        regid = source.cell(row=row, column=regid_col).value
        customer = str(source.cell(
            row=row, column=get_column_index_by_header(source, "Customer")).value or "")

        for sheet_name, keyword, fill, excludes in mapping:
            if keyword.lower() in val.lower() and not any(ex in val.lower() for ex in excludes):

                tgt = service_sheets[sheet_name]

                # Copy main service row
                color_row(source, row, fill)
                service_row_pos = insert_row_above_regid(
                    source, tgt["sheet"], row, fill, regid, regid_col
                )
                tgt["row"] += 1

                next_row = row + 1
                next_item = ""
                next_regid = None

                if next_row <= source.max_row:
                    next_item = str(source.cell(
                        next_row, column=item_col).value or "")
                    next_regid = source.cell(next_row, column=regid_col).value

                # ---------------------------------------------------
                # copy discount/coupon if present
                # ---------------------------------------------------
                has_existing_discount = (
                    next_regid == regid and
                    any(kw in next_item.lower() for kw in ["discount", "coupon"]) and
                    "void" not in next_item.lower()
                )

                if has_existing_discount:
                    color_row(source, next_row, fill)
                    insert_at = service_row_pos + 1
                    tgt["sheet"].insert_rows(insert_at)

                    for col in range(1, source.max_column + 1):
                        src_cell = source.cell(row=next_row, column=col)
                        tgt_cell = tgt["sheet"].cell(row=insert_at, column=col)

                        tgt_cell.value = src_cell.value
                        tgt_cell.fill = fill
                        tgt_cell.number_format = src_cell.number_format
                    tgt["row"] += 1

                # ---------------------------------------------------
                # Empire Merchants Chelsea 50% discount logic
                # ---------------------------------------------------
                if customer.strip().lower() == "empire merchants chelsea" and not has_existing_discount:
                    insert_at = service_row_pos + 1
                    tgt["sheet"].insert_rows(insert_at)

                    for col in range(1, source.max_column + 1):
                        src_cell = source.cell(row=row, column=col)
                        tgt_cell = tgt["sheet"].cell(row=insert_at, column=col)
                        tgt_cell.value = src_cell.value
                        tgt_cell.fill = fill
                        tgt_cell.number_format = src_cell.number_format

                    # Modify the row to represent 50% discount
                    tgt["sheet"].cell(
                        row=insert_at, column=item_col).value = "50% discount"
                    amount_col = get_column_index_by_header(source, "Amount")
                    original_amount = source.cell(
                        row=row, column=amount_col).value
                    if isinstance(original_amount, (int, float)):
                        tgt["sheet"].cell(
                            row=insert_at, column=amount_col).value = -abs(original_amount) / 2

                    tgt["row"] += 1

                break

    # 2nd pass: handle "Declared value"
    for row in range(2, source.max_row):
        item_val = str(source.cell(row=row, column=item_col).value or "")
        if "declared value" in item_val.lower():
            regid = source.cell(row=row, column=regid_col).value

            # Prefer the row below
            next_row = row + 1
            if next_row <= source.max_row:
                next_item = str(source.cell(
                    next_row, column=item_col).value or "")
                next_regid = source.cell(next_row, column=regid_col).value

                if regid == next_regid:
                    for sheet_name, keyword, _, excludes in mapping:
                        if keyword.lower() in next_item.lower() and not any(ex in next_item.lower() for ex in excludes):
                            color_row(
                                source, row, service_sheets[sheet_name]["fill"])
                            tgt = service_sheets[sheet_name]
                            # fixed: insert ABOVE so declared value comes before service row
                            insert_row_above_regid(
                                source_sheet=source,
                                target_sheet=tgt["sheet"],
                                source_row=row,
                                target_fill=tgt["fill"],
                                regid=regid,
                                regid_col=regid_col
                            )
                            tgt["row"] += 1
                            break
                    continue

            # Else search entire regid group
            for search_row in range(2, source.max_row + 1):
                if source.cell(search_row, column=regid_col).value == regid:
                    search_item = str(source.cell(
                        search_row, column=item_col).value or "")
                    for sheet_name, keyword, _, excludes in mapping:
                        if keyword.lower() in search_item.lower() and not any(ex in search_item.lower() for ex in excludes):
                            color_row(
                                source, row, service_sheets[sheet_name]["fill"])
                            tgt = service_sheets[sheet_name]
                            # fixed: insert ABOVE so declared value comes before service row
                            insert_row_above_regid(
                                source_sheet=source,
                                target_sheet=tgt["sheet"],
                                source_row=row,
                                target_fill=tgt["fill"],
                                regid=regid,
                                regid_col=regid_col
                            )
                            tgt["row"] += 1
                            break
                    break

    # --- Build 3PL sheet: all services combined, NO Account tender ---
    # (auto-excludes E-Scribers, Empire, Feshaire - all use Account tender)
    build_3pl_sheet(workbook, service_sheets)
    # -----------------------------------------------------------------

    # --- Build Account sheets: E-Scribers, Empire, Feshaire ---
    # Only created if data exists for that customer
    build_account_sheets(source, workbook, mapping)
    # ----------------------------------------------------------


def build_3pl_sheet(workbook, service_sheets):
    """
    3PL = all service sheets (UPS, FedEx, USPS, DHL) combined.
    Excludes:
      - Rows where Tender == 'Account' (E-Scribers, Empire, Feshaire)
      - Rows where Item contains 'void', 'discount', 'coupon'
      - Rows with no UID (blank col A) — these are discount/coupon inserted rows
    """
    SERVICE_ORDER = ["DHL", "USPS", "FedEx", "UPS"]
    EXCLUDED_TENDER = "account"
    EXCLUDED_ITEM_KEYWORDS = ["void", "discount", "coupon"]

    if "3PL" in workbook.sheetnames:
        del workbook["3PL"]
    sheet_3pl = workbook.create_sheet("3PL")

    tender_col = None
    item_col_3pl = None
    uid_col = None
    headers_written = False

    for sname in SERVICE_ORDER:
        if sname not in workbook.sheetnames:
            continue
        src = workbook[sname]

        if not headers_written:
            for col in range(1, src.max_column + 1):
                sheet_3pl.cell(row=1, column=col).value = src.cell(row=1, column=col).value
            for col in range(1, src.max_column + 1):
                header_val = str(src.cell(row=1, column=col).value or "").strip().lower()
                if header_val == "tender":
                    tender_col = col
                if header_val == "item":
                    item_col_3pl = col
                if header_val == "uid":
                    uid_col = col
            format_header(sheet_3pl, header_row=1)
            freeze_top_and_filter(sheet_3pl)
            highlight_rows(sheet_3pl, header_row=1)
            autofit_columns(sheet_3pl)
            headers_written = True

        write_row = sheet_3pl.max_row + 1

        for src_row in range(2, src.max_row + 1):
            if tender_col:
                tender_val = str(src.cell(src_row, column=tender_col).value or "").strip().lower()
                if tender_val == EXCLUDED_TENDER:
                    continue
            if item_col_3pl:
                item_val = str(src.cell(src_row, column=item_col_3pl).value or "").strip().lower()
                if any(kw in item_val for kw in EXCLUDED_ITEM_KEYWORDS):
                    continue
            if uid_col:
                uid_val = src.cell(src_row, column=uid_col).value
                if not uid_val:
                    continue

            for col in range(1, src.max_column + 1):
                src_cell = src.cell(row=src_row, column=col)
                tgt_cell = sheet_3pl.cell(row=write_row, column=col)
                tgt_cell.value = src_cell.value
                tgt_cell.fill = copy(src_cell.fill)
                tgt_cell.number_format = src_cell.number_format
            write_row += 1


def build_account_sheets(source, workbook, mapping):
    """
    Build separate sheets for Account tender customers:
    E-Scribers, Empire, Feshaire.

    E-Scribers: all rows where Customer contains escriber AND Tender=Account
    Empire:     all rows where Customer contains empire AND Tender=Account
                PLUS auto-generate a 50% discount row after every service row
                that doesn't already have one
    Feshaire:   all rows where Customer contains feshaire AND Tender=Account

    Only creates sheet if data exists. Row color matches service type.
    """

    ACCOUNT_CUSTOMERS = [
        ("E-Scribers", ["e-scriber", "escriber"]),
        ("Empire",     ["empire"]),
        ("Feshaire",   ["feshaire", "fashaire"]),
    ]

    EXCLUDED_ITEM_KEYWORDS = ["void"]

    item_col     = get_column_index_by_header(source, "Item", 1)
    tender_col   = get_column_index_by_header(source, "Tender", 1)
    customer_col = get_column_index_by_header(source, "Customer", 1)
    amount_col   = get_column_index_by_header(source, "Amount", 1)
    regid_col    = get_column_index_by_header(source, "RegID", 1)

    def get_service_fill(item_val):
        item_lower = item_val.lower()
        for sheet_name, keyword, fill, excludes in mapping:
            if keyword.lower() in item_lower and not any(ex in item_lower for ex in excludes):
                return fill
        return FILL_LIGHT_GREEN  # default green if no service match

    for sheet_name, customer_keywords in ACCOUNT_CUSTOMERS:

        is_empire = (sheet_name == "Empire")

        # --- Collect all matching rows from source by customer name + Account tender ---
        matching_rows = []
        for row in range(2, source.max_row + 1):
            tender_val   = str(source.cell(row=row, column=tender_col).value or "").strip().lower()
            customer_val = str(source.cell(row=row, column=customer_col).value or "").lower()
            item_val     = str(source.cell(row=row, column=item_col).value or "").lower()

            # Must be Account tender
            if tender_val != "account":
                continue
            # Must match customer — purely by customer name, regardless of RegID
            if not any(kw in customer_val for kw in customer_keywords):
                continue
            # Exclude void
            if any(kw in item_val for kw in EXCLUDED_ITEM_KEYWORDS):
                continue

            matching_rows.append(row)

        if not matching_rows:
            continue

        # Delete and recreate clean
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
        ws = workbook.create_sheet(sheet_name)

        # Copy headers
        for col in range(1, source.max_column + 1):
            ws.cell(row=1, column=col).value = source.cell(row=1, column=col).value

        format_header(ws, header_row=1)
        freeze_top_and_filter(ws)
        highlight_rows(ws, header_row=1)
        autofit_columns(ws)

        write_row = 2

        # --- Write rows ---
        # For Empire: after each non-discount service row, ensure a 50% discount row follows
        i = 0
        while i < len(matching_rows):
            src_row = matching_rows[i]
            item_val     = str(source.cell(src_row, column=item_col).value or "")
            item_lower   = item_val.lower()
            amount_val   = source.cell(src_row, column=amount_col).value
            row_fill     = get_service_fill(item_val)

            # Write the current row
            for col in range(1, source.max_column + 1):
                src_cell = source.cell(row=src_row, column=col)
                tgt_cell = ws.cell(row=write_row, column=col)
                tgt_cell.value = src_cell.value
                tgt_cell.number_format = src_cell.number_format
                tgt_cell.fill = row_fill
            write_row += 1

            if is_empire:
                # Check if this is a service row (not already a discount/coupon row)
                is_service_row = (
                    "discount" not in item_lower and
                    "coupon" not in item_lower and
                    "void" not in item_lower
                )

                if is_service_row:
                    # Check if next row in matching_rows is already a discount for same RegID
                    regid = source.cell(src_row, column=regid_col).value
                    next_is_discount = False
                    if i + 1 < len(matching_rows):
                        next_src_row = matching_rows[i + 1]
                        next_item = str(source.cell(next_src_row, column=item_col).value or "").lower()
                        next_regid = source.cell(next_src_row, column=regid_col).value
                        if next_regid == regid and ("discount" in next_item or "coupon" in next_item):
                            next_is_discount = True

                    if not next_is_discount:
                        # Auto-generate 50% discount row
                        for col in range(1, source.max_column + 1):
                            src_cell = source.cell(row=src_row, column=col)
                            tgt_cell = ws.cell(row=write_row, column=col)
                            tgt_cell.value = src_cell.value
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.fill = row_fill
                        # Clear UID, set Item = '50% discount', set Amount = -50%
                        ws.cell(row=write_row, column=1).value = None  # no UID
                        ws.cell(row=write_row, column=item_col).value = "50% discount"
                        if isinstance(amount_val, (int, float)):
                            ws.cell(row=write_row, column=amount_col).value = -abs(amount_val) / 2
                        write_row += 1

            i += 1

        # --- Net total row at bottom ---
        total_row = write_row + 1
        net_total = 0
        for r in range(2, write_row):
            val = ws.cell(row=r, column=amount_col).value
            if isinstance(val, (int, float)):
                net_total += val

        total_label_col = amount_col - 1 if amount_col > 1 else amount_col
        ws.cell(row=total_row, column=total_label_col).value = "Total:"
        ws.cell(row=total_row, column=total_label_col).font = Font(bold=True)
        total_cell = ws.cell(row=total_row, column=amount_col)
        total_cell.value = round(net_total, 2)
        total_cell.font = Font(bold=True)
        total_cell.number_format = '$#,##0.00'


def insert_row_above_regid(source_sheet, target_sheet, source_row, target_fill, regid, regid_col):
    insert_at = None
    for row in range(2, target_sheet.max_row + 1):
        if target_sheet.cell(row=row, column=regid_col).value == regid:
            insert_at = row
            break

    if insert_at is None:
        insert_at = target_sheet.max_row + 1

    target_sheet.insert_rows(insert_at)

    for col in range(1, source_sheet.max_column + 1):
        src_cell = source_sheet.cell(row=source_row, column=col)
        tgt_cell = target_sheet.cell(row=insert_at, column=col)

        tgt_cell.value = src_cell.value
        tgt_cell.fill = target_fill
        tgt_cell.number_format = src_cell.number_format

    return insert_at


def insert_row_below_regid(source_sheet, target_sheet, source_row, target_fill, regid, regid_col):
    insert_at = None

    # Find LAST occurrence of this RegID
    for row in range(2, target_sheet.max_row + 1):
        if target_sheet.cell(row=row, column=regid_col).value == regid:
            insert_at = row

    if insert_at is None:
        insert_at = target_sheet.max_row + 1
    else:
        insert_at += 1  # 👈 insert BELOW

    target_sheet.insert_rows(insert_at)

    for col in range(1, source_sheet.max_column + 1):
        src_cell = source_sheet.cell(row=source_row, column=col)
        tgt_cell = target_sheet.cell(row=insert_at, column=col)

        tgt_cell.value = src_cell.value
        tgt_cell.fill = target_fill
        tgt_cell.number_format = src_cell.number_format


def copy_headers(src, tgt):
    for col in range(1, src.max_column + 1):
        tgt.cell(row=1, column=col).value = src.cell(row=1, column=col).value


def copy_row_with_fill(src, tgt, src_row, tgt_row, fill):
    for col in range(1, src.max_column + 1):
        src_cell = src.cell(row=src_row, column=col)
        tgt_cell = tgt.cell(row=tgt_row, column=col)

        tgt_cell.value = src_cell.value
        tgt_cell.number_format = src_cell.number_format  # 🔥 key fix
        tgt_cell.fill = fill