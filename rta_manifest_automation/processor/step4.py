from openpyxl.styles import PatternFill
from .utils import *

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
    remove_columns_by_header(step4, ["SubTotal", "Tax", "Total"])
    drop_rows_with_empty_item(step4)
    remove_footer_and_mech_rows(step4)
    clear_all_highlighting(step4)

    format_header(step4, header_row=1)
    highlight_header_row(step4, header_row=1)
    autofit_columns(step4)
    distribute_items_to_sheets(step4, workbook)


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
                            insert_row_below_regid(
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
                            insert_row_below_regid(
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
