import re
import time
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font


def clean_text(val):
    if val is None:
        return ""
    return re.sub(r"[^a-zA-Z0-9 ]", " ", str(val)).lower().strip()


def parse_amount(val):
    if val is None or val == "":
        return 0.0
    val = str(val).replace("$", "").replace(",", "").strip()
    try:
        return float(val)
    except:
        return 0.0


def process_step_6(workbook):
    t0 = time.time()

    step5 = workbook["Step 5"]

    if "Step 6" in workbook.sheetnames:
        del workbook["Step 6"]

    step6 = workbook.create_sheet("Step 6")

    headers = [
        "UID", "RegID", "Date", "Time", "Item",
        "Tender", "Customer", "Amount", "Tax", "Total Amount", "Taxable"
    ]
    step6.append(headers)

    item_col   = 5
    amount_col = 8

    TAX_RATE = 0.08875

    # Excludes shipping services (was previously skipped via fill color)
    EXCLUDE_KEYWORDS = [
        "coupon", "discount", "void", "term",
        "late fee", "mailbox", "setup fee", "renew",
        "ups", "usps", "fedex", "dhl",           # shipping services
        "declared value",                         # paired with shipping
    ]

    # Special-case items: bypass EXCLUDE_KEYWORDS and go into Step 6 / Retail
    # (e.g., DHL DROP OFF is retail, not a shipping service)
    INCLUDE_OVERRIDES = [
        "dhl drop off",
    ]

    TAXABLE_KEYWORDS = [
        "copies", "misc  taxable", "fax", "lamination",
        "passport", "postcard",
        "printing", "scan", "office rental"
    ]

    # Keywords that identify a coupon/discount row
    COUPON_DISCOUNT_KEYWORDS = ["coupon", "discount"]

    def is_no_fill_from_cell(cell):
        fill = cell.fill
        if fill is None:
            return True
        if fill.fill_type is None:
            return True
        if getattr(fill, "patternType", None) is None:
            return True
        return False

    # Load Step 5 into memory once
    t = time.time()
    rows_values = []
    rows_no_fill = []
    max_col = step5.max_column
    for row_cells in step5.iter_rows(min_row=1, max_row=step5.max_row, max_col=max_col):
        rows_values.append([c.value for c in row_cells])
        if len(row_cells) >= item_col:
            rows_no_fill.append(is_no_fill_from_cell(row_cells[item_col - 1]))
        else:
            rows_no_fill.append(True)
    print(f"  [Step6] loaded {len(rows_values)} rows in {time.time()-t:.2f}s")

    processed_rows = []
    total_amount = 0.0
    total_tax    = 0.0
    total_total  = 0.0

    # Tracks whether the last accepted (non-coupon/discount) item was a retail item.
    # If True, any immediately following coupon/discount rows are included.
    last_item_was_retail = False

    t = time.time()
    for r_idx in range(1, len(rows_values)):
        row_vals = rows_values[r_idx]

        item_raw   = row_vals[item_col - 1] if len(row_vals) >= item_col else None
        item_clean = clean_text(item_raw)
        amount     = parse_amount(row_vals[amount_col - 1] if len(row_vals) >= amount_col else None)

        # Skip completely empty rows
        if item_clean == "" and amount == 0:
            continue

        # Skip colored rows (purple/green/blue from Step 5)
        if not rows_no_fill[r_idx]:
            continue

        # Detect if this row is a coupon/discount row
        is_coupon_discount = any(k in item_clean for k in COUPON_DISCOUNT_KEYWORDS)

        if is_coupon_discount:
            # Only include if the last accepted non-coupon item was a retail item
            if last_item_was_retail:
                row_data = []
                for c in range(1, 11):
                    if c <= len(row_vals):
                        row_data.append(row_vals[c - 1])
                    else:
                        row_data.append(None)

                row_data[7] = amount   # coupon amounts are already negative
                row_data[8] = 0.0     # no tax on coupons/discounts
                row_data[9] = amount  # total = amount (negative)

                processed_rows.append((item_clean, row_data, "n"))

                total_amount += amount
                total_tax    += 0.0
                total_total  += amount
            # Either way, do NOT update last_item_was_retail — keep it as-is
            # so multiple consecutive coupons after a retail item are all included
            continue

        # --- Non-coupon/discount row from here ---

        # Check INCLUDE_OVERRIDES first — these bypass EXCLUDE_KEYWORDS
        is_override = any(ov in item_clean for ov in INCLUDE_OVERRIDES)

        # Skip excluded keywords (unless it's an override item like DHL DROP OFF)
        if not is_override and any(k in item_clean for k in EXCLUDE_KEYWORDS):
            last_item_was_retail = False  # excluded item resets the flag
            continue

        # Safety: skip Mechanical Totals if any slipped through
        if "mechanical total" in item_clean:
            last_item_was_retail = False
            continue

        # Tax
        is_taxable   = any(k in item_clean for k in TAXABLE_KEYWORDS)
        tax          = round(amount * TAX_RATE, 2) if is_taxable else 0.0
        taxable_flag = "y" if is_taxable else "n"
        total        = round(amount + tax, 2)

        row_data = []
        for c in range(1, 11):
            if c <= len(row_vals):
                row_data.append(row_vals[c - 1])
            else:
                row_data.append(None)

        row_data[7] = amount
        row_data[8] = tax
        row_data[9] = total

        processed_rows.append((item_clean, row_data, taxable_flag))

        total_amount += amount
        total_tax    += tax
        total_total  += total

        # This was a retail item — coupons immediately following it should be included
        last_item_was_retail = True

    print(f"  [Step6] processed {len(processed_rows)} rows in {time.time()-t:.2f}s")

    processed_rows.sort(key=lambda x: x[0])

    for _, row_data, taxable_flag in processed_rows:
        step6.append(row_data + [taxable_flag])

    for row in range(2, step6.max_row + 1):
        for col in [8, 9, 10]:
            cell = step6.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'

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

    for col in range(1, 12):
        step6.column_dimensions[get_column_letter(col)].width = 20

    step6.freeze_panes    = "A2"
    step6.auto_filter.ref = f"A1:K{step6.max_row}"

    _build_retail_tab(workbook, step6, headers, processed_rows, total_amount, total_tax, total_total)

    print(f"  [Step6] TOTAL: {time.time()-t0:.2f}s")
    return step6


def _build_retail_tab(workbook, step6, headers, processed_rows, total_amount, total_tax, total_total):
    TAB_NAME = "Retail"

    if TAB_NAME in workbook.sheetnames:
        del workbook[TAB_NAME]

    retail = workbook.create_sheet(TAB_NAME)
    retail.append(headers)

    for _, row_data, taxable_flag in processed_rows:
        retail.append(row_data + [taxable_flag])

    for row in range(2, retail.max_row + 1):
        for col in [8, 9, 10]:
            cell = retail.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'

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

    for col in range(1, 12):
        retail.column_dimensions[get_column_letter(col)].width = 20

    retail.freeze_panes    = "A2"
    retail.auto_filter.ref = f"A1:K{retail.max_row}"