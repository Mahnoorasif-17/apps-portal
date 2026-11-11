import re
import time
import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill


DUP_YELLOW = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
VOID_PURPLE = PatternFill(start_color="B44CB5", end_color="B44CB5", fill_type="solid")  # NEW (same purple as Step 5)


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


def clean_date(val):
    if val is None:
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val  # real date/datetime object, formatting handles display
    s = str(val).strip()
    # strip trailing time like " 0:00:00" or " 00:00:00"
    s = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?$", "", s)
    return s


def process_step_6(workbook):
    t0 = time.time()

    step5 = workbook["Step 5"]

    if "Step 6" in workbook.sheetnames:
        del workbook["Step 6"]

    step6 = workbook.create_sheet("Step 6")

    # New headers include 3 NS columns at the end
    headers = [
        "UID", "RegID", "Date", "Time", "Item",
        "Tender", "Customer", "Amount", "Tax", "Total Amount", "Taxable",
        "Item NetSuite ID", "Item NetSuite Name", "Customer NetSuite ID"
    ]
    step6.append(headers)

    item_col   = 5
    amount_col = 8
    regid_col  = 2   # NEW

    TAX_RATE = 0.08875

    EXCLUDE_KEYWORDS = [
        "coupon", "discount", "void", "term",
        "late fee", "mailbox", "setup fee", "renew",
        "ups", "usps", "fedex", "dhl",
        "declared value",
    ]

    INCLUDE_OVERRIDES = ["dhl drop off"]

    TAXABLE_KEYWORDS = [
        "copies", "misc  taxable", "fax", "lamination",
        "passport", "postcard",
        "printing", "scan", "office rental"
    ]

    COUPON_DISCOUNT_KEYWORDS = ["coupon", "discount"]

    # NEW: account customers (same keywords as Step 4 build_account_sheets); their Cash rows are copied to Extra
    ACCOUNT_CUSTOMER_KEYWORDS = ["e-scriber", "escriber", "empire", "feshaire", "fashaire"]
    tender_col   = 6
    customer_col = 7

    def is_no_fill_from_cell(cell):
        fill = cell.fill
        if fill is None:
            return True
        if fill.fill_type is None:
            return True
        if getattr(fill, "patternType", None) is None:
            return True
        return False

    # --- Find NS column positions in Step 5 by header name ---
    ns_item_id_src = ns_item_name_src = ns_cust_id_src = None
    for col in range(1, step5.max_column + 1):
        h = str(step5.cell(row=1, column=col).value or "").strip().lower()
        if h == "item netsuite id":
            ns_item_id_src = col
        elif h == "item netsuite name":
            ns_item_name_src = col
        elif h == "customer netsuite id":
            ns_cust_id_src = col

    # --- Load Step 5 into memory (values + fill info + NS cust ID cell fills) ---
    t = time.time()
    rows_values = []
    rows_no_fill = []
    rows_ns_cust_yellow = []  # True if the source NS Customer ID cell was yellow (duplicate)
    max_col = step5.max_column

    for row_cells in step5.iter_rows(min_row=1, max_row=step5.max_row, max_col=max_col):
        rows_values.append([c.value for c in row_cells])
        if len(row_cells) >= item_col:
            rows_no_fill.append(is_no_fill_from_cell(row_cells[item_col - 1]))
        else:
            rows_no_fill.append(True)

        # Check if the Customer NetSuite ID cell in this row was yellow
        is_yellow = False
        if ns_cust_id_src is not None and len(row_cells) >= ns_cust_id_src:
            cell = row_cells[ns_cust_id_src - 1]
            if cell.fill and cell.fill.fill_type == "solid":
                color = cell.fill.start_color.rgb if cell.fill.start_color else None
                if color and "FFFF00" in str(color).upper():
                    is_yellow = True
        rows_ns_cust_yellow.append(is_yellow)

    print(f"  [Step6] loaded {len(rows_values)} rows in {time.time()-t:.2f}s")

    # NEW: coupon/discount rows that cancel another row with the same RegID and amount -> "Extra" sheet
    extra_idx = set()
    coupons_by_regid = {}
    items_by_regid = {}
    for i in range(1, len(rows_values)):
        if not rows_no_fill[i]:
            continue
        rv = rows_values[i]
        rid = rv[regid_col - 1] if len(rv) >= regid_col else None
        if rid in (None, ""):
            continue
        ic = clean_text(rv[item_col - 1] if len(rv) >= item_col else None)
        amt = parse_amount(rv[amount_col - 1] if len(rv) >= amount_col else None)
        if any(k in ic for k in COUPON_DISCOUNT_KEYWORDS):
            coupons_by_regid.setdefault(rid, []).append((i, amt))
        elif ic and amt != 0 and "petty" not in ic and "mechanical total" not in ic:
            items_by_regid.setdefault(rid, []).append((i, amt))
    for rid, coupon_list in coupons_by_regid.items():
        available = list(items_by_regid.get(rid, []))
        for c_idx, c_amt in coupon_list:
            if c_amt == 0:
                continue
            for k, (i_idx, i_amt) in enumerate(available):
                if round(c_amt + i_amt, 2) == 0:  # same amount, coupon is the negative side
                    extra_idx.update((c_idx, i_idx))
                    del available[k]
                    break

    processed_rows = []
    total_amount = 0.0
    total_tax    = 0.0
    total_total  = 0.0

    # NEW: rows going to the Extra sheet and Petty rows going to the Void sheet
    extra_rows = []
    petty_rows = []

    last_item_was_retail = False

    t = time.time()
    for r_idx in range(1, len(rows_values)):
        row_vals = rows_values[r_idx]

        item_raw   = row_vals[item_col - 1] if len(row_vals) >= item_col else None
        item_clean = clean_text(item_raw)
        amount     = parse_amount(row_vals[amount_col - 1] if len(row_vals) >= amount_col else None)

        if item_clean == "" and amount == 0:
            continue

        if not rows_no_fill[r_idx]:
            continue

        is_coupon_discount = any(k in item_clean for k in COUPON_DISCOUNT_KEYWORDS)

        # Extract NS values for this source row
        ns_item_id   = row_vals[ns_item_id_src - 1]   if ns_item_id_src   is not None and len(row_vals) >= ns_item_id_src   else None
        ns_item_name = row_vals[ns_item_name_src - 1] if ns_item_name_src is not None and len(row_vals) >= ns_item_name_src else None
        ns_cust_id   = row_vals[ns_cust_id_src - 1]   if ns_cust_id_src   is not None and len(row_vals) >= ns_cust_id_src   else None
        cust_is_yellow = rows_ns_cust_yellow[r_idx]

        # NEW: Petty -> Void sheet (not part of Step 6 / Retail)
        if "petty" in item_clean:
            petty_rows.append((row_vals, ns_item_id, ns_item_name, ns_cust_id, cust_is_yellow))
            continue

        if is_coupon_discount:
            if last_item_was_retail:
                row_data = []
                for c in range(1, 11):
                    if c <= len(row_vals):
                        row_data.append(row_vals[c - 1])
                    else:
                        row_data.append(None)

                row_data[2] = clean_date(row_data[2])

                row_data[7] = amount
                row_data[8] = 0.0
                row_data[9] = amount

                # Append taxable_flag + NS values
                full_row = row_data + ["n", ns_item_id, ns_item_name, ns_cust_id]

                if r_idx in extra_idx:  # NEW: also copy to Extra (row stays in Step 6 / Retail)
                    extra_rows.append((item_clean, full_row, cust_is_yellow))

                processed_rows.append((item_clean, full_row, cust_is_yellow))

                total_amount += amount
                total_total  += amount
            continue

        is_override = any(ov in item_clean for ov in INCLUDE_OVERRIDES)

        if not is_override and any(k in item_clean for k in EXCLUDE_KEYWORDS):
            last_item_was_retail = False
            continue

        if "mechanical total" in item_clean:
            last_item_was_retail = False
            continue

        is_taxable   = any(k in item_clean for k in TAXABLE_KEYWORDS)
        # NEW: Rent / Rental in Item -> not taxable (n)
        if re.search(r"\brent", item_clean):
            is_taxable = False
        tax          = round(amount * TAX_RATE, 2) if is_taxable else 0.0
        taxable_flag = "y" if is_taxable else "n"
        total        = round(amount + tax, 2)

        row_data = []
        for c in range(1, 11):
            if c <= len(row_vals):
                row_data.append(row_vals[c - 1])
            else:
                row_data.append(None)

        row_data[2] = clean_date(row_data[2])

        row_data[7] = amount
        row_data[8] = tax
        row_data[9] = total

        full_row = row_data + [taxable_flag, ns_item_id, ns_item_name, ns_cust_id]

        if r_idx in extra_idx:  # NEW: also copy to Extra (row stays in Step 6 / Retail)
            extra_rows.append((item_clean, full_row, cust_is_yellow))

        processed_rows.append((item_clean, full_row, cust_is_yellow))

        total_amount += amount
        total_tax    += tax
        total_total  += total

        last_item_was_retail = True

    print(f"  [Step6] processed {len(processed_rows)} rows in {time.time()-t:.2f}s")

    # NEW: Cash rows of E-Scribers / Empire / Feshaire -> Extra (Tax / Total / Taxable left blank)
    already_in_extra = {fr[0] for _, fr, _ in extra_rows if fr[0]}
    for r_idx in range(1, len(rows_values)):
        row_vals = rows_values[r_idx]
        if len(row_vals) < max(amount_col, customer_col, tender_col):
            continue

        if str(row_vals[tender_col - 1] or "").strip().lower() != "cash":
            continue
        customer_val = str(row_vals[customer_col - 1] or "").lower()
        if not any(kw in customer_val for kw in ACCOUNT_CUSTOMER_KEYWORDS):
            continue

        item_raw = row_vals[item_col - 1]
        amount   = parse_amount(row_vals[amount_col - 1])
        if str(item_raw or "").strip() == "" and amount == 0:
            continue
        if "void" in str(item_raw or "").lower():
            continue
        if row_vals[0] and row_vals[0] in already_in_extra:
            continue

        ns_item_id   = row_vals[ns_item_id_src - 1]   if ns_item_id_src   is not None and len(row_vals) >= ns_item_id_src   else None
        ns_item_name = row_vals[ns_item_name_src - 1] if ns_item_name_src is not None and len(row_vals) >= ns_item_name_src else None
        ns_cust_id   = row_vals[ns_cust_id_src - 1]   if ns_cust_id_src   is not None and len(row_vals) >= ns_cust_id_src   else None

        full_row = [row_vals[0], row_vals[1], clean_date(row_vals[2]), row_vals[3], item_raw,
                    row_vals[5], row_vals[6], amount, None, None, None,
                    ns_item_id, ns_item_name, ns_cust_id]
        extra_rows.append((clean_text(item_raw), full_row, rows_ns_cust_yellow[r_idx]))

    print(f"  [Step6] UPDATED VERSION -> extra rows: {len(extra_rows)}, petty rows: {len(petty_rows)}")  # NEW (diagnostic)

    processed_rows.sort(key=lambda x: x[0])

    # Column indices in Step 6 output
    COL_AMOUNT_OUT = 8
    COL_TAX_OUT    = 9
    COL_TOTAL_OUT  = 10
    COL_TAXABLE    = 11
    COL_NS_CUST_ID_OUT = 14  # Customer NetSuite ID

    for _, full_row, cust_is_yellow in processed_rows:
        step6.append(full_row)
        # Highlight NS Customer ID cell yellow if the source was flagged
        if cust_is_yellow:
            step6.cell(row=step6.max_row, column=COL_NS_CUST_ID_OUT).fill = DUP_YELLOW

    for row in range(2, step6.max_row + 1):
        for col in [COL_AMOUNT_OUT, COL_TAX_OUT, COL_TOTAL_OUT]:
            cell = step6.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'
        date_cell = step6.cell(row=row, column=3)
        if isinstance(date_cell.value, (datetime.datetime, datetime.date)):
            date_cell.number_format = 'mm/dd/yyyy'

    step6.append([])
    step6.append([
        "", "", "", "", "", "", "TOTALS",
        total_amount,
        total_tax,
        total_total,
        "", "", "", ""
    ])

    last_row = step6.max_row
    for col in [COL_AMOUNT_OUT, COL_TAX_OUT, COL_TOTAL_OUT]:
        step6.cell(row=last_row, column=col).number_format = '$#,##0.00'
        step6.cell(row=last_row, column=col).font          = Font(bold=True)
    step6.cell(row=last_row, column=7).font = Font(bold=True)

    for col in range(1, 15):
        step6.column_dimensions[get_column_letter(col)].width = 20

    step6.freeze_panes    = "A2"
    step6.auto_filter.ref = f"A1:N{step6.max_row}"

    _build_retail_tab(workbook, step6, headers, processed_rows,
                      total_amount, total_tax, total_total)

    # NEW (Extra is built last so it sits at the very end, after Retail)
    _move_petty_to_void(workbook, petty_rows)
    _build_extra_tab(workbook, headers, extra_rows)

    print(f"  [Step6] TOTAL: {time.time()-t0:.2f}s")
    return step6


def _build_retail_tab(workbook, step6, headers, processed_rows,
                      total_amount, total_tax, total_total):
    TAB_NAME = "Retail"

    if TAB_NAME in workbook.sheetnames:
        del workbook[TAB_NAME]

    retail = workbook.create_sheet(TAB_NAME)
    retail.append(headers)

    COL_AMOUNT_OUT = 8
    COL_TAX_OUT    = 9
    COL_TOTAL_OUT  = 10
    COL_NS_CUST_ID_OUT = 14

    for _, full_row, cust_is_yellow in processed_rows:
        retail.append(full_row)
        if cust_is_yellow:
            retail.cell(row=retail.max_row, column=COL_NS_CUST_ID_OUT).fill = DUP_YELLOW

    for row in range(2, retail.max_row + 1):
        for col in [COL_AMOUNT_OUT, COL_TAX_OUT, COL_TOTAL_OUT]:
            cell = retail.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'
        date_cell = retail.cell(row=row, column=3)
        if isinstance(date_cell.value, (datetime.datetime, datetime.date)):
            date_cell.number_format = 'mm/dd/yyyy'

    retail.append([])
    retail.append([
        "", "", "", "", "", "", "TOTALS",
        total_amount,
        total_tax,
        total_total,
        "", "", "", ""
    ])

    last_row = retail.max_row
    for col in [COL_AMOUNT_OUT, COL_TAX_OUT, COL_TOTAL_OUT]:
        retail.cell(row=last_row, column=col).number_format = '$#,##0.00'
        retail.cell(row=last_row, column=col).font          = Font(bold=True)
    retail.cell(row=last_row, column=7).font = Font(bold=True)

    for col in range(1, 15):
        retail.column_dimensions[get_column_letter(col)].width = 20

    retail.freeze_panes    = "A2"
    retail.auto_filter.ref = f"A1:N{retail.max_row}"


# NEW: "Extra" sheet (same 14 columns as Step 6) for coupon/discount + item pairs with same RegID and amount
def _build_extra_tab(workbook, headers, extra_rows):
    TAB_NAME = "Extra"

    if TAB_NAME in workbook.sheetnames:
        del workbook[TAB_NAME]

    ws = workbook.create_sheet(TAB_NAME)  # created last -> lands after Retail
    ws.append(headers)

    COL_AMOUNT_OUT = 8
    COL_TAX_OUT    = 9
    COL_TOTAL_OUT  = 10
    COL_NS_CUST_ID_OUT = 14

    for _, full_row, cust_is_yellow in extra_rows:
        ws.append(full_row)
        if cust_is_yellow:
            ws.cell(row=ws.max_row, column=COL_NS_CUST_ID_OUT).fill = DUP_YELLOW

    for row in range(2, ws.max_row + 1):
        for col in [COL_AMOUNT_OUT, COL_TAX_OUT, COL_TOTAL_OUT]:
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '$#,##0.00'
        date_cell = ws.cell(row=row, column=3)
        if isinstance(date_cell.value, (datetime.datetime, datetime.date)):
            date_cell.number_format = 'mm/dd/yyyy'

    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = 20

    ws.freeze_panes    = "A2"
    ws.auto_filter.ref = f"A1:N{ws.max_row}"


# NEW: appends Petty rows to the Void sheet built in Step 5 (same 11-column layout, purple)
def _move_petty_to_void(workbook, petty_rows):
    if not petty_rows:
        return

    TAB_NAME = "Void-Discount-Coupons"
    void_headers = ["UID", "RegID", "Date", "Time", "Item", "Tender", "Customer", "Amount",
                    "Item NetSuite ID", "Item NetSuite Name", "Customer NetSuite ID"]

    if TAB_NAME in workbook.sheetnames:
        ws = workbook[TAB_NAME]
    else:
        ws = workbook.create_sheet(TAB_NAME)
        for col_idx, h in enumerate(void_headers, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = h
            cell.fill = VOID_PURPLE
            cell.font = Font(bold=True, color="FFFFFF")
        ws.freeze_panes = "A2"

    # Skip UIDs already on the sheet so re-running Step 6 doesn't duplicate rows
    existing_uids = set()
    for (v,) in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        if v:
            existing_uids.add(v)

    for row_vals, ns_item_id, ns_item_name, ns_cust_id, cust_is_yellow in petty_rows:
        def g(i):
            return row_vals[i] if len(row_vals) > i else None

        uid = g(0)
        if uid and uid in existing_uids:
            continue

        r = ws.max_row + 1
        values = [uid, g(1), clean_date(g(2)), g(3), g(4), g(5), g(6), g(7),
                  ns_item_id, ns_item_name, ns_cust_id]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c)
            cell.value = v
            cell.fill = VOID_PURPLE

        ws.cell(row=r, column=1).number_format = '0'
        ws.cell(row=r, column=3).number_format = 'mm/dd/yyyy'
        amt_cell = ws.cell(row=r, column=8)
        amt_cell.number_format = '$#,##0.00'
        if isinstance(g(7), (int, float)) and g(7) < 0:
            amt_cell.font = Font(color="FF0000")
        if cust_is_yellow:
            ws.cell(row=r, column=11).fill = DUP_YELLOW