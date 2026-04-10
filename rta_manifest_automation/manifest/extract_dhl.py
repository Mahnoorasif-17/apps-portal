import re
import pandas as pd
from rta_manifest_automation.manifest.extract_common import read_pdf_content

def extract_dhl(file):
    """
    Extracts transaction data from a DHL Manifest PDF.

    Args:
        file (UploadedFile or BytesIO): The uploaded PDF file.

    Returns:
        pd.DataFrame: Parsed transaction records.
    """

    def parse_transactions(content_string):
        if not content_string:
            return []

        all_records = []
        transaction_blocks = content_string.split("Summary Totals:")[0].split('Ship To:')[1:]

        for i, block in enumerate(transaction_blocks, 1):
            lines = block.strip().split('\n')
            address = f"{re.search(rf'(.*?)Recorded:.*?', lines[3]).group(1)} "
            address += f"{re.search(rf'(.*?)Picked up:.*?', lines[4]).group(1)} "
            address += f"{re.search(rf'(.*?)Security Services:.*?', lines[5]).group(1)}"
            try:
                last_line = next((line for line in lines if "Package ID" in line), None)
                base_info = {
                    'Recorded': re.search(r'Recorded:\s*(.*?)\s*Over', lines[3]).group(1),
                    'Package ID': re.search(r'Package ID No\.:\s*(\d+)', last_line).group(1) if last_line else None,
                    'Actual Weight': re.search(r'ActualWeight:\s*(.*?lbs)', lines[1]).group(1),
                    "Recipient Name & Comp": f"{re.search(r'(.*?)ActualWeight.*?:', lines[1]).group(1)}, {re.search(r'(.*?)Billable Weight:.*?', lines[2]).group(1)}".strip().strip(","),
                    "Recipient Address": address,
                    'Billable Weight': re.search(r'Billable Weight:\s*(.*?)\s*Add', lines[2]).group(1),
                    'Picked Up': re.search(r'Picked up:\s*([0-9/]+)', lines[4]).group(1),
                    'Tracking No': re.search(r'Tracking No\.:\s*([A-Z0-9]+)', last_line).group(1) if last_line else None,
                    'Service Type': re.search(r'Service Type:\s*(.*?)\s*Service', lines[0]).group(1)
                }
            except (IndexError, AttributeError):
                continue

            charge_types = [
                "Service Charge", "Insured Val.", "Add-on Charges", "Over Sized Piece",
                "Over Weight Piece", "Security Services", "Elevated Risk",
                "Restricted Destination", "Exporter Validation", "Residential Delivery"
            ]

            for line in lines:
                charge_type = next((ct for ct in charge_types if ct in line), None)
                if charge_type:
                    match = re.search(f"({charge_type}.*?)\s*:\s*\$\s*([\d\.]+)", line)
                    if match:
                        amount = match.group(2)
                        if amount != "0.00":
                            record = base_info.copy()
                            record['Charge Type'] = match.group(1)
                            record['Amount'] = amount
                            all_records.append(record)

        return all_records

    content = read_pdf_content(file, header_lines_to_skip=7)
    records = parse_transactions(content)
    return pd.DataFrame(records)
