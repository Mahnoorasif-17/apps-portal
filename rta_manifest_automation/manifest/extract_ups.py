import re
import pandas as pd
from rta_manifest_automation.manifest.extract_common import read_pdf_content

def extract_ups(file):
    """
    Extracts transaction data from a UPS Manifest PDF (streamed file).

    Args:
        file (UploadedFile or BytesIO): The uploaded PDF file.

    Returns:
        pd.DataFrame: Parsed transaction records.
    """

    def parse_transactions(content_string):
        if not content_string:
            return []

        all_records = []
        transaction_blocks = content_string.split('Ship To:')[1:]

        for i, block in enumerate(transaction_blocks, 1):
            lines = block.strip().split('\n')
            try:
                base_info = {
                    'Service Type': lines[0].split('UPS Total Charge:')[0][14:].strip(),
                    'Shipment ID': lines[2].split('Shipment ID:')[1].strip(),
                    'Recipient Name': lines[2].split('Shipment ID:')[0].strip(),
                    'Recipient Company':  lines[1].split('Total Packages:')[0].strip(),
                    'Recipient Address': lines[3].split('Billable Weight:')[0].strip() + lines[4].split('Billing Option:')[0].strip(),
                    'Billable Weight': lines[3].split('Billable Weight:')[1].strip(),
                    'Recorded': lines[8].split('Recorded:')[1].strip().split(' ')[0],
                    'Picked Up': lines[9].split(" ")[2],
                    'Actual Weight': re.search(r'Actual Weight:\s*(.*?lbs)', lines[7]).group(1),
                    'Package ID': re.search(r'Package ID\.:\s*(\d+)', lines[8]).group(1),
                    'Tracking No': re.search(r'Tracking No\.:\s*([A-Z0-9]+)', lines[5]).group(1)
                }
            except (IndexError, AttributeError):
                continue

            charge_lines_to_check = lines[5:9]
            for line in charge_lines_to_check:
                match = re.search(r'(?:Tracking No\.:|Package Type:|Actual Weight:|Recorded:)\s*.*?([A-Za-z\s\(\)\$\d\.,/]+:)\s*\$\s*([\d\.]+)$', line)
                if match:
                    charge_type_list = match.group(1).strip().split(' ')
                    charge_type_list.pop(0)
                    charge_type = ' '.join(charge_type_list).strip(":").strip("lbs ")
                    amount = match.group(2)
                    record = base_info.copy()
                    record['Charge Type'] = charge_type
                    record['Amount'] = amount
                    all_records.append(record)

        return all_records
    
    pdf_content = read_pdf_content(file, header_lines_to_skip=7)
    records = parse_transactions(pdf_content)
    df = pd.DataFrame(records)
    return df