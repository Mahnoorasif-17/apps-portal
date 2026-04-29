import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Alignment

# --- CORE LOGIC ---
def process_excel(input_file):
    # Load the CSV
    df = pd.read_csv(input_file, dtype={'VoicePhoneNo': str, 'VoicePhoneNo2': str})

    repl_map = {
        'street': 'ST', 'st': 'ST', 'avenue': 'AVE', 'ave': 'AVE',
        'east': 'E', 'west': 'W', 'north': 'N', 'south': 'S',
        '1st': '1ST', '2nd': '2ND', '3rd': '3RD',
        'apartment': 'APT', 'apt': 'APT', 'drive': 'DR', 
        'room': 'RM', 'floor': 'FL', 'panthouse': 'PH', 
        'suite': 'STE', 'road': 'RD'
    }

    def apply_formatting(text):
        if pd.isna(text) or text == "": return ""
        text = str(text).strip()
        for word, sub in repl_map.items():
            text = re.sub(rf'\b{word}\b', sub, text, flags=re.IGNORECASE)
        text = re.sub(rf'(\d+)(st|nd|rd|th)', lambda m: m.group(1) + m.group(2).upper(), text, flags=re.IGNORECASE)
        return text

    split_keywords = ['ST', 'AVE', 'RD', 'DR', 'PL', 'BLVD', 'LN', 'ROAD', 'STREET', 'AVENUE', 'APT', 'STE', 'FRNT', 'OFC']
    
    def split_address_final(row):
        addr1 = str(row['Address1']).strip() if pd.notna(row['Address1']) else ""
        addr2 = str(row['Address2']).strip() if pd.notna(row['Address2']) else ""
        if addr2.lower() == 'nan': addr2 = ""
        split_pattern = rf'\b({"|".join(split_keywords)})\b|#'
        match = re.search(split_pattern, addr1, flags=re.IGNORECASE)
        if match:
            split_idx = match.end()
            new_addr1 = addr1[:split_idx].strip().rstrip(',')
            new_addr2 = addr1[split_idx:].strip()
            dir_match = re.match(r'^([NSEW])\b', new_addr2, flags=re.IGNORECASE)
            if dir_match:
                direction = dir_match.group(1).upper()
                new_addr1 = f"{new_addr1} {direction}"
                new_addr2 = new_addr2[len(direction):].strip()
            if addr2:
                new_addr2 = (new_addr2 + " " + addr2).strip()
            new_addr2 = new_addr2.replace("# ", "#")
            return new_addr1, new_addr2
        return addr1, addr2

    addr_splits = df.apply(split_address_final, axis=1, result_type='expand')
    df['Address1_Mod'] = addr_splits[0].apply(apply_formatting)
    df['Address2_Mod'] = addr_splits[1].astype(str).str.upper()

    def clean_date(val):
        if pd.isna(val) or val == "": return ""
        try:
            d = pd.to_datetime(val)
            return f"{d.month}/{d.day}/{d.year}"
        except: return val

    output_df = pd.DataFrame(index=df.index)
    output_df['Source'] = 'postalmate'
    
    def calculate_rta(row):
        fn = str(row['FirstName']).strip() if pd.notna(row['FirstName']) and str(row['FirstName']) != 'nan' else ""
        ln = str(row['LastName']).strip() if pd.notna(row['LastName']) and str(row['LastName']) != 'nan' else ""
        cn = str(row['CompanyName']).strip() if pd.notna(row['CompanyName']) and str(row['CompanyName']) != 'nan' else ""
        name_concat = fn + ln
        return pd.Series([cn if cn else name_concat, name_concat if name_concat else cn])

    rta_cols = df.apply(calculate_rta, axis=1)
    output_df['Original RTA String'] = rta_cols[0]
    output_df['Updated RTA String'] = rta_cols[1]
    output_df['CustomerID'] = pd.to_numeric(df['CustomerID'], errors='coerce')
    output_df['AddDate'] = df['AddDate'].apply(clean_date)
    output_df['NamePre'] = df['NamePre']
    output_df['FirstName'] = df['FirstName']
    output_df['LastName'] = df['LastName']
    output_df['CompanyName'] = df['CompanyName']
    output_df['Address1 (Unmodified)'] = df['Address1']
    output_df['Address2 (Unmodified)'] = df['Address2']
    output_df['Address3 (Unmodified)'] = df['Address3']
    output_df['City (Unmodified)'] = df['City']
    output_df['Address1 (Modified)'] = df['Address1_Mod']
    output_df['Address2 (Modified)'] = df['Address2_Mod'].replace('NAN', '')
    output_df['Address3 (Modified)'] = df['Address3']
    output_df['City (Modified)'] = df['City'].apply(apply_formatting)
    output_df['City (Modified)'] = output_df['City (Modified)'].str.replace('Manhattan', 'NEW YORK', case=False, regex=True)
    output_df['StateDisplay'] = df['StateDisplay']
    output_df['ZipDisplay'] = df['ZipDisplay']
    output_df['Zip4'] = df['Zip4']
    output_df['CountryName'] = df['CountryName'].replace(['USA', 'US'], 'United States')
    
    def phone_to_num(val):
        clean = re.sub(r'[\s\-\(\)]+', '', str(val))
        if clean == 'nan' or clean == "": return ""
        try: return int(float(clean))
        except: return clean

    output_df['VoicePhoneNo'] = df['VoicePhoneNo'].apply(phone_to_num)
    output_df['VoicePhoneNo2'] = df['VoicePhoneNo2'].apply(phone_to_num)
    output_df['Email'] = df['Email']
    output_df['LastShipDTG'] = df['LastShipDTG'].apply(clean_date)
    output_df['LastActivityDTG'] = df['LastActivityDTG'].apply(clean_date)
    output_df['Note'] = df['Note']

    # --- EXCEL FORMATTING IN MEMORY ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        for column_cells in ws.columns:
            max_length = max(len(str(cell.value)) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = max_length + 3
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='right')
        ws['O1'].alignment = Alignment(horizontal='left')
    
    return output.getvalue()

# --- WRAPPER FUNCTION FOR PORTAL ---
def show_customer_formatter():
    # CSS to make file uploader labels look original and big
    st.markdown("""
        <style>
            .stFileUploader label p {
                font-size: 22px !important;
                font-weight: 600 !important;
                color: #2c3e50 !important;
            }
        </style>
    """, unsafe_allow_html=True)

    st.title("📊 ONHO Customer Data Formatter")
    st.write("Upload your **ONHO Customer Report.csv** file to format it automatically.")

    uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

    if uploaded_file:
        st.success("File Uploaded Successfully!")
        if st.button("Process and Format"):
            with st.spinner("Processing..."):
                try:
                    processed_data = process_excel(uploaded_file)
                    st.download_button(
                        label="📥 Download Formatted Excel",
                        data=processed_data,
                        file_name="ONHO_Customer_Formatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"An error occurred: {e}")