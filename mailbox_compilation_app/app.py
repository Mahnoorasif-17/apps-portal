import os
import tempfile

import pandas as pd
import streamlit as st


def show_mailbox():
    def process_excel(input_path, output_path):
        # Read all sheets into a dictionary
        sheets = pd.read_excel(input_path, sheet_name=None, header=None)

        final_df = pd.DataFrame()

        for sheet_index, (sheet_name, df) in enumerate(sheets.items()):
            # ---- SHEET 1 RULES ----
            if sheet_index == 0:
                # Keep first 6 rows + header row
                header_row = df.iloc[6]  # Row 7 = header
                df = df.iloc[7:]  # Data starts from row 8
                df.columns = header_row  # Set header properly

            # ---- SHEET 2+ RULES ----
            else:
                # Ignore first 6 rows + header row, so skip 7 rows
                header_row = sheets[list(sheets.keys())[0]].iloc[6]  # use sheet1 header
                df = df.iloc[7:]  # Skip metadata + header
                df.columns = header_row

            # Remove completely empty rows (a result of skipping)
            df = df.dropna(how="all")

            # # ---- REMOVE the "Total Storage Days all Packages" row and everything after ----
            # drop_index = None
            # for i, row in df.iterrows():
            #     if row.astype(str).str.contains("Total Storage Days all Packages", case=False, na=False).any():
            #         drop_index = i
            #         break
            df = df.iloc[:-8]

            # if drop_index is not None:
            #     df = df.iloc[:drop_index]

            # Append processed sheet data to final
            final_df = pd.concat([final_df, df], ignore_index=True)

            # Add empty row between sheets (EXCEPT last sheet)
            if sheet_index < len(sheets) - 1:
                empty_row = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
                final_df = pd.concat([final_df, empty_row], ignore_index=True)

        final_df.to_excel(output_path, index=False)
        return output_path


    st.title("Excel Sheet Processor")

    uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
            tmp_input.write(uploaded.read())
            tmp_input_path = tmp_input.name

        # Build output filename based on original name
        base_name, ext = os.path.splitext(uploaded.name)
        output_filename = f"{base_name}-processed{ext}"

        output_path = tmp_input_path.replace(".xlsx", "-processed.xlsx")
        process_excel(tmp_input_path, output_path)

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Processed File",
                data=f,
                file_name=output_filename,
            )