import pandas as pd
import streamlit as st
from rta_manifest_automation.manifest.extract_dhl import extract_dhl
from rta_manifest_automation.manifest.extract_fedex import extract_fedex
from rta_manifest_automation.manifest.extract_ups import extract_ups

def manifest_page():
    st.header("Manifest PDF File Processor")

    uploaded_files = st.file_uploader(
        "Upload Manifest PDF(s)", type=["pdf"], accept_multiple_files=True
    )

    if uploaded_files:
        st.info(f"{len(uploaded_files)} file(s) uploaded.")
        extracted_data = []
        outname = ""
        for file in uploaded_files:
            filename = file.name.lower()

            if "ups" in filename:
                outname += "UPS"
                df = extract_ups(file)
            elif "fedex" in filename:
                outname += "FEDEX"
                df = extract_fedex(file)
            elif "dhl" in filename:
                outname += "DHL"
                df = extract_dhl(file)
            else:
                st.warning(f"Unknown courier for file: {file.name}")
                continue

            if isinstance(df, pd.DataFrame):
                extracted_data.append(df)

        if extracted_data:
            full_df = pd.concat(extracted_data, ignore_index=True)
            st.success("Extraction complete!")
            st.dataframe(full_df)
            csv = full_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download Extracted Data",
                csv,
                f"{outname} manifest extracted data.csv",
                "text/csv",
            )
