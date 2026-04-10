import streamlit as st
import tempfile
import os
from rta_manifest_automation.processor.pipeline import run_processing_pipeline


class ValidationError(Exception):
    def __init__(self, message, workbook=None):
        super().__init__(message)
        self.workbook = workbook

st.title("RTA - Manifest Automation")

def rta_page():
    st.header("📄 RTA Excel File Processor")
    uploaded_file   = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file:
        st.success("File uploaded!")

        original_name = os.path.splitext(uploaded_file.name)[0]
        processed_name = f"{original_name} - Processed.xlsx"

        if st.button("Process File"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name

            try:
                output_path, error_message = run_processing_pipeline(
                    tmp_path, return_output_path=True)

                if error_message:
                    st.error(f"⚠️ {error_message}")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="Download File (processed until error)",
                            data=f,
                            file_name=processed_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="Download Processed File",
                            data=f,
                            file_name=processed_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.exception(e)

            os.remove(tmp_path)
            if 'output_path' in locals():
                os.remove(output_path)