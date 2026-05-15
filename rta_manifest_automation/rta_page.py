import streamlit as st
import tempfile
import os
import gc
from rta_manifest_automation.processor.pipeline import run_processing_pipeline


def rta_page():
    st.header("📄 RTA Excel File Processor")
    st.info("Supports files up to 1 Year of data")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if not uploaded_file:
        return

    st.success("File uploaded!")
    original_name = os.path.splitext(uploaded_file.name)[0]
    processed_name = f"{original_name} - Processed.xlsx"

    if st.button("Process File"):
        tmp_path = None
        output_path = None
        try:
            # Stream to disk in 8MB chunks — avoids loading whole file in RAM
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                while True:
                    chunk = uploaded_file.read(8 * 1024 * 1024)
                    if not chunk:
                        break
                    tmp.write(chunk)
                tmp_path = tmp.name

            output_path, error_message = run_processing_pipeline(
                tmp_path, return_output_path=True
            )

            if error_message:
                st.error(f"⚠️ {error_message}")
                if output_path and os.path.exists(output_path):
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="Download File (processed until error)",
                            data=f.read(),
                            file_name=processed_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
            else:
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="Download Processed File",
                        data=f.read(),
                        file_name=processed_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
        except Exception as e:
            st.exception(e)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except:
                    pass
            gc.collect()