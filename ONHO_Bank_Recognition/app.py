import streamlit as st
from ONHO_Bank_Recognition.Summary_recon import run_reconciliation

def show_bank_recon():
    st.title("Bank Reconciliation Tool")

    batch_file = st.file_uploader("Upload 1881 Batch File", type=["xlsx"])
    rta_file = st.file_uploader("Upload RTA File", type=["xlsx"])

    if st.button("Run Reconciliation"):
        if batch_file and rta_file:
            # Your original logic call
            output_file = run_reconciliation(batch_file, rta_file)
            st.success("Reconciliation complete!")

            with open(output_file, "rb") as f:
                st.download_button("Download Excel", f, file_name="Reconciliation_Output.xlsx")
        else:
            st.warning("Please upload both files.")