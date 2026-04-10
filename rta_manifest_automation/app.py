import streamlit as st

# st.set_page_config(page_title="RTA - Manifest Automation", layout="wide")

def show_rta_manifest():
    st.title("RTA - Manifest Automation")
    st.markdown("Choose a file type to process:")

    # Navigation menu
    page = st.sidebar.selectbox("Choose processing type", ["Home", "RTA File", "Manifest Files"])

    if page == "Home":
        st.markdown("""
        ### Welcome!
        Use the sidebar to choose:
        - **RTA File**: Upload and process an Excel RTA file
        - **Manifest Files**: Upload one or more Manifest PDFs for data extraction
        """)

    elif page == "RTA File":
        from rta_manifest_automation.rta_page import rta_page
        rta_page()

    elif page == "Manifest Files":
        from rta_manifest_automation.manifest_page import manifest_page
        manifest_page()