import streamlit as st
from Clockify_Automation.streamlit_app import show_clockify
from fedex_file_formatter.fedex_bill_formatter import show_fedex
from mailbox_compilation_app.app import show_mailbox
from ONHO_Bank_Recognition.app import show_bank_recon
from rta_manifest_automation.app import show_rta_manifest
from xmltoexcel.splittingxml_mon_year_newest import show_xml_converter
from PIL import Image
import base64
from io import BytesIO
from pathlib import Path

# --- 1. CONFIG & REGISTRY ---
st.set_page_config(page_title="Ei1 Portal", layout="wide")

APPS = {
    "clockify": {"title": "Clockify Formatter", "caption": "Timesheet formatting.", "function": show_clockify, "client": "EION", "icon": "🕒"},
    "fedex": {"title": "FedEx Billing", "caption": "FedEx billing IDs.", "function": show_fedex, "client": "ONHO", "icon": "📦"},
    "mailbox": {"title": "Mailbox Data", "caption": "Mailbox data compilation.", "function": show_mailbox, "client": "ONHO", "icon": "📬"},
    "bank": {"title": "Bank Recon", "caption": "Bank matching summary.", "function": show_bank_recon, "client": "ONHO", "icon": "🏦"},
    "rta_manifest": {"title": "RTA Manifest", "caption": "Manifest extractor.", "function": show_rta_manifest, "client": "ONHO", "icon": "📄"},
    "xml_converter": {"title": "XML Converter", "caption": "XML to Excel conversion.", "function": show_xml_converter, "client": "REDH", "icon": "📑"}
}

# --- 2. LOGO HELPER ---
def get_base64_image(image_path):
    try:
        img = Image.open(image_path)
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode()
    except Exception:
        return ""

# --- 3. CUSTOM CSS (No Blue, Centered, Spaced Cards) ---
st.markdown("""
    <style>
        /* 1. KILL ALL BLUE BUTTONS (EXTREME PRIORITY) */
        .main .stButton button, .main .stDownloadButton button {
            background-color: #2c3e50 !important;
            color: #ffffff !important;
            border: 1px solid #2c3e50 !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            text-transform: uppercase !important;
            width: 100% !important;
            height: 45px !important;
        }
        
        .main .stButton button:hover {
            background-color: #1a252f !important;
            color: #f1c40f !important; /* Gold Text */
            border: 1px solid #f1c40f !important;
        }

        /* 2. CARD DESIGN & SPACING */
        .portal-card {
            background-color: #ffffff;
            border-radius: 15px;
            border: 1px solid #e9ecef;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
            display: flex;
            flex-direction: column;
            min-height: 320px;
            margin-bottom: 25px; /* SPACE BETWEEN ROWS */
            overflow: hidden;
        }

        .hero-logo-box {
            background-color: #000000;
            height: 130px;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 15px;
        }

        .card-text-area {
            padding: 20px 15px;
            text-align: center;
            flex-grow: 1;
        }
        
        /* 3. CENTER BUTTON CONTAINER */
        .button-container {
            padding: 0 25px 25px 25px;
        }
    </style>
""", unsafe_allow_html=True)

# --- 4. HEADER HELPER ---
def show_header():
    client = st.session_state.get("selected_client")
    logo_path = "onho-logo.png" if client == "ONHO" else "redh-logo.png" if client == "REDH" else "ei1-logo.png"
    logo_b64 = get_base64_image(logo_path)
    
    st.markdown(f"""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px;">
            <h1 style="color: #2c3e50; margin: 0; font-weight: 800;">Ei1 Portal</h1>
            <div style="background-color: black; padding: 10px; border-radius: 8px;">
                <img src="data:image/png;base64,{logo_b64}" width="130">
            </div>
        </div>
    """, unsafe_allow_html=True)

# --- 5. NAVIGATION ---
if "view" not in st.session_state: st.session_state.view = "client_selection"

if st.session_state.view == "client_selection":
    show_header()
    st.markdown("<h3 style='text-align: center; color: #7f8c8d;'>Select Client</h3>", unsafe_allow_html=True)
    st.write("---")
    
    col1, col2, col3 = st.columns(3)
    clients = [
        {"id": "EION", "logo": "ei1-logo.png", "desc": "Internal Management & Timesheets"},
        {"id": "ONHO", "logo": "onho-logo.png", "desc": "Logistics, Banking & Billing"},
        {"id": "REDH", "logo": "redh-logo.png", "desc": "XML Conversion & File Tools"}
    ]

    for idx, client in enumerate(clients):
        with [col1, col2, col3][idx]:
            img_b64 = get_base64_image(client['logo'])
            st.markdown(f"""
                <div class="portal-card">
                    <div class="hero-logo-box"><img src="data:image/png;base64,{img_b64}" style="max-width: 85%; max-height: 90px; object-fit: contain;"></div>
                    <div class="card-text-area">
                        <h2 style="margin: 0; color: #2c3e50; font-size: 1.5rem;">{client['id']}</h2>
                        <p style="color: #95a5a6; font-size: 18px; margin-top: 10px;">{client['desc']}</p>
                    </div>
                    <div class="button-container">
            """, unsafe_allow_html=True)
            if st.button(f"Enter {client['id']}", key=f"btn_{client['id']}"):
                st.session_state.selected_client = client['id']
                st.session_state.view = "app_list"
                st.rerun()
            st.markdown("</div></div>", unsafe_allow_html=True)

elif st.session_state.view == "app_list":
    show_header()
    if st.button("⬅️ Back to Clients", key="nav_back"):
        st.session_state.selected_client = None
        st.session_state.view = "client_selection"
        st.rerun()
    st.markdown(f"<h2 style='text-align:center;'>{st.session_state.selected_client} Solutions</h2>", unsafe_allow_html=True)
    st.write("---")
    
    client_apps = {k: v for k, v in APPS.items() if v['client'] == st.session_state.selected_client}
    app_keys = list(client_apps.keys())
    
    # GRID LOGIC: Always use 3 columns to maintain size
    for i in range(0, len(app_keys), 3):
        row_keys = app_keys[i:i+3]
        cols = st.columns(3) # This is fixed at 3
        
        for idx, key in enumerate(row_keys):
            info = APPS[key]
            logo_path = "ei1-logo.png" if info['client'] == "EION" else "onho-logo.png" if info['client'] == "ONHO" else "redh-logo.png"
            img_b64 = get_base64_image(logo_path)
            
            with cols[idx]:
                st.markdown(f"""
                    <div class="portal-card">
                        <div class="hero-logo-box"><img src="data:image/png;base64,{img_b64}" style="max-width: 80%; max-height: 80px; object-fit: contain;"></div>
                        <div class="card-text-area">
                            <h3 style="margin:0; color: #2c3e50;">{info['icon']} {info['title']}</h3>
                            <p style="color: #95a5a6; margin-top: 10px; min-height: 40px; font-size: 18px;">{info['caption']}</p>
                        </div>
                        <div class="button-container">
                """, unsafe_allow_html=True)
                if st.button(f"Open {info['title'].split()[0]}", key=f"run_{key}"):
                    st.session_state.selected_app = key
                    st.session_state.view = "app_view"
                    st.rerun()
                st.markdown("</div></div>", unsafe_allow_html=True)

elif st.session_state.view == "app_view":
    if st.button("⬅️ Back to Menu"):
        st.session_state.view = "app_list"
        st.rerun()
    st.divider()
    if st.session_state.selected_app in APPS:
        APPS[st.session_state.selected_app]["function"]()