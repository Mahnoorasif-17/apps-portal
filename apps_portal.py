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

# --- 1. CONFIG & REGISTRY ---
st.set_page_config(page_title="Ei1 Portal", layout="wide")

# Initialize session state variables so they don't crash
if "selected_client" not in st.session_state: st.session_state.selected_client = None
if "selected_app" not in st.session_state: st.session_state.selected_app = None
if "view" not in st.session_state: st.session_state.view = "client_selection"

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

# --- 3. CUSTOM CSS (Exact UI) ---
st.markdown("""
    <style>
        .card-link { text-decoration: none !important; color: inherit !important; cursor: pointer !important; }
        .portal-card {
            background-color: #ffffff;
            border-radius: 15px;
            border: 1px solid #e9ecef;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
            display: flex;
            flex-direction: column;
            height: 300px;
            overflow: hidden;
            transition: all 0.3s ease;
        }
        .portal-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 12px 20px rgba(0,0,0,0.1);
            border-color: #2c3e50;
        }
        .hero-logo-box {
            background-color: #000000;
            height: 130px;
            display: flex; justify-content: center; align-items: center; padding: 15px;
        }
        .card-text-area { padding: 30px 15px; text-align: center; flex-grow: 1; }
        .card-text-area h2, .card-text-area h3 { color: #2c3e50 !important; font-size: 26px !important; font-weight: 800; }
        .card-text-area p { color: #95a5a6 !important; font-size: 19px !important; }
        
        /* Back Button Design */
        .stButton button {
            background-color: #2c3e50 !important;
            color: white !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- 4. NAVIGATION LOGIC ---
query_params = st.query_params
if "nav" in query_params:
    nav_val = query_params["nav"]
    if nav_val in ["EION", "ONHO", "REDH"]:
        st.session_state.selected_client = nav_val
        st.session_state.view = "app_list"
    elif nav_val in APPS:
        st.session_state.selected_app = nav_val
        st.session_state.view = "app_view"

# --- 5. HEADER ---
def show_header():
    client = st.session_state.selected_client
    logo_path = "onho-logo.png" if client == "ONHO" else "redh-logo.png" if client == "REDH" else "ei1-logo.png"
    logo_b64 = get_base64_image(logo_path)
    st.markdown(f"""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px;">
            <h1 style="color: #2c3e50; margin: 0; font-weight: 800; font-size: 36px;">Ei1 Portal</h1>
            <div style="background-color: black; padding: 10px; border-radius: 8px;">
                <img src="data:image/png;base64,{logo_b64}" width="200">
            </div>
        </div>
    """, unsafe_allow_html=True)

# --- 6. PAGE RENDERING ---
if st.session_state.view == "client_selection":
    show_header()
    st.markdown("<h3 style='text-align: center; color: #7f8c8d; font-size: 28px;'>Select Client</h3>", unsafe_allow_html=True)
    st.write("---")
    cols = st.columns(3, gap="large")
    clients = [
        {"id": "EION", "logo": "ei1-logo.png", "desc": "Internal Management & Timesheets"},
        {"id": "ONHO", "logo": "onho-logo.png", "desc": "Logistics, Banking & Billing"},
        {"id": "REDH", "logo": "redh-logo.png", "desc": "XML Conversion & File Tools"}
    ]
    for idx, client in enumerate(clients):
        with cols[idx]:
            img_b64 = get_base64_image(client['logo'])
            st.markdown(f"""
                <a href="/?nav={client['id']}" target="_self" class="card-link">
                    <div class="portal-card">
                        <div class="hero-logo-box"><img src="data:image/png;base64,{img_b64}" style="max-width: 85%; max-height: 90px; object-fit: contain;"></div>
                        <div class="card-text-area"><h2>{client['id']}</h2><p>{client['desc']}</p></div>
                    </div>
                </a>
            """, unsafe_allow_html=True)

elif st.session_state.view == "app_list":
    show_header()
    if st.button("⬅️ Back to Clients"):
        st.query_params.clear() 
        st.session_state.selected_client = None
        st.session_state.view = "client_selection"
        st.rerun()
    
    st.markdown(f"<h2 style='text-align:center; font-size: 34px;'>{st.session_state.selected_client} Solutions</h2>", unsafe_allow_html=True)
    st.write("---")
    
    client_apps = {k: v for k, v in APPS.items() if v['client'] == st.session_state.selected_client}
    app_keys = list(client_apps.keys())
    
    # GRID LOGIC
    for i in range(0, len(app_keys), 3):
        row_keys = app_keys[i:i+3]
        cols = st.columns(3, gap="large")
        
        for idx, key in enumerate(row_keys):
            info = APPS[key]
            l_path = "ei1-logo.png" if info['client'] == "EION" else "onho-logo.png" if info['client'] == "ONHO" else "redh-logo.png"
            img_b64 = get_base64_image(l_path)
            with cols[idx]:
                st.markdown(f"""
                    <a href="/?nav={key}" target="_self" class="card-link">
                        <div class="portal-card">
                            <div class="hero-logo-box"><img src="data:image/png;base64,{img_b64}" style="max-width: 80%; max-height: 80px; object-fit: contain;"></div>
                            <div class="card-text-area">
                                <h3>{info['icon']} {info['title']}</h3>
                                <p>{info['caption']}</p>
                            </div>
                        </div>
                    </a>
                """, unsafe_allow_html=True)
        
        # --- THE FIX: FORCE A GAP AFTER EVERY ROW ---
        st.markdown('<div style="margin-bottom: 60px;"></div>', unsafe_allow_html=True)

elif st.session_state.view == "app_view":
    # FIXED BACK BUTTON LOGIC
    if st.button("⬅️ Back to Menu"):
        # Make sure we know which client to go back to
        client_to_return = st.session_state.selected_client
        if not client_to_return and st.session_state.selected_app in APPS:
            client_to_return = APPS[st.session_state.selected_app]["client"]
        
        st.query_params.update({"nav": client_to_return})
        st.session_state.view = "app_list"
        st.rerun()
        
    st.divider()
    if st.session_state.selected_app in APPS:
        APPS[st.session_state.selected_app]["function"]()