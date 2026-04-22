import streamlit as st
from PIL import Image
import base64
from io import BytesIO
# Import the data from our registry
from registry import CLIENTS, APPS

# --- 1. CONFIG ---
st.set_page_config(page_title="Ei1 Portal", layout="wide")

if "selected_client" not in st.session_state: st.session_state.selected_client = None
if "selected_app" not in st.session_state: st.session_state.selected_app = None
if "view" not in st.session_state: st.session_state.view = "client_selection"

# --- 2. LOGO HELPER ---
def get_base64_image(image_path):
    import os
    try:
        # Check if file exists to avoid silent errors
        if not os.path.exists(image_path):
            print(f"DEBUG: File not found at {image_path}")
            return ""
            
        img = Image.open(image_path)
        buffered = BytesIO()
        
        # Convert to RGB if it's a JPG to avoid transparency issues during base64 encoding
        if image_path.lower().endswith(('jpg', 'jpeg')):
            img = img.convert('RGB')
            img.save(buffered, format="JPEG")
        else:
            img.save(buffered, format="PNG")
            
        return base64.b64encode(buffered.getvalue()).decode()
    except Exception as e:
        print(f"DEBUG: Error loading {image_path}: {e}")
        return ""

# --- 3. CUSTOM CSS ---
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
            height: 340px; 
            overflow: hidden;
            transition: all 0.3s ease;
        }
        
        .portal-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 12px 20px rgba(0,0,0,0.1);
            border-color: #2c3e50;
        }
        
        /* THE LOGO BOX - Made much bigger */
        .hero-logo-box {
            background-color: #000000;
            height: 160px;
            display: flex; 
            justify-content: center; 
            align-items: center; 
            padding: 10px; /* Reduced padding to let image grow */
        }
        
        /* THE IMAGE - Forced to fill space */
        .hero-logo-box img {
            max-width: 95% !important;  /* Allowed to take more width */
            max-height: 95% !important; /* Allowed to take more height */
            object-fit: contain !important;
            display: block;
        }
        
        .card-text-area { padding: 20px 15px; text-align: center; flex-grow: 1; }
        .card-text-area h2, .card-text-area h3 { color: #2c3e50 !important; font-size: 26px !important; font-weight: 800; margin-bottom: 5px;}
        .card-text-area p { color: #95a5a6 !important; font-size: 19px !important; }
        
        /* White Back Buttons */
        .stButton button {
            background-color: #ffffff !important;
            color: #2c3e50 !important;
            border: 2px solid #2c3e50 !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            padding: 8px 20px !important;
        }
        
        .stButton button:hover {
            background-color: #2c3e50 !important;
            color: #ffffff !important;
        }
            
        h1 {
            font-size: 50px !important; /* Adjust this number to go even bigger */
            font-weight: 900 !important;
            color: #2c3e50 !important;
            margin-bottom: 30px !important;
            text-transform: uppercase !important;
            letter-spacing: -1px !important;
        }

     
        .header-logo {
            width: 200px !important; 
        }
    </style>
""", unsafe_allow_html=True)

# --- 4. NAVIGATION LOGIC ---
query_params = st.query_params
if "nav" in query_params:
    nav_val = query_params["nav"]
    if any(c['id'] == nav_val for c in CLIENTS):
        st.session_state.selected_client = nav_val
        st.session_state.view = "app_list"
    elif nav_val in APPS:
        st.session_state.selected_app = nav_val
        st.session_state.view = "app_view"

# --- 5. HEADER ---
def show_header():
    client_id = st.session_state.selected_client
    client_info = next((c for c in CLIENTS if c['id'] == client_id), None)
    logo_path = client_info['logo'] if client_info else "logos/ei1-logo.png"
    logo_b64 = get_base64_image(logo_path)
    
    st.markdown(f"""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 40px;">
            <h1>Ei1 Portal</h1>
            <div style="background-color: black; padding: 15px; border-radius: 12px;">
                <img src="data:image/png;base64,{logo_b64}" class="header-logo" width="150">
            </div>
        </div>
    """, unsafe_allow_html=True)

# --- 6. PAGE RENDERING ---
if st.session_state.view == "client_selection":
    show_header()
    st.markdown("<h3 style='text-align: center; color: #7f8c8d; font-size: 38px;'>Select Client</h2>", unsafe_allow_html=True)
    st.write("---")
    cols = st.columns(3, gap="large")
    for idx, client in enumerate(CLIENTS):
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
    
    for i in range(0, len(app_keys), 3):
        row_keys = app_keys[i:i+3]
        cols = st.columns(3, gap="large")
        for idx, key in enumerate(row_keys):
            info = APPS[key]
            # Calls the specific automation logo (Bank, FedEx, etc.)
            img_b64 = get_base64_image(info['logo'])
            with cols[idx]:
                st.markdown(f"""
                    <a href="/?nav={key}" target="_self" class="card-link">
                        <div class="portal-card">
                            <div class="hero-logo-box"><img src="data:image/png;base64,{img_b64}" style="max-width: 80%; max-height: 80px; object-fit: contain;"></div>
                            <div class="card-text-area">
                                <h3>{info['title']}</h3>
                                <p>{info['caption']}</p>
                            </div>
                        </div>
                    </a>
                """, unsafe_allow_html=True)
        st.markdown('<div style="margin-bottom: 60px;"></div>', unsafe_allow_html=True)

elif st.session_state.view == "app_view":
    if st.button("⬅️ Back to Menu"):
        client_to_return = st.session_state.selected_client
        if not client_to_return and st.session_state.selected_app in APPS:
            client_to_return = APPS[st.session_state.selected_app]["client"]
        st.query_params.update({"nav": client_to_return})
        st.session_state.view = "app_list"
        st.rerun()
    st.divider()
    if st.session_state.selected_app in APPS:
        APPS[st.session_state.selected_app]["function"]()