# ==============================================================================
# 🚀 EI1 INTERNAL AUTOMATIONS PORTAL - DEVELOPER GUIDELINES
# ==============================================================================
#
# DESCRIPTION:
# This portal is a centralized hub for all Ei1 automation tools. 
# It is designed to be DYNAMIC. You do not need to touch the UI code
# to add new tools.
#
# ------------------------------------------------------------------------------
# 📂 PROJECT STRUCTURE
# ------------------------------------------------------------------------------
# /APPS PORTAL
# ├── apps_portal.py          <-- Main Entry Point (The "Brain")
# ├── ei1-logo.png            <-- Company Branding
# ├── requirements.txt        <-- Dependencies for Cloud Deployment
# └── [App_Folders]/         <-- Individual automation tool directories
#
# ------------------------------------------------------------------------------
# 🛠️ HOW TO ADD A NEW APP (3-STEP PROCESS)
# ------------------------------------------------------------------------------
#
# STEP 1: ADD YOUR FOLDER
# -----------------------
# Drop your new automation folder into the root directory. 
# Inside your main script, wrap your  code in a function.
# Example: 
# def show_my_new_tool():
#     st.title("My Tool")
#     ... (your code)
#
# STEP 2: IMPORT THE FUNCTION
# ---------------------------
# Open 'apps_portal.py'. At the top with the other imports, add:
# from your_folder_name.your_script import show_my_new_tool
#
# STEP 3: UPDATE THE APP REGISTRY ("THE BRAIN")
# -------------------------------------------
# Locate the 'APPS' dictionary at the top of 'apps_portal.py'. 
# Add a new key-value pair for your app. The portal will automatically:
#   1. Create a Card on the dashboard.
#   2. Set the Title and Description.
#   3. Handle all navigation and 'Back' button logic.
#
# EXAMPLE ADDITION:
# APPS = {
#     ...
#     "new_tool_key": {
#         "title": "🤖 Tool Name",
#         "caption": "Briefly describe what this tool automates.",
#         "function": show_my_new_tool
#     }
# }
#
# ------------------------------------------------------------------------------
# 🚀 DEPLOYMENT NOTES
# ------------------------------------------------------------------------------
# 1. Push all folders to GitHub.
# 2. Deploy via Streamlit Community Cloud.
# 3. Ensure 'requirements.txt' is updated with any new libraries your 
#    added app requires (e.g., pdfplumber, openpyxl).
#
# ==============================================================================