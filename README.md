# 🛠️ Ei1 Developer Guidelines

```python
# ==========================================================
# EI1 INTERNAL AUTOMATIONS PORTAL
# ==========================================================
#
# DESCRIPTION:
# A centralized, dynamic hub for all Ei1 automation tools.
#
# HOW TO ADD A NEW APP:
# ----------------------------------------------------------
# # STEP 1: PREPARE THE LOGO
# - Save your automation logo in the '/logos' directory.
# - IMPORTANT: Use a transparent PNG so it blends with the black header.
# - Note the exact filename (e.g., 'my-tool-logo.png').
#
# STEP 2: IMPORT THE FUNCTION
# - Open 'registry.py'.
# - Import your main function at the top of the file.
#   Example: from my_folder.app import show_my_tool
#
# STEP 3: REGISTER IN THE APPS DICTIONARY
# - Go to the 'APPS' dictionary in 'registry.py'.
# - Add your app details using the following structure:
#
#   "my_tool_key": {
#       "title": "My New Tool",
#       "caption": "Short description of what it does.",
#       "function": show_my_tool,
#       "client": "ONHO",  # Options: 'EION', 'ONHO', or 'REDH'
#       "logo": "logos/my-tool-logo.png"
#   },
#
# STEP 4: ADD A NEW CLIENT (IF NEEDED)
# - If you are adding a new client portfolio, update the 'CLIENTS' list:
#
#   {
#       "id": "NEW_CLIENT",
#       "logo": "logos/client-main-logo.png",
#       "desc": "Portfolio description."
#   }
#
# ----------------------------------------------------------
# No changes are needed in 'apps_portal.py' logic! 
# Everything is managed through the 'registry.py' file.
# ----------------------------------------------------------