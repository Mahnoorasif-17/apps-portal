from Clockify_Automation.streamlit_app import show_clockify
from fedex_file_formatter.fedex_bill_formatter import show_fedex
from mailbox_compilation_app.app import show_mailbox
from ONHO_Bank_Recognition.app import show_bank_recon
from rta_manifest_automation.app import show_rta_manifest
from xmltoexcel.splittingxml_mon_year_newest import show_xml_converter

# --- CLIENT logos ---
CLIENTS = [
    {"id": "EION", "logo": "logos/ei1-logo.png", "desc": "Internal Management & Timesheets"},
    {"id": "ONHO", "logo": "logos/onho-logo.png", "desc": "Logistics, Banking & Billing"},
    {"id": "REDH", "logo": "logos/redh-logo.png", "desc": "XML Conversion & File Tools"}
]

# --- AUTOMATION APPS ---
APPS = {
    "clockify": {
        "title": "Clockify Formatter", 
        "caption": "Timesheet formatting.", 
        "function": show_clockify, 
        "client": "EION", 
        "logo": "logos/clockify-logo.png"
    },
    "fedex": {
        "title": "FedEx Billing", 
        "caption": "FedEx billing IDs.", 
        "function": show_fedex, 
        "client": "ONHO", 
        "logo": "logos/fedex-logo.png"
    },
    "mailbox": {
        "title": "Mailbox Data", 
        "caption": "Mailbox data compilation.", 
        "function": show_mailbox, 
        "client": "ONHO", 
        "logo": "logos/mailbox-logo.png"
    },
    "bank": {
        "title": "Bank Recon", 
        "caption": "Bank matching summary.", 
        "function": show_bank_recon, 
        "client": "ONHO", 
        "logo": "logos/Bank-logo.png"
    },
    "rta_manifest": {
        "title": "RTA Manifest", 
        "caption": "Manifest extractor.", 
        "function": show_rta_manifest, 
        "client": "ONHO", 
        "logo": "logos/rta-logo.png"
    },
    "xml_converter": {
        "title": "XML Converter", 
        "caption": "XML to Excel conversion.", 
        "function": show_xml_converter, 
        "client": "REDH", 
        "logo": "logos/XML-logo.png"
    }
}