import streamlit as st
import streamlit.components.v1 as components
import os
import pandas as pd
from lxml import etree
import xml.etree.ElementTree as ET

# Import your existing logic from the uploaded file
# Ensure tableau_comparator.py is in the same directory
from tableau_comparator import (
    sign_in, 
    get_workbook_id_in_project, 
    download_latest_workbook_revision,
    extract_sections,
    parse_twb,
    # Add other necessary imports from your file here:
    # generate_html_report, xmldiff_text, etc.
)

st.set_page_config(page_title="Tableau Workbook Comparator", layout="wide")

st.title("📊 Tableau Workbook Comparator")
st.sidebar.header("Tableau Connectivity")

# --- 1. Connection Section ---
with st.sidebar:
    server_url = st.text_input("Server URL", value="https://prod-useast-b.online.tableau.com")
    site_id = st.text_input("Site ID (Content URL)", value="")
    token_name = st.text_input("Token Name")
    token_secret = st.text_input("Token Secret", type="password")
    
    if st.button("Connect to Tableau"):
        try:
            token, site_id_resp = sign_in(server_url, site_id, token_name, token_secret)
            st.session_state['tableau_token'] = token
            st.session_state['tableau_site_id'] = site_id_resp
            st.success("Connected!")
        except Exception as e:
            st.error(f"Connection failed: {e}")

# --- 2. Selection Section ---
if 'tableau_token' in st.session_state:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Source Selection")
        src_project = st.text_input("Source Project Name", key="src_proj")
        src_workbook = st.text_input("Source Workbook Name", key="src_wb")
        
    with col2:
        st.subheader("Target Selection")
        tgt_project = st.text_input("Target Project Name", key="tgt_proj")
        tgt_workbook = st.text_input("Target Workbook Name", key="tgt_wb")

    # --- 3. Comparison Execution ---
    if st.button("🚀 Compare Workbooks", use_container_width=True):
        with st.spinner("Downloading and analyzing workbooks..."):
            try:
                token = st.session_state['tableau_token']
                sid = st.session_state['tableau_site_id']

                # Download Source
                src_data = download_latest_workbook_revision(token, sid, src_project, src_workbook)
                # Download Target
                tgt_data = download_latest_workbook_revision(token, sid, tgt_project, tgt_workbook)

                # Parsing and Comparison Logic (Simplified for UI flow)
                root_old = parse_twb(src_data['twb_path'])
                root_new = parse_twb(tgt_data['twb_path'])
                
                # ... (Insert the rest of your comparison logic here) ...
                # Use your existing 'generate_html_report' function
                
                report_path = "compare_report.html"
                # Assuming your generate_html_report saves to a file
                # generate_html_report(...) 

                # --- 4. Display Result ---
                if os.path.exists(report_path):
                    with open(report_path, 'r', encoding='utf-8') as f:
                        html_content = f.read()
                    
                    st.success("Comparison Complete!")
                    components.html(html_content, height=800, scrolling=True)
                
            except Exception as e:
                st.error(f"Error during comparison: {str(e)}")
else:
    st.info("Please connect to Tableau using the sidebar to begin.")