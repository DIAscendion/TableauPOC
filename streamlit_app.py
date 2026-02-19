import streamlit as st
import os
import xml.etree.ElementTree as ET
# Import your existing functions from your script
# (Assuming your script is named tableau_comparator.py)
from tableau_comparator import sign_in, download_latest_workbook_revision, extract_sections

st.set_page_config(page_title="Tableau Workbook Comparator", layout="wide")

# --- SIDEBAR: Server Details ---
with st.sidebar:
    st.title("🔐 Tableau Server Details")
    server_url = st.text_input("Tableau Site URL", placeholder="https://prod-useast-a.online.tableau.com")
    site_id = st.text_input("Site Content URL", placeholder="YourSiteName")
    token_name = st.text_input("PAT Name")
    token_secret = st.text_input("PAT Secret", type="password")
    
    st.divider()
    st.title("📊 Workbooks to Compare")
    proj_a = st.text_input("Source Project")
    wb_a = st.text_input("Source Workbook")
    
    st.divider()
    proj_b = st.text_input("Target Project")
    wb_b = st.text_input("Target Workbook")

    run_button = st.button("🚀 Compare Workbooks", use_container_width=True)

# --- MAIN UI ---
st.title("🔍 Tableau Workbook Comparator")
st.info("This tool compares two Tableau workbooks and identifies changes in calculations, filters, and layout.")

if run_button:
    if not all([server_url, site_id, token_name, token_secret]):
        st.error("Please provide all Tableau server details in the sidebar.")
    else:
        with st.spinner("Signing in and downloading workbooks..."):
            try:
                # 1. Sign In (Injecting UI values into your existing logic)
                # You'll need to modify your sign_in function to accept parameters
                token, s_id = sign_in(server_url, site_id, token_name, token_secret)
                
                # 2. Download 
                data_a = download_latest_workbook_revision(token, s_id, proj_a, wb_a)
                data_b = download_latest_workbook_revision(token, s_id, proj_b, wb_b)
                
                # 3. Parse and Compare
                root_a = ET.parse(data_a['twb_path']).getroot()
                root_b = ET.parse(data_b['twb_path']).getroot()
                
                # 4. Display Results in Tabs
                tab1, tab2, tab3 = st.tabs(["📊 Visual Change Tree", "📄 XML Structural Diff", "👥 Permissions"])
                
                with tab1:
                    st.subheader("High-Level Summary")
                    # Call your 'render_visual_change_tree' here
                    # st.write(your_change_logic_results)
                    
                with tab2:
                    st.subheader("Raw XML Differences")
                    # Display the xmldiff output
                    
                st.success("Comparison Complete!")
                
            except Exception as e:
                st.error(f"Error: {e}")
