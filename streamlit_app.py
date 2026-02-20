import streamlit as st
import streamlit.components.v1 as components
import os
import xml.etree.ElementTree as ET
import requests

# Import your existing logic from the uploaded file
import tableau_comparator as tc 

st.set_page_config(page_title="Tableau Workbook Comparator", layout="wide")

st.title("📊 Tableau Workbook Comparator")

# --- 1. Connection Section (Sidebar) ---
with st.sidebar:
    st.header("🔐 Connection")
    server_url = st.text_input("Server URL", value="https://prod-useast-b.online.tableau.com")
    site_id_input = st.text_input("Site Content URL (ID)", help="The part of the URL after /site/")
    token_name = st.text_input("PAT Name")
    token_secret = st.text_input("PAT Secret", type="password")
    
    if st.button("Connect to Tableau"):
        try:
            # Sign in and store session details
            token, site_id_resp = tc.sign_in(server_url, site_id_input, token_name, token_secret)
            
            # CRITICAL FIX: Update the global variable in your module so other functions work
            tc.TABLEAU_SITE_URL = server_url.rstrip('/')
            
            st.session_state['tableau_token'] = token
            st.session_state['tableau_site_id'] = site_id_resp
            st.session_state['server_url'] = server_url.rstrip('/')
            st.success("Connected!")
        except Exception as e:
            st.error(f"Connection failed: {e}")

# Helper to fetch projects for the dropdown
def get_projects(token, site_id, server_url):
    url = f"{server_url}/api/3.25/sites/{site_id}/projects"
    headers = {"X-Tableau-Auth": token}
    r = requests.get(url, headers=headers, verify=False)
    r.raise_for_status()
    root = ET.fromstring(r.text)
    ns = {"t": "http://tableau.com/api"}
    return {p.attrib['name']: p.attrib['id'] for p in root.findall(".//t:project", ns)}

# Helper to fetch workbooks for a project
def get_workbooks_in_project(token, site_id, server_url, project_id):
    url = f"{server_url}/api/3.25/sites/{site_id}/workbooks?filter=projectId:eq:{project_id}"
    headers = {"X-Tableau-Auth": token}
    r = requests.get(url, headers=headers, verify=False)
    r.raise_for_status()
    root = ET.fromstring(r.text)
    ns = {"t": "http://tableau.com/api"}
    return [wb.attrib['name'] for wb in root.findall(".//t:workbook", ns)]

# --- 2. Dynamic Selection Section ---
if 'tableau_token' in st.session_state:
    token = st.session_state['tableau_token']
    sid = st.session_state['tableau_site_id']
    srv = st.session_state['server_url']

    # Refresh project list
    try:
        project_map = get_projects(token, sid, srv)
        project_names = sorted(list(project_map.keys()))
    except Exception as e:
        st.error(f"Could not fetch projects: {e}")
        project_names = []

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📘 Source Selection")
        src_proj = st.selectbox("Source Project", project_names, key="src_p_sel")
        if src_proj:
            src_wb_list = get_workbooks_in_project(token, sid, srv, project_map[src_proj])
            src_wb = st.selectbox("Source Workbook", sorted(src_wb_list), key="src_w_sel")
        
    with col2:
        st.subheader("📗 Target Selection")
        tgt_proj = st.selectbox("Target Project", project_names, key="tgt_p_sel")
        if tgt_proj:
            tgt_wb_list = get_workbooks_in_project(token, sid, srv, project_map[tgt_proj])
            tgt_wb = st.selectbox("Target Workbook", sorted(tgt_wb_list), key="tgt_w_sel")

    # --- 3. Comparison Execution ---
    if st.button("🚀 Run Comparison", use_container_width=True):
        if src_wb and tgt_wb:
            with st.spinner("Analyzing differences..."):
                try:
                    # Sync the internal URL variable again just in case
                    tc.TABLEAU_SITE_URL = srv
                    
                    # Call the existing logic from your script
                    src_data = tc.download_latest_workbook_revision(token, sid, src_proj, src_wb)
                    tgt_data = tc.download_latest_workbook_revision(token, sid, tgt_proj, tgt_wb)

                    # Note: You need to call your report generation function here 
                    # based on the download results above.
                    
                    report_file = "compare_SOURCE_vs_TARGET_latest.html"
                    if os.path.exists(report_file):
                        with open(report_file, 'r', encoding='utf-8') as f:
                            html_content = f.read()
                        st.success("Comparison Successful!")
                        components.html(html_content, height=1000, scrolling=True)
                except Exception as e:
                    st.error(f"Error during comparison: {e}")
        else:
            st.warning("Please select both a source and target workbook.")
else:
    st.info("Please connect to Tableau via the sidebar to begin.")