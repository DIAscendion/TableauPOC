import streamlit as st
import streamlit.components.v1 as components
import os
import xml.etree.ElementTree as ET
import requests
from datetime import datetime

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
# --- HELPER FUNCTIONS ---
def get_projects(token, site_id, server_url):
    # Ensure URL is clean and headers are standard for Tableau Online
    url = f"{server_url.rstrip('/')}/api/3.25/sites/{site_id}/projects?pageSize=1000"
    headers = {
        "X-Tableau-Auth": token,
        "Accept": "application/xml" # Force XML to match your ET.fromstring logic
    }
    
    try:
        r = requests.get(url, headers=headers, verify=False, timeout=30)
        
        # Check if we got a valid response before parsing
        if r.status_code != 200:
            st.error(f"Tableau API Error ({r.status_code}): {r.text[:200]}")
            return {}
            
        if not r.text.strip():
            st.error("Tableau returned an empty response for projects.")
            return {}

        root = ET.fromstring(r.content) # Use .content (bytes) for better encoding handling
        ns = {"t": "http://tableau.com/api"}
        
        projects = {p.attrib['name']: p.attrib['id'] for p in root.findall(".//t:project", ns)}
        return projects

    except ET.ParseError as e:
        st.error("Failed to parse Tableau response. The server might be sending HTML instead of XML.")
        with st.expander("Show raw response for debugging"):
            st.code(r.text)
        return {}
    except Exception as e:
        st.error(f"Unexpected error fetching projects: {e}")
        return {}

def get_workbooks_in_project(token, site_id, server_url, project_id):
    if not project_id:
        return []
    
    # We fetch ALL workbooks for the site and filter locally to avoid API filter bugs
    url = f"{server_url.rstrip('/')}/api/3.25/sites/{site_id}/workbooks?pageSize=1000"
    headers = {
        "X-Tableau-Auth": token,
        "Accept": "application/xml"
    }
    
    try:
        r = requests.get(url, headers=headers, verify=False, timeout=30)
        if r.status_code != 200:
            return []
            
        root = ET.fromstring(r.content)
        ns = {"t": "http://tableau.com/api"}
        
        # Find all workbooks where the parent project ID matches our selection
        workbooks = []
        for wb in root.findall(".//t:workbook", ns):
            # Look for the project tag inside the workbook tag
            proj_tag = wb.find("t:project", ns)
            if proj_tag is not None and proj_tag.attrib.get('id') == project_id:
                workbooks.append(wb.attrib.get('name'))
        
        return workbooks
    except Exception as e:
        st.error(f"Error fetching workbooks: {e}")
        return []

# --- MAIN UI ---
if 'tableau_token' in st.session_state:
    token = st.session_state['tableau_token']
    sid = st.session_state['tableau_site_id']
    srv = st.session_state['server_url']

    # Load projects into session state to avoid repeated API calls
    if 'project_map' not in st.session_state:
        st.session_state.project_map = get_projects(token, sid, srv)
    
    project_names = sorted(list(st.session_state.project_map.keys()))

    if not project_names:
        st.error("No projects found. Please check your user permissions in Tableau.")
        st.stop()

    # Layout for Source and Target
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📘 Source Selection")
        src_proj = st.selectbox("Select Project", project_names, key="src_p_sel")
        src_proj_id = st.session_state.project_map.get(src_proj)
        
        # Dynamic Workbook Dropdown
        src_wb_list = get_workbooks_in_project(token, sid, srv, src_proj_id)
        src_wb = st.selectbox("Select Workbook", sorted(src_wb_list) if src_wb_list else ["No workbooks found"], key="src_w_sel")
        
    with col2:
        st.subheader("📗 Target Selection")
        # Reuse the same project names
        tgt_proj = st.selectbox("Select Project", project_names, key="tgt_p_sel")
        tgt_proj_id = st.session_state.project_map.get(tgt_proj)
        
        tgt_wb_list = get_workbooks_in_project(token, sid, srv, tgt_proj_id)
        tgt_wb = st.selectbox("Select Workbook", sorted(tgt_wb_list) if tgt_wb_list else ["No workbooks found"], key="tgt_w_sel")

    # --- 3. Comparison Execution ---
    if st.button("🚀 Run Comparison", use_container_width=True, type="primary"):
        if src_wb and tgt_wb and src_wb != "No workbooks found":
            with st.spinner("🕵️ Analyzing XML Heuristics & Permissions..."):
                try:
                    # 1. Environment Setup
                    tc.TABLEAU_SITE_URL = srv
                    report_file = "compare_SOURCE_vs_TARGET_latest.html"
                    
                    # 2. Download and Parse
                    src_data = tc.download_latest_workbook_revision(token, sid, src_proj, src_wb)
                    tgt_data = tc.download_latest_workbook_revision(token, sid, tgt_proj, tgt_wb)
                    
                    root_old = tc.parse_twb(src_data['twb_path'])
                    root_new = tc.parse_twb(tgt_data['twb_path'])

                    # 3. Fetch Missing Metadata (Emails & Permissions)
                    src_owner = tc.get_workbook_owner(token, sid, src_data.get('workbook_id'))
                    tgt_owner = tc.get_workbook_owner(token, sid, tgt_data.get('workbook_id'))
                    
                    src_perms = tc.get_users_and_permissions_for_workbook(token, sid, src_proj, src_wb)
                    tgt_perms = tc.get_users_and_permissions_for_workbook(token, sid, tgt_proj, tgt_wb)

                    # 4. Core Heuristics
                    sec_old = tc.extract_sections(root_old)
                    sec_new = tc.extract_sections(root_new)
                    
                    # Reset Global States inside the tool
                    tc.CHANGE_REGISTRY = {
                        "workbook": [], "datasources": {}, "calculations": {}, 
                        "parameters": {}, "worksheets": {}, "dashboards": {}, "stories": {}
                    }
                    # IMPORTANT: Clear the internal datasource summaries before building cards
                    if hasattr(tc, 'DATASOURCE_CARDS'): tc.DATASOURCE_CARDS = [] 

                    # 5. Build Content
                    cards = tc.build_cards(sec_old, sec_new)
                    tc.populate_change_registry_from_cards(cards)
                    
                    # Manual addition of Summary Card
                    overall_summary = tc.build_overall_workbook_summary_card(sec_old, sec_new, cards, root_old, root_new)
                    if overall_summary:
                        cards.insert(0, overall_summary)

                    # 6. Generate HTML Components
                    kpi_old = tc.build_workbook_kpi_snapshot(sec_old)
                    kpi_new = tc.build_workbook_kpi_snapshot(sec_new)
                    kpi_html = tc.render_workbook_kpi_table(kpi_old, kpi_new)
                    
                    visual_tree = tc.render_visual_change_tree(sec_new, tc.CHANGE_REGISTRY, tgt_wb)
                    
                    perm_html = (
                        tc.build_users_permissions_card_with_context(src_proj, src_wb, src_perms, context="source")
                        + "<hr/>" +
                        tc.build_users_permissions_card_with_context(tgt_proj, tgt_wb, tgt_perms, context="Target")
                    )

                    # 7. Call generate_html_report with EXACT positional arguments (15 total)
                    # This matches your def generate_html_report(...) signature exactly
                    tc.generate_html_report(
                        src_wb,              # title_a
                        tgt_wb,              # title_b
                        cards,               # cards
                        None,                # structural_ops
                        report_file,         # out_file
                        kpi_html,            # kpi_html
                        root_new,            # root_new
                        visual_tree,         # visual_tree_text
                        tgt_owner,           # latest_publisher
                        "Latest",            # latest_revision
                        datetime.now().strftime("%Y-%m-%d"), # latest_published_at
                        src_owner,           # old_publisher
                        tgt_owner,           # new_publisher
                        perm_html,           # users_permissions_html
                        ""                   # site_users_html
                    )

                    # 8. Render in Streamlit
                    if os.path.exists(report_file):
                        with open(report_file, 'r', encoding='utf-8') as f:
                            html_content = f.read()
                        st.success("✅ Comparison Complete")
                        components.html(html_content, height=1200, scrolling=True)
                        st.download_button("📥 Download Full HTML", html_content, 
                                         file_name=f"Diff_{tgt_wb}.html", mime="text/html")
                
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)
        else:
            st.warning("Please select both a source and target workbook.")
else:
    st.info("Please connect to Tableau via the sidebar to begin.")