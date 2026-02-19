import streamlit as st
import os
import xml.etree.ElementTree as ET
import requests
import urllib3
from pathlib import Path

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

st.set_page_config(page_title="Tableau Workbook Comparator", layout="wide")

# ============================================================================
# TABLEAU API HELPERS (parameterized versions)
# ============================================================================

@st.cache_data
def sign_in_with_params(site_url, site_content_url, token_name, token_secret):
    """Sign in to Tableau using provided credentials."""
    try:
        url = f"{site_url}/api/3.25/auth/signin"
        payload = {"credentials": {
            "personalAccessTokenName": token_name,
            "personalAccessTokenSecret": token_secret,
            "site": {"contentUrl": site_content_url}
        }}
        r = requests.post(url, json=payload, verify=VERIFY_SSL, timeout=60)
        r.raise_for_status()
        root = ET.fromstring(r.text)
        ns = {"t": "http://tableau.com/api"}
        token = root.find(".//t:credentials", ns).attrib["token"]
        site_id = root.find(".//t:site", ns).attrib["id"]
        return token, site_id
    except Exception as e:
        st.error(f"Sign-in failed: {str(e)}")
        return None, None

def list_projects(site_url, site_id, token):
    """List all projects."""
    try:
        url = f"{site_url}/api/3.25/sites/{site_id}/projects"
        r = requests.get(url, headers={"X-Tableau-Auth": token}, verify=VERIFY_SSL, timeout=60)
        r.raise_for_status()
        root = ET.fromstring(r.text)
        ns = {"t": "http://tableau.com/api"}
        projects = []
        for proj in root.findall(".//t:project", ns):
            projects.append({
                "id": proj.attrib.get("id"),
                "name": proj.attrib.get("name")
            })
        return projects
    except Exception as e:
        st.error(f"Failed to list projects: {str(e)}")
        return []

def list_workbooks_in_project(site_url, site_id, token, project_id):
    """List all workbooks in a project."""
    try:
        url = f"{site_url}/api/3.25/sites/{site_id}/projects/{project_id}/workbooks"
        r = requests.get(url, headers={"X-Tableau-Auth": token}, verify=VERIFY_SSL, timeout=60)
        r.raise_for_status()
        root = ET.fromstring(r.text)
        ns = {"t": "http://tableau.com/api"}
        workbooks = []
        for wb in root.findall(".//t:workbook", ns):
            workbooks.append({
                "id": wb.attrib.get("id"),
                "name": wb.attrib.get("name")
            })
        return workbooks
    except Exception as e:
        st.error(f"Failed to list workbooks: {str(e)}")
        return []

def get_workbook_revisions(site_url, site_id, token, workbook_id):
    """Get revision history for a workbook."""
    try:
        url = f"{site_url}/api/3.25/sites/{site_id}/workbooks/{workbook_id}/revisions"
        r = requests.get(url, headers={"X-Tableau-Auth": token}, verify=VERIFY_SSL, timeout=60)
        r.raise_for_status()
        root = ET.fromstring(r.text)
        ns = {"t": "http://tableau.com/api"}
        revisions = []
        for rev in root.findall(".//t:revision", ns):
            pub_elem = rev.find("t:publisher", ns)
            pub_name = pub_elem.attrib.get("name") if pub_elem is not None else "Unknown"
            revisions.append({
                "number": rev.attrib.get("revisionNumber"),
                "publishedAt": rev.attrib.get("publishedAt"),
                "publisher": pub_name
            })
        return sorted(revisions, key=lambda x: x["publishedAt"], reverse=True)
    except Exception as e:
        st.error(f"Failed to get revisions: {str(e)}")
        return []

# ============================================================================
# UI INITIALIZATION & STATE MANAGEMENT
# ============================================================================

if "auth_token" not in st.session_state:
    st.session_state.auth_token = None
    st.session_state.site_id = None
    st.session_state.authenticated = False
    st.session_state.projects = []
    st.session_state.workbooks = {}

# ============================================================================
# MAIN UI
# ============================================================================

st.title("🔍 Tableau Workbook Comparator")
st.markdown("Compare two Tableau workbooks and identify changes in calculations, filters, layout, permissions, and more.")

# --- SIDEBAR: Tableau Credentials ---
with st.sidebar:
    st.header("🔐 Tableau Credentials")
    
    site_url = st.text_input(
        "Tableau Site URL",
        placeholder="https://prod-useast-a.online.tableau.com",
        help="Full URL of your Tableau Server/Online site"
    )
    
    site_content_url = st.text_input(
        "Site Content URL (ID)",
        placeholder="yoursitenamehere",
        help="The site ID/URL slug (without https://)"
    )
    
    token_name = st.text_input(
        "PAT Token Name",
        placeholder="Enter your Personal Access Token name",
        help="Tableau personal access token username"
    )
    
    token_secret = st.text_input(
        "PAT Token Secret",
        type="password",
        placeholder="Enter your Personal Access Token secret",
        help="Tableau personal access token secret/password"
    )
    
    if st.button("🔓 Connect to Tableau", use_container_width=True):
        if not all([site_url, site_content_url, token_name, token_secret]):
            st.error("❌ Please fill in all credential fields.")
        else:
            with st.spinner("Connecting to Tableau..."):
                token, site_id = sign_in_with_params(site_url, site_content_url, token_name, token_secret)
                if token and site_id:
                    st.session_state.auth_token = token
                    st.session_state.site_id = site_id
                    st.session_state.site_url = site_url
                    st.session_state.authenticated = True
                    st.success("✅ Successfully connected to Tableau!")
                    st.rerun()

# ============================================================================
# MAIN CONTENT (only show if authenticated)
# ============================================================================

if st.session_state.authenticated:
    
    # Create two columns for Source and Target workbooks
    col1, col2 = st.columns(2)
    
    # ========== SOURCE WORKBOOK (LEFT COLUMN) ==========
    with col1:
        st.subheader("📘 Source Workbook")
        
        # Project Selection
        projects = list_projects(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token)
        project_names = [p["name"] for p in projects]
        
        if not projects:
            st.warning("No projects available. Check your access permissions.")
        else:
            selected_project_source = st.selectbox(
                "Select Source Project",
                options=project_names,
                key="source_project"
            )
            
            # Get project ID
            source_project_id = next((p["id"] for p in projects if p["name"] == selected_project_source), None)
            
            if source_project_id:
                # Workbook Selection
                workbooks_source = list_workbooks_in_project(
                    st.session_state.site_url,
                    st.session_state.site_id,
                    st.session_state.auth_token,
                    source_project_id
                )
                
                workbook_names_source = [wb["name"] for wb in workbooks_source]
                
                if workbooks_source:
                    selected_workbook_source = st.selectbox(
                        "Select Source Workbook",
                        options=workbook_names_source,
                        key="source_workbook"
                    )
                    
                    # Get workbook ID
                    source_workbook_id = next((wb["id"] for wb in workbooks_source if wb["name"] == selected_workbook_source), None)
                    
                    if source_workbook_id:
                        # Revision Selection
                        st.write("**Select Revision:**")
                        revisions_source = get_workbook_revisions(
                            st.session_state.site_url,
                            st.session_state.site_id,
                            st.session_state.auth_token,
                            source_workbook_id
                        )
                        
                        if revisions_source:
                            revision_options_source = [
                                f"v{rev['number']} - {rev['publishedAt']} (by {rev['publisher']})"
                                for rev in revisions_source
                            ]
                            selected_revision_source = st.selectbox(
                                "Choose revision",
                                options=revision_options_source,
                                key="source_revision"
                            )
                            
                            # Extract revision number
                            source_rev_num = revision_options_source.index(selected_revision_source)
                            st.info(f"✅ Source: **{selected_workbook_source}** - v{revisions_source[source_rev_num]['number']}")
                        else:
                            st.warning("No revisions found for this workbook.")
                else:
                    st.warning("No workbooks in this project.")
    
    # ========== TARGET WORKBOOK (RIGHT COLUMN) ==========
    with col2:
        st.subheader("📗 Target Workbook")
        
        projects = list_projects(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token)
        project_names = [p["name"] for p in projects]
        
        if not projects:
            st.warning("No projects available. Check your access permissions.")
        else:
            selected_project_target = st.selectbox(
                "Select Target Project",
                options=project_names,
                key="target_project"
            )
            
            # Get project ID
            target_project_id = next((p["id"] for p in projects if p["name"] == selected_project_target), None)
            
            if target_project_id:
                # Workbook Selection
                workbooks_target = list_workbooks_in_project(
                    st.session_state.site_url,
                    st.session_state.site_id,
                    st.session_state.auth_token,
                    target_project_id
                )
                
                workbook_names_target = [wb["name"] for wb in workbooks_target]
                
                if workbooks_target:
                    selected_workbook_target = st.selectbox(
                        "Select Target Workbook",
                        options=workbook_names_target,
                        key="target_workbook"
                    )
                    
                    # Get workbook ID
                    target_workbook_id = next((wb["id"] for wb in workbooks_target if wb["name"] == selected_workbook_target), None)
                    
                    if target_workbook_id:
                        # Revision Selection
                        st.write("**Select Revision:**")
                        revisions_target = get_workbook_revisions(
                            st.session_state.site_url,
                            st.session_state.site_id,
                            st.session_state.auth_token,
                            target_workbook_id
                        )
                        
                        if revisions_target:
                            revision_options_target = [
                                f"v{rev['number']} - {rev['publishedAt']} (by {rev['publisher']})"
                                for rev in revisions_target
                            ]
                            selected_revision_target = st.selectbox(
                                "Choose revision",
                                options=revision_options_target,
                                key="target_revision"
                            )
                            
                            # Extract revision number
                            target_rev_num = revision_options_target.index(selected_revision_target)
                            st.info(f"✅ Target: **{selected_workbook_target}** - v{revisions_target[target_rev_num]['number']}")
                        else:
                            st.warning("No revisions found for this workbook.")
                else:
                    st.warning("No workbooks in this project.")
    
    # ========== COMPARE BUTTON ==========
    st.divider()
    col_btn = st.columns([1, 3, 1])
    
    with col_btn[1]:
        if st.button("🚀 Compare Workbooks", use_container_width=True, type="primary"):
            st.info("📊 Comparison feature coming soon! Integration with your main comparison engine in progress...")
            # TODO: Integrate with your dev_2_prod 3.py comparison logic here

else:
    # Not authenticated
    st.warning("👆 Please enter your Tableau credentials in the sidebar to get started.")
