import streamlit as st
import os
import xml.etree.ElementTree as ET
import requests
import urllib3
from pathlib import Path

# import core comparison helpers from the CLI module
from tableau_comparator import (
    sign_in as _placeholder_signin,  # avoid accidental use
    get_workbook_id_in_project,
    get_revisions,
    get_workbook_owner,
    get_latest_revision_info,
    download_rev,
    parse_twb,
    extract_sections,
    build_cards,
    populate_change_registry_from_cards,
    build_overall_workbook_summary_card,
    build_users_permissions_card_with_context,
    ensure_change_registry_keys,
    CHANGE_REGISTRY,
    build_workbook_kpi_snapshot,
    render_workbook_kpi_table,
    render_visual_change_tree,
    xmldiff_text,
    sanitize_name,
    write_text,
    generate_html_report,
    get_users_and_permissions_for_workbook,
)

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


# --------------------------------------------------------------------------
# comparison orchestration for the Streamlit UI
# --------------------------------------------------------------------------
def perform_comparison(
    source_project,
    source_workbook,
    target_project,
    target_workbook,
    source_revision_number,
    target_revision_number,
):
    """Run the full comparator logic and return path to generated HTML report.

    Arguments are all strings (revision numbers may be numeric strings).
    """
    token = st.session_state.get("auth_token")
    site_id = st.session_state.get("site_id")
    site_url = st.session_state.get("site_url")

    if not token or not site_id or not site_url:
        st.error("Authentication state lost. Please reconnect.")
        return None

    # look up workbook ids
    try:
        source_wid, source_project_id = get_workbook_id_in_project(
            token, site_id, source_project, source_workbook
        )
        target_wid, target_project_id = get_workbook_id_in_project(
            token, site_id, target_project, target_workbook
        )
    except Exception as e:
        st.error(f"Lookup failed: {str(e)}")
        return None

    st.write(f"🔎 Source workbook ID: {source_wid}")
    st.write(f"🔎 Target workbook ID: {target_wid}")

    # download the specific revisions
    twb_old = download_rev(token, site_id, source_wid, source_revision_number, force=False)
    twb_new = download_rev(token, site_id, target_wid, target_revision_number, force=False)

    root_old = parse_twb(twb_old)
    root_new = parse_twb(twb_new)
    if root_old is None or root_new is None:
        st.error("Unable to parse downloaded workbook(s).")
        return None

    sec_old = extract_sections(root_old)
    sec_new = extract_sections(root_new)

    cards = build_cards(sec_old, sec_new)
    populate_change_registry_from_cards(cards)

    overall_summary_card = build_overall_workbook_summary_card(
        sec_old, sec_new, cards, root_old, root_new
    )
    if overall_summary_card:
        cards.insert(0, overall_summary_card)

    # permissions
    source_permissions = get_users_and_permissions_for_workbook(
        token, site_id, source_project, source_workbook
    )
    target_permissions = get_users_and_permissions_for_workbook(
        token, site_id, target_project, target_workbook
    )

    users_permissions_html = (
        build_users_permissions_card_with_context(
            source_project,
            source_workbook,
            source_permissions,
            context="source",
        )
        + "<hr/>"
        + build_users_permissions_card_with_context(
            target_project,
            target_workbook,
            target_permissions,
            context="target",
        )
    )

    # structural diff
    structural = xmldiff_text(
        ET.tostring(root_old, encoding="unicode"),
        ET.tostring(root_new, encoding="unicode"),
    )
    safe_wb = sanitize_name(f"{source_workbook}_VS_{target_workbook}")
    struct_path = f"{safe_wb}_STRUCT.txt"
    write_text(struct_path, structural)

    # kpis + visual tree
    kpi_old = build_workbook_kpi_snapshot(sec_old)
    kpi_new = build_workbook_kpi_snapshot(sec_new)
    kpi_html = render_workbook_kpi_table(kpi_old, kpi_new)
    visual_tree_text = render_visual_change_tree(sec_new, CHANGE_REGISTRY, target_workbook)

    # revision metadata
    source_owner = get_workbook_owner(token, site_id, source_wid)
    target_owner = get_workbook_owner(token, site_id, target_wid)
    source_revs = get_revisions(token, site_id, source_wid)
    target_revs = get_revisions(token, site_id, target_wid)
    source_latest = get_latest_revision_info(source_revs, source_owner)
    target_latest = get_latest_revision_info(target_revs, target_owner)

    OLD_PUBLISHER = source_latest.get("publisher")
    NEW_PUBLISHER = target_latest.get("publisher")
    LATEST_PUBLISHER = NEW_PUBLISHER
    LATEST_REVISION = target_revision_number
    LATEST_PUBLISHED_AT = target_latest.get("publishedAt", "")

    # generate HTML report
    report_name = sanitize_name(
        f"{source_workbook}_v{source_revision_number}_vs_{target_workbook}_v{target_revision_number}.html"
    )
    out_file = report_name
    generate_html_report(
        f"{source_workbook} v{source_revision_number}",
        f"{target_workbook} v{target_revision_number}",
        cards,
        None,
        out_file,
        kpi_html,
        root_new,
        visual_tree_text,
        LATEST_PUBLISHER,
        LATEST_REVISION,
        LATEST_PUBLISHED_AT,
        OLD_PUBLISHER,
        NEW_PUBLISHER,
        users_permissions_html,
    )

    return out_file


# ============================================================================
# UI INITIALIZATION & STATE MANAGEMENT
# ============================================================================

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
            # determine revision numbers from the selections (they're stored earlier in the UI)
            try:
                source_rev_num = revisions_source[
                    revision_options_source.index(selected_revision_source)
                ]["number"]
                target_rev_num = revisions_target[
                    revision_options_target.index(selected_revision_target)
                ]["number"]
            except Exception:
                st.error("Unable to determine selected revision numbers.")
                source_rev_num = None
                target_rev_num = None

            if not source_rev_num or not target_rev_num:
                st.warning("Please make sure both revisions are selected.")
            else:
                with st.spinner("Running comparison… this may take a minute"):
                    report_path = perform_comparison(
                        selected_project_source,
                        selected_workbook_source,
                        selected_project_target,
                        selected_workbook_target,
                        source_rev_num,
                        target_rev_num,
                    )
                if report_path:
                    st.success(f"✅ Comparison complete, report saved to `{report_path}`")
                    # display the generated HTML
                    try:
                        with open(report_path, "r", encoding="utf-8") as f:
                            html_content = f.read()
                        st.components.v1.html(html_content, height=800, scrolling=True)
                    except Exception as e:
                        st.error(f"Failed to render report: {e}")
        

else:
    # Not authenticated
    st.warning("👆 Please enter your Tableau credentials in the sidebar to get started.")
