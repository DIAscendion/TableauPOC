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

@st.cache_data(show_spinner="Refreshing projects...", ttl=600) # Cache for 10 minutes
def list_projects(site_url, site_id, token):
    try:
        base_url = site_url.rstrip('/')
        url = f"{base_url}/api/3.25/sites/{site_id}/projects?pageSize=1000"
        
        headers = {
            "X-Tableau-Auth": token,
            "Accept": "application/json"
        }
        
        r = requests.get(url, headers=headers, verify=VERIFY_SSL, timeout=60)
        
        if r.status_code == 401:
            # This is the "Unauthorized" flag
            return "AUTH_EXPIRED"
            
        r.raise_for_status()
        root = ET.fromstring(r.text)
        ns = {"t": "http://tableau.com/api"}
        
        projects = []
        for proj in root.findall(".//t:project", ns):
            projects.append({
                "id": proj.attrib.get("id"),
                "name": proj.attrib.get("name")
            })
        return sorted(projects, key=lambda x: x['name'])
    except Exception as e:
        st.error(f"Project Fetch Error: {str(e)}")
        return []
    
@st.cache_data(show_spinner="Fetching workbooks...")
def list_workbooks_in_project(site_url, site_id, token, project_id):
    """List all workbooks in a project."""
    try:
        base_url = site_url.rstrip('/')
        # Correct way to filter workbooks by project in the REST API
        url = f"{base_url}/api/3.25/sites/{site_id}/workbooks?filter=projectId:eq:{project_id}"
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
    # 1. FETCH PROJECTS & VALIDATE SESSION
    projects = list_projects(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token)
    
    if projects == "AUTH_EXPIRED":
        st.session_state.authenticated = False
        st.sidebar.error("⚠️ Session Expired. Please reconnect.")
        st.rerun()
    
    if not projects:
        st.warning("No projects available.")
        st.stop()

    project_names = [p["name"] for p in projects]
    col1, col2 = st.columns(2)
    
    # We initialize these as None so the button knows if selections are complete
    source_rev_final = None
    target_rev_final = None

    # ========== SOURCE WORKBOOK (LEFT) ==========
    with col1:
        st.subheader("📘 Source Workbook")
        sel_proj_src = st.selectbox("Select Project", options=project_names, key="src_proj_sel")
        src_proj_id = next((p["id"] for p in projects if p["name"] == sel_proj_src), None)
        
        if src_proj_id:
            wbs_src = list_workbooks_in_project(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token, src_proj_id)
            if wbs_src:
                sel_wb_src = st.selectbox("Select Workbook", options=[w["name"] for w in wbs_src], key="src_wb_sel")
                src_wb_id = next((w["id"] for w in wbs_src if w["name"] == sel_wb_src), None)
                
                revs_src = get_workbook_revisions(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token, src_wb_id)
                if revs_src:
                    rev_opts_src = [f"v{r['number']} - {r['publishedAt']} ({r['publisher']})" for r in revs_src]
                    sel_rev_src = st.selectbox("Choose Revision", options=rev_opts_src, key="src_rev_sel")
                    # Store the actual revision number
                    source_rev_final = revs_src[rev_opts_src.index(sel_rev_src)]["number"]
                else:
                    st.warning("No revisions found.")

    # ========== TARGET WORKBOOK (RIGHT) ==========
    with col2:
        st.subheader("📗 Target Workbook")
        sel_proj_tgt = st.selectbox("Select Project", options=project_names, key="tgt_proj_sel")
        tgt_proj_id = next((p["id"] for p in projects if p["name"] == sel_proj_tgt), None)
        
        if tgt_proj_id:
            wbs_tgt = list_workbooks_in_project(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token, tgt_proj_id)
            if wbs_tgt:
                sel_wb_tgt = st.selectbox("Select Workbook", options=[w["name"] for w in wbs_tgt], key="tgt_wb_sel")
                tgt_wb_id = next((w["id"] for w in wbs_tgt if w["name"] == sel_wb_tgt), None)
                
                revs_tgt = get_workbook_revisions(st.session_state.site_url, st.session_state.site_id, st.session_state.auth_token, tgt_wb_id)
                if revs_tgt:
                    rev_opts_tgt = [f"v{r['number']} - {r['publishedAt']} ({r['publisher']})" for r in revs_tgt]
                    sel_rev_tgt = st.selectbox("Choose Revision", options=rev_opts_tgt, key="tgt_rev_sel")
                    target_rev_final = revs_tgt[rev_opts_tgt.index(sel_rev_tgt)]["number"]
                else:
                    st.warning("No revisions found.")

    # ========== COMPARE BUTTON ==========
    st.divider()
    if st.button("🚀 Compare Workbooks", use_container_width=True, type="primary"):
        if not source_rev_final or not target_rev_final:
            st.warning("Please ensure both workbooks and revisions are selected.")
        else:
            with st.spinner("Analyzing differences..."):
                report_path = perform_comparison(
                    sel_proj_src, sel_wb_src, 
                    sel_proj_tgt, sel_wb_tgt, 
                    source_rev_final, target_rev_final
                )
                if report_path:
                    st.success(f"Report Generated: {report_path}")
                    with open(report_path, "r", encoding="utf-8") as f:
                        st.components.v1.html(f.read(), height=900, scrolling=True)
else:
    # Not authenticated
    st.warning("👆 Please enter your Tableau credentials in the sidebar to get started.")
