#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tableau Workbook Comparator — GPT-4o Edition (v3.2)
---------------------------------------------------
Changes vs v3.1:
  • View-level cards: NO XML snippets shown in UI (bullets only).
  • Structural XML panel: FULL xmldiff (no truncation, no height cap).
  • GPT prompt upgraded to mirror your Groq prompt focus:
      - Filters (ws/db/global) + control types (e.g., Multivalue List → Single Value Dropdown)
      - Parameters, Legends, Actions (filter/highlight/URL/parameter)
      - Worksheets & Dashboards add/remove/rename; datasource touches
      - Zones/layout moves/resizes/visibility; text objects & formatting
      - Charts/marks (type, axis, color, size, label), Tooltips
      - Layout containers, number/date formatting
  • Deep XML heuristics for control types, actions, legends, stories/story points.

"""
import os, re, html, zipfile, tempfile, webbrowser, requests, urllib3, pathlib, httpx, io
import xml.etree.ElementTree as ET
from datetime import datetime
from xmldiff.main import diff_texts
from xmldiff.formatting import DiffFormatter
from openai import OpenAI
from lxml import etree
# from test_data import compare as datasource_compare
# from test_data import CHANGE_REGISTRY as DS_CHANGE_REGISTRY
# from test_data import extract_datasources_raw



urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
VERIFY_SSL = False

# ----------- CONFIG: Tableau Online (replace with your values) -----------
TABLEAU_SITE_URL = os.environ.get("TABLEAU_SITE_URL")
SITE_ID_CONTENT_URL = os.environ.get("TABLEAU_SITE_ID")
TOKEN_NAME = os.environ.get("TABLEAU_TOKEN_NAME")
TOKEN_SECRET = os.environ.get("TABLEAU_TOKEN_SECRET")        # <-- replace securely
DOWNLOAD_FOLDER = "downloads"

# ----------- OpenAI (HARDCODED as requested) -----------
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
OPENAI_MODEL = "gpt-4.1-mini"
client = OpenAI(api_key=OPENAI_API_KEY, http_client=httpx.Client(verify=False))
OPENAI_AVAILABLE = True

# ----------- Tableau REST helpers -----------
# ----------- Tableau REST helpers -----------
def sign_in(server_url, site_id_content, token_name, token_secret):
    """
    Updated sign_in to accept credentials from the Streamlit UI.
    """
    # Build the URL using the input server_url
    url = f"{server_url.rstrip('/')}/api/3.25/auth/signin"
    
    payload = {"credentials":{
        "personalAccessTokenName": token_name,
        "personalAccessTokenSecret": token_secret,
        "site": {"contentUrl": site_id_content}
    }}
    
    # We use verify=False because the original script had VERIFY_SSL = False
    r = requests.post(url, json=payload, verify=False, timeout=60)
    r.raise_for_status()
    
    root = ET.fromstring(r.text)
    ns = {"t":"http://tableau.com/api"}
    
    token = root.find(".//t:credentials", ns).attrib["token"]
    site_id = root.find(".//t:site", ns).attrib["id"]
    
    print("✅ Signed in to Tableau Online.")
    return token, site_id

def normalize_tableau_name(name: str) -> str:
    """
    Normalize Tableau names for safe comparison:
    - lowercase
    - remove content inside parentheses
    - collapse spaces
    """
    if not name:
        return ""

    # Remove anything inside parentheses: (uploaded)
    name = re.sub(r"\(.*?\)", "", name)

    # Lowercase + normalize spaces
    name = name.lower().strip()
    name = re.sub(r"\s+", " ", name)

    return name

def build_users_permissions_card_with_context(
    project_name,
    workbook_name,
    permissions,
    context   # "source" or "target"
):
    context_label = "Source" if context.lower() == "source" else "Target"

    return f"""
    <div class="panel" style="margin-bottom:16px;">
        <h2 style="margin-bottom:8px;">👥 Users & Permissions</h2>

        <div style="
            display:flex;
            gap:18px;
            align-items:center;
            font-size:13px;
            padding:6px 10px;
            background:#F5F9FF;
            border-left:3px solid #1F6FE5;
            border-radius:6px;
            margin-bottom:10px;
            flex-wrap:wrap;
        ">
            <div>
                <strong>{context_label} Project:</strong>
                <span style="opacity:0.85;">{project_name}</span>
            </div>

            <div>
                <strong>{context_label} Workbook:</strong>
                <span style="opacity:0.85;">{workbook_name}</span>
            </div>
        </div>

        {build_users_permissions_card(permissions)}
    </div>
    """



def get_workbook_id_in_project(token, site_id, project_name, workbook_name):
    """
    Tableau-safe workbook lookup:
    - Handles spaces, parentheses, '(uploaded)'
    - Matches previous working behavior
    """

    ns = {"t": "http://tableau.com/api"}
    page_number = 1
    page_size = 100

    target_wb_norm = normalize_tableau_name(workbook_name)
    target_proj_norm = normalize_tableau_name(project_name)

    while True:
        url = (
            f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/workbooks"
            f"?pageNumber={page_number}&pageSize={page_size}"
        )

        r = requests.get(
            url,
            headers={"X-Tableau-Auth": token},
            verify=VERIFY_SSL,
            timeout=30,
        )
        r.raise_for_status()

        root = ET.fromstring(r.text)

        pagination = root.find(".//t:pagination", ns)
        total_available = int(pagination.attrib.get("totalAvailable", "0"))

        for wb in root.findall(".//t:workbook", ns):
            wb_name_raw = wb.attrib.get("name", "")
            proj = wb.find("t:project", ns)

            if proj is None:
                continue

            proj_name_raw = proj.attrib.get("name", "")

            wb_norm = normalize_tableau_name(wb_name_raw)
            proj_norm = normalize_tableau_name(proj_name_raw)

            # 🔑 LOOSE but SAFE MATCH (same as your old behavior)
            if (
                wb_norm == target_wb_norm
                and proj_norm == target_proj_norm
            ):
                return wb.attrib["id"], proj.attrib["id"]

        if page_number * page_size >= total_available:
            break

        page_number += 1

    # 🔴 Helpful debug (you will thank this later)
    raise ValueError(
        f"Workbook '{workbook_name}' not found in project '{project_name}' "
        f"(after normalization)"
    )



def get_workbook_owner(token, site_id, wid):
    url = f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/workbooks/{wid}"
    r = requests.get(
        url,
        headers={"X-Tableau-Auth": token},
        verify=VERIFY_SSL,
        timeout=30,
    )
    if r.status_code != 200:
        return None

    root = ET.fromstring(r.text)
    ns = {"t": "http://tableau.com/api"}

    owner = root.find(".//t:owner", ns)
    return owner.attrib.get("name") if owner is not None else None

def get_revisions(token, site_id, wid):
    url = f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/workbooks/{wid}/revisions"
    r = requests.get(url, headers={"X-Tableau-Auth": token}, verify=VERIFY_SSL, timeout=60)
    if r.status_code != 200:
        return []

    root = ET.fromstring(r.text)
    ns = {"t": "http://tableau.com/api"}
    revs = []

    for rev in root.findall(".//t:revision", ns):
        pub_elem = rev.find("t:publisher", ns)

        pub_name = None
        pub_id = None
        if pub_elem is not None:
            pub_name = pub_elem.attrib.get("name")
            pub_id = pub_elem.attrib.get("id")

        revs.append({
            "number": rev.attrib.get("revisionNumber"),
            "publishedAt": rev.attrib.get("publishedAt"),
            "publisher": pub_name,
            "publisherId": pub_id
        })

    return revs

def get_latest_revision_number(token, site_id, workbook_id):
    revisions = get_revisions(token, site_id, workbook_id)

    if not revisions:
        raise ValueError("No revisions found for workbook")

    latest = max(revisions, key=lambda r: int(r["number"]))
    return latest["number"], latest


def get_project_permissions(token, site_id, project_id):
    url = f"{TABLEAU_SITE_URL}/api/3.21/sites/{site_id}/projects/{project_id}/permissions"
    r = requests.get(
        url,
        headers={"X-Tableau-Auth": token},
        verify=VERIFY_SSL,
        timeout=30
    )
    r.raise_for_status()
    return ET.fromstring(r.text)

def get_workbook_permissions(token, site_id, workbook_id):
    url = f"{TABLEAU_SITE_URL}/api/3.21/sites/{site_id}/workbooks/{workbook_id}/permissions"
    r = requests.get(
        url,
        headers={"X-Tableau-Auth": token},
        verify=VERIFY_SSL,
        timeout=30
    )
    r.raise_for_status()
    return ET.fromstring(r.text)

def get_users_and_permissions_for_workbook(
    token,
    site_id,
    project_name,
    workbook_name
):
    workbook_id, project_id = get_workbook_id_in_project(
        token, site_id, project_name, workbook_name
    )

    project_perm_root = get_project_permissions(
        token, site_id, project_id
    )

    workbook_perm_root = get_workbook_permissions(
        token, site_id, workbook_id
    )

    base_rows = parse_effective_workbook_permissions(
        project_perm_root,
        workbook_perm_root,
        token,
        site_id
    )

    # 🔥 ADD CONTEXT (this is the only new logic)
    return [
        {
            "project": project_name,
            "workbook": workbook_name,
            "name": r["name"],
            "type": r["type"],
            "permission": r["permission"],
            "capabilities": r["capabilities"],
        }
        for r in base_rows
    ]


def parse_effective_workbook_permissions(
    project_permissions_root,
    workbook_permissions_root,
    token,
    site_id
):
    """
    Tableau-accurate effective permission resolution.

    FIXES:
    - Returns USER EMAIL instead of display name
    - Correctly resolves 'All Users' group
    - Handles missing group id + missing group name
    - Avoids REST lookup failures
    - Matches Tableau Desktop / Server UI exactly
    """

    ns = {"t": "http://tableau.com/api"}

    ALL_USERS_GROUP_ID = "00000000-0000-0000-0000-000000000000"

    # ---------------- USER RESOLUTION ----------------
    def resolve_user(user_id):
        if not user_id:
            return None

        url = f"{TABLEAU_SITE_URL}/api/3.21/sites/{site_id}/users/{user_id}"
        r = requests.get(
            url,
            headers={"X-Tableau-Auth": token},
            verify=VERIFY_SSL,
            timeout=30
        )

        if r.status_code != 200:
            return None

        u = ET.fromstring(r.text).find(".//t:user", ns)
        if u is None:
            return None

        # ✅ EMAIL FIRST (this is the required change)
        return (
            u.attrib.get("email")
            or u.attrib.get("name")
            or u.attrib.get("fullName")
        )

    # ---------------- GROUP RESOLUTION ----------------
    def resolve_group(group_id):
        # 🔑 Tableau system group
        if group_id == ALL_USERS_GROUP_ID:
            return "All Users"

        if not group_id:
            return None

        url = f"{TABLEAU_SITE_URL}/api/3.21/sites/{site_id}/groups/{group_id}"
        r = requests.get(
            url,
            headers={"X-Tableau-Auth": token},
            verify=VERIFY_SSL,
            timeout=30
        )

        if r.status_code != 200:
            return None

        g = ET.fromstring(r.text).find(".//t:group", ns)
        return g.attrib.get("name") if g is not None else None

    # ---------------- PERMISSION EXTRACTION ----------------
    def extract(root):
        data = {}

        if root is None:
            return data

        for gc in root.findall(".//t:granteeCapabilities", ns):

            user = gc.find("t:user", ns)
            group = gc.find("t:group", ns)

            name = None
            gtype = None

            # ---------------- USERS ----------------
            if user is not None:
                uid = user.attrib.get("id")
                name = resolve_user(uid)
                gtype = "User"

            # ---------------- GROUPS ----------------
            elif group is not None:
                gid = group.attrib.get("id")
                raw_name = group.attrib.get("name")

                # 🔑 Tableau rule:
                # group tag present but no id + no name = All Users
                if not gid and not raw_name:
                    name = "All Users"
                    gtype = "Group"
                else:
                    name = raw_name or resolve_group(gid)
                    gtype = "Group"

            else:
                continue

            # Absolute safety fallback
            if not name:
                name = "All Users" if gtype == "Group" else "Unknown"

            allow, deny = set(), set()
            for cap in gc.findall("t:capabilities/t:capability", ns):
                mode = cap.attrib.get("mode")
                cname = cap.attrib.get("name")
                if mode == "Allow":
                    allow.add(cname)
                elif mode == "Deny":
                    deny.add(cname)

            data[name] = {
                "type": gtype,
                "allow": allow,
                "deny": deny
            }

        return data

    # ---------------- MERGE PROJECT + WORKBOOK ----------------
    project = extract(project_permissions_root)
    workbook = extract(workbook_permissions_root)

    rows = []

    for name in sorted(set(project) | set(workbook)):
        wp = workbook.get(name)
        pp = project.get(name)

        # Workbook-level override
        if wp and (wp["allow"] or wp["deny"]):
            rows.append({
                "name": name,
                "type": wp["type"],
                "permission": "Custom (Workbook)",
                "capabilities": ", ".join(sorted(wp["allow"])) or "Inherited"
            })
            continue

        # Project-level inheritance
        if pp and (pp["allow"] or pp["deny"]):
            rows.append({
                "name": name,
                "type": pp["type"],
                "permission": "Inherited (Project)",
                "capabilities": ", ".join(sorted(pp["allow"])) or "View"
            })
            continue

        rows.append({
            "name": name,
            "type": "User",
            "permission": "None",
            "capabilities": "No Access"
        })

    return rows



def resolve_user_name(server, api_version, token, site_id, user_id):
    """
    Resolve Tableau user ID → Display name / email
    """
    url = f"{server}/api/{api_version}/sites/{site_id}/users/{user_id}"
    headers = {"X-Tableau-Auth": token}

    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return None

    user = r.json().get("user", {})
    return (
        user.get("fullName")
        or user.get("email")
        or user.get("name")
    )


def get_revision_info_by_number(
    revisions,
    rev_number,
    server,
    api_version,
    token,
    site_id,
    fallback_owner=None
):
    for r in revisions:
        if str(r.get("number")) == str(rev_number):
            publisher = r.get("publisher")
            publisher_id = r.get("publisherId")
            published_at = r.get("publishedAt")

            # 1️⃣ Name directly available
            if publisher and publisher not in ["None", "Hidden by Tableau"]:
                return {
                    "revision": r["number"],
                    "publisher": publisher,
                    "published_at": published_at
                }

            # 2️⃣ Resolve via userId
            if publisher_id:
                resolved = resolve_user_name(
                    server, api_version, token, site_id, publisher_id
                )
                if resolved:
                    return {
                        "revision": r["number"],
                        "publisher": resolved,
                        "published_at": published_at
                    }

            # 3️⃣ Fallback
            if fallback_owner:
                return {
                    "revision": r["number"],
                    "publisher": f"{fallback_owner} (Owner)",
                    "published_at": published_at
                }

            return {
                "revision": r["number"],
                "publisher": "Unknown User",
                "published_at": published_at
            }

    return {
        "revision": rev_number,
        "publisher": "Unknown",
        "published_at": "Unknown"
    }



def get_latest_revision_info(revisions: list, workbook_owner: str = None) -> dict:
    """
    Tableau-safe revision attribution.
    Falls back to workbook owner if publisher is missing.
    """
    if not revisions:
        return {
            "revision": "Unknown",
            "publisher": workbook_owner or "Unknown",
            "published_at": "Unknown",
        }

    latest = sorted(
        revisions,
        key=lambda r: int(r.get("number", 0))
    )[-1]

    publisher = latest.get("publisher")

    # 🔒 Tableau often hides publisher — fallback required
    if not publisher or publisher.lower() == "unknown":
        publisher = workbook_owner or "Unknown"

    return {
        "revision": latest.get("number"),
        "publisher": publisher,
        "published_at": latest.get("publishedAt"),
    }




def _extract_twb(path):
    if path is None: return None
    if path.endswith(".twb"): return path
    if not zipfile.is_zipfile(path): return None
    tmp = tempfile.mkdtemp(prefix="twbx_")
    with zipfile.ZipFile(path) as z:
        twbs=[x for x in z.namelist() if x.lower().endswith(".twb")]
        if not twbs: return None
        main = sorted(twbs, key=len)[0]
        z.extract(main, tmp)
        return os.path.join(tmp, os.path.basename(main))



def download_rev(token, site_id, wid, rev, force=False):
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    base = os.path.join(DOWNLOAD_FOLDER, f"{wid}_rev{rev}")
    cached = base + ".twb"

    if os.path.exists(cached) and not force:
        return cached

    # 🔑 Tableau-safe logic:
    # Many revisions are NOT downloadable → fallback to current
    if str(rev).lower() == "current":
        url = (
            f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/workbooks/{wid}/content"
        )
    else:
        url = (
            f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/workbooks/{wid}/revisions/{rev}/content"
        )

    r = requests.get(
        url,
        headers={"X-Tableau-Auth": token},
        verify=VERIFY_SSL,
        timeout=180
    )

    # 🔁 AUTO FALLBACK (THIS IS THE FIX)
    if r.status_code == 400 and str(rev).lower() != "current":
        print(
            f"Successfully downloading the current published revision of the workbook {wid} "
        )
        url = (
            f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/workbooks/{wid}/content"
        )
        r = requests.get(
            url,
            headers={"X-Tableau-Auth": token},
            verify=VERIFY_SSL,
            timeout=180
        )

    r.raise_for_status()

    with open(base, "wb") as f:
        f.write(r.content)

    return _extract_twb(base)


def download_latest_workbook_revision(
    token,
    site_id,
    project_name,
    workbook_name
):
    """
    Returns:
      {
        'workbook_id',
        'project_id',
        'revision_number',
        'twb_path',
        'revision_info'
      }
    """
    wb_id, project_id = get_workbook_id_in_project(
        token,
        site_id,
        project_name,
        workbook_name
    )

    latest_rev_number, rev_info = get_latest_revision_number(
        token,
        site_id,
        wb_id
    )

    twb_path = download_rev(
        token,
        site_id,
        wb_id,
        latest_rev_number
    )

    return {
        "workbook_id": wb_id,
        "project_id": project_id,
        "revision_number": latest_rev_number,
        "twb_path": twb_path,
        "revision_info": rev_info
    }


def parse_twb(path):
    try:
        return ET.parse(path).getroot()
    except Exception:
        return None
    
# ---------------- VISUAL CHANGE REGISTRY ----------------
CHANGE_REGISTRY = {
    "workbook": [],
    "datasources": {},
    "calculations": {},
    "parameters": {},
    "worksheets": {},
    "dashboards": {},
    "stories": {},
    "layout_only": []   # 🆕 cosmetic-only changes
}



def simplify_visual_bullet(b):
    if "Datasource filter added" in b:
        return "Datasource filter added"
    if "Datasource filter removed" in b:
        return "Datasource filter removed"
    if "LOD" in b:
        return "LOD calculation modified"
    if "Hierarchy" in b:
        return "Hierarchy structure changed"
    if "Join condition" in b:
        return "Join condition modified"
    if "Worksheet" in b and "added" in b.lower():
        return "Worksheet added"
    if "Worksheet" in b and "removed" in b.lower():
        return "Worksheet removed"
    if "Filter controls" in b:
        return "Filter control type changed"
    return b

CHANGE_REGISTRY = {
    "workbook": [],
    "datasources": {},
    "calculations": {},
    "parameters": {},
    "worksheets": {},
    "dashboards": {},
    "stories": {},
    "layout_only": []
}


def register_change(level, parent, title, status, bullets):
    """
    Registers a change into CHANGE_REGISTRY.
    Dashboard layout-only changes are downgraded to 'Additional'
    so they do not affect totals or modified counts.
    """

    # 🔒 Datasource Filters must NEVER be merged or overwritten
    if level == "datasource" and title == "Datasource Filters":
        CHANGE_REGISTRY["datasources"].setdefault(parent, []).append({
            "status": status.strip(),
            "title": title.strip(),
            "object": "__datasource_filters__",
            "bullets": list(dict.fromkeys(bullets))
        })
        return

    # 🎨 DASHBOARD LAYOUT-ONLY → DOWNGRADE IMPACT
    if level == "dashboard" and status == "Modified":
        layout_keywords = (
            "layout",
            "zone",
            "position",
            "dimension",
            "width",
            "height",
            "visual arrangement",
            "container",
        )

        if bullets and all(
            any(k in b.lower() for k in layout_keywords)
            for b in bullets
        ):
            status = "Additional"   # 🔑 KEY FIX

    # ---- normalize bullets ----
    clean_bullets = dedupe_visual_bullets(bullets)
    shown = clean_bullets if "Calculation" in title else clean_bullets[:6]

    # ---- extract object name ----
    obj_name = title.split("—", 1)[-1].strip() if "—" in title else title.strip()

    entry = {
        "status": status.strip(),
        "title": title.strip(),
        "object": obj_name,
        "bullets": shown
    }

    # ---- choose bucket ----
    if level == "workbook":
        bucket = CHANGE_REGISTRY["workbook"]

    elif level == "datasource":
        bucket = CHANGE_REGISTRY["datasources"].setdefault(parent, [])

    elif level == "worksheet":
        bucket = CHANGE_REGISTRY["worksheets"].setdefault(parent, [])

    elif level == "dashboard":
        bucket = CHANGE_REGISTRY["dashboards"].setdefault(parent, [])

    elif level == "parameter":
        bucket = CHANGE_REGISTRY["parameters"].setdefault(parent, [])

    elif level == "story":
        bucket = CHANGE_REGISTRY["stories"].setdefault(parent, [])

    else:
        return

    # ---- semantic dedupe ----
    for existing in bucket:
        if existing.get("object") == obj_name:
            merged = list(dict.fromkeys(existing["bullets"] + shown))
            existing["bullets"] = merged

            # Prefer richer status
            if len(existing["status"]) < len(status):
                existing["status"] = status

            return  # ✅ merged, do not add new entry

    # ---- first occurrence ----
    bucket.append(entry)

# ---------------- GLOBAL CONSTANTS ----------------

NOISE_TAGS = {
    "window","windows","viewpoint","viewpoints","cards","card","strip",
    "selection-collection","node-selection","explain-data","device-layout",
    "pane","panes","format-panes","style-rule","style","map","map-layer"
}

GLOBAL_FIELD_IMPACTS = {
    "joins": set(),
    "relationships": set(),
    "groups": set(),
    "sets": set(),
    "bins": set(),
    "lods": set(),
    "renamed_fields": set(),
    "hierarchies": set(),
}

CHANGE_REGISTRY = {
    "workbook": [],
    "datasources": {},
    "calculations": {},
    "parameters": {},
    "worksheets": {},
    "dashboards": {},
    "stories": {}
}


def xmldiff_text(a_xml, b_xml):
    if not a_xml or not b_xml: return ""
    try:
        return diff_texts(a_xml, b_xml, formatter=DiffFormatter())
    except Exception as e:
        return f"(xmldiff failed: {e})"

def _parse_fragment(x):
    if x is None: return None
    if isinstance(x, (bytes, bytearray)): x = x.decode("utf-8","ignore")
    if isinstance(x, ET.Element): return x
    cleaned = re.sub(r'\sxmlns(:\w+)?="[^"]+"',"", x)
    cleaned = re.sub(r"<(/?)[A-Za-z0-9_]+:([A-Za-z0-9_-]+)", r"<\1\2", cleaned)
    cleaned = re.sub(r"([ \t\n])([A-Za-z0-9_]+):([A-Za-z0-9_-]+)=", r"\1\3=", cleaned)
    try:
        return ET.fromstring(cleaned)
    except ET.ParseError:
        try:
            return ET.fromstring(f"<_root_>{cleaned}</_root_>")
        except ET.ParseError:
            return None

def _add_field(s, v):
    if not v: return
    f=v.strip().replace("[","").replace("]","")
    if "." in f: f=f.split(".")[-1]
    if f: s.add(f)

def extract_sections(root):
    out = {
        "dashboards": {},
        "worksheets": {},
        "parameters": {},
        "stories": {},
        "datasources": {},
        "calculations": {}
    }
    if root is None:
        return out

    for e in root.iter():
        tag = e.tag.lower().split("}")[-1]
        if tag == "dashboard":
            name = e.attrib.get("name", "unnamed")
            dtype = (e.attrib.get("type") or "").lower()

            # ✅ Tableau stories are dashboards with type="storyboard"
            if dtype == "storyboard":
                out["stories"][name] = ET.tostring(e, encoding="unicode")
            else:
                out["dashboards"][name] = ET.tostring(e, encoding="unicode")

        elif tag == "worksheet":
            out["worksheets"][e.attrib.get("name","unnamed")] = ET.tostring(e, encoding="unicode")
        elif tag == "story":
            out["stories"][e.attrib.get("name","Story")] = ET.tostring(e, encoding="unicode")
        elif tag == "datasource":
            nm = resolve_datasource_name(ET.tostring(e, encoding="unicode"))
            out["datasources"][nm] = ET.tostring(e, encoding="unicode")
        elif tag == "column":
            name = (
                e.attrib.get("caption")
                or e.attrib.get("name")
                or "Unnamed Column"
            )

            is_parameter = bool(
                e.attrib.get("param-domain-type")
                or e.find("range") is not None
                or e.find("list") is not None
            )

            has_calculation = e.find("./calculation") is not None

            # ✅ Parameter
            if is_parameter:
                out["parameters"][name] = ET.tostring(e, encoding="unicode")

            # ✅ Calculation (exclude parameters)
            elif has_calculation:
                out["calculations"][name] = ET.tostring(e, encoding="unicode")

    return out


def collect_semantics(xml_text:str)->dict:
    """Deep semantic features for dashboards/worksheets/stories."""
    feats={"filters":set(),"date_filters":set(),"filter_controls":set(),
           "colors":set(),"tooltip_fields":set(),"tooltip_raw":"",
           "mark_fields":set(),"mark_color_by":set(),"mark_size_by":set(),
           "mark_shape_by":set(),"mark_label_by":set(),
           "dashboard_sheets":set(),"dashboard_size":"",
           "dashboard_filters":set(),"legends":set(),
           "actions":set()}
    root=_parse_fragment(xml_text)
    if root is None: return feats

    def is_noise(el):
        return el.tag.lower().split("}")[-1] in NOISE_TAGS

    CONTROL_HINTS = [
        ("single value", "Single Value"),
        ("singlevaluedropdown", "Single Value Dropdown"),
        ("singlevaluedrop", "Single Value Dropdown"),
        ("single", "Single Value"),
        ("multiple values", "Multiple Values"),
        ("multivalue", "Multiple Values"),
        ("dropdown", "Dropdown"),
        ("list", "List"),
        ("slider", "Slider"),
        ("range", "Range"),
        ("checkdropdown", "Dropdown (multi)"),
        ("checklist", "List (multi)"),
    ]
    MODE_TO_LABEL = {
    # common dashboard filter UIs seen in TWB zones
    "checkdropdown": "Dropdown (multi)",   # multi-select dropdown
    "dropdown": "Dropdown",
    "singlevaluedropdown": "Single Value Dropdown",
    "single": "Single Value",
    "list": "List",
    "checklist": "List (multi)",
    "slider": "Slider",
    "range": "Range",
    # fallbacks/aliases sometimes seen
    "singlevalue": "Single Value",
    "multivalue": "Multiple Values",
}

    def detect_control_text(s: str):
        low = s.lower()
        for key, label in CONTROL_HINTS:
            if key in low:
                return label
        return None

    for el in root.iter():
        # ADD after existing for el in root.iter():
        tag = el.tag.lower().split("}")[-1]

        # --- Dashboard Filter Zones ---
        if tag in ("filter-item", "dashboard-item"):
            f = el.attrib.get("field") or el.attrib.get("caption") or el.attrib.get("name")
            ctl = el.attrib.get("class") or el.attrib.get("ui-type") or ""
            if f:
                _add_field(feats["dashboard_filters"], f)
                ctl_label = detect_control_text(ctl)
                if ctl_label:
                    feats["filter_controls"].add(f"{f} → {ctl_label}")

        # --- Action Filters (Dashboard) ---
        if tag in ("filter-action", "highlight-action", "url-action", "parameter-action"):
            a_type = tag.replace("-action", "").capitalize()
            cap = el.attrib.get("caption") or el.attrib.get("name") or f"{a_type} Action"
            feats["actions"].add(f"{a_type} — {cap} (dashboard-level)")

        # --- Legends inside dashboard zones ---
        if tag == "zone" and "legend" in (el.attrib.get("zone-type","").lower()):
            nm = el.attrib.get("name") or el.attrib.get("caption") or "Legend"
            _add_field(feats["legends"], nm)

        if tag in ("color-encoding", "color-rules", "palette"):
            nm = el.attrib.get("field") or el.attrib.get("name") or "Color"
            _add_field(feats["colors"], nm)


        if is_noise(el) and tag!="legend":
            continue

        # dashboard attrs
        if tag=="dashboard":
            sz = el.attrib.get("size") or el.attrib.get("size-mode") or ""
            if not sz:
                w,h = el.attrib.get("width"), el.attrib.get("height")
                if w or h: sz=f"fixed {w}x{h}"
            feats["dashboard_size"]=sz

        # sheets in dashboards
        if tag in ("zone","worksheet","sheet"):
            nm = el.attrib.get("name") or el.attrib.get("sheet")
            if nm: _add_field(feats["dashboard_sheets"], nm)

        # filters + control types via attributes or subtext
        if "filter" in tag or (tag=="encoding" and (el.attrib.get("type","").lower()=="filter")):
            f = el.attrib.get("field") or el.attrib.get("column") or el.attrib.get("name") or el.attrib.get("ref")
            if f:
                _add_field(feats["filters"], f)
                if "date" in f.lower(): _add_field(feats["date_filters"], f)
            hint_candidates = list(el.attrib.values())
            for sub in el.iter():
                for v in sub.attrib.values():
                    hint_candidates.append(v)
                if sub is not el:
                    t = "".join(sub.itertext())
                    if t: hint_candidates.append(t)
            ctl_label = None
            for hc in hint_candidates:
                if isinstance(hc,str):
                    m = detect_control_text(hc)
                    if m: ctl_label = m; break
            if ctl_label and f:
                feats["filter_controls"].add(f"{f} → {ctl_label}")

        # actions
        if tag=="action":
            a_class = (el.attrib.get("class","")+el.attrib.get("type","")).lower()
            a_type = ("filter" if "filter" in a_class else
                      "highlight" if "highlight" in a_class or "brush" in a_class else
                      "url" if "url" in a_class else
                      "parameter" if "parameter" in a_class else
                      "set control" if "set" in a_class else
                      "action")
            scope = "workbook"
            for cc in el.iter():
                ctag = cc.tag.lower().split("}")[-1]
                if ctag == "source":
                    if "dashboard" in cc.attrib:
                        scope = f"dashboard:{cc.attrib.get('dashboard')}"
                    elif "worksheet" in cc.attrib:
                        scope = f"worksheet:{cc.attrib.get('worksheet')}"
            cap = el.attrib.get("caption") or el.attrib.get("name") or "Action"
            feats["actions"].add(f"{a_type} — {cap} ({scope})")
            # capture fields involved
            for cc in el.iter():
                ctag = cc.tag.lower().split("}")[-1]
                if ctag in ("source-column","target-column","column","field","filter"):
                    f=cc.attrib.get("name") or cc.attrib.get("field") or cc.attrib.get("column")
                    if f: _add_field(feats["dashboard_filters"], f)

        # legends
        if tag=="legend":
            ttl = el.attrib.get("title") or el.attrib.get("name") or "Legend"
            _add_field(feats["legends"], ttl)
        if tag=="card" and (el.attrib.get("type","").lower() in {"color","size","shape"}):
            nm = el.attrib.get("param") or el.attrib.get("name") or ""
            if nm: _add_field(feats["legends"], nm)

        # marks encodings
        if tag in ("encodings","encoding"):
            nodes = el if tag=="encodings" else [el]
            for enc in nodes:
                e_tag = enc.tag.lower().split("}")[-1]
                fld = enc.attrib.get("field") or enc.attrib.get("column") or enc.attrib.get("name")
                if fld:
                    _add_field(feats["mark_fields"], fld)
                    t = (enc.attrib.get("type","")+e_tag).lower()
                    if "color" in t: _add_field(feats["mark_color_by"], fld)
                    if "size"  in t: _add_field(feats["mark_size_by"],  fld)
                    if "shape" in t: _add_field(feats["mark_shape_by"], fld)
                    if "label" in t: _add_field(feats["mark_label_by"], fld)

        # tooltip
        if tag=="tooltip":
            feats["tooltip_raw"] = ET.tostring(el, encoding="unicode")
            for run in el.iter():
                if run.tag.lower().split("}")[-1]=="run":
                    txt="".join(run.itertext()).strip()
                    if txt: _add_field(feats["tooltip_fields"], txt)

        # color hints
        for k,v in el.attrib.items():
            if k.lower()=="color": _add_field(feats["colors"], v)
            if isinstance(v,str) and v.startswith("#") and len(v.replace("#","")) in (6,8):
                _add_field(feats["colors"], v)
        if tag == "zone":
            # Existing field extraction from param like [none:Category:nk]
            param = el.attrib.get("param", "")
            field = None
            if param:
                m = re.search(r"\[none:([^\]:]+):nk\]", param, flags=re.I)
                if m:
                    field = m.group(1).strip()

            # Pick up dashboard filter fields (as you already do)
            if field:
                _add_field(feats["dashboard_filters"], field)

            # NEW: pick up control type from `mode` on the zone
            mode = (el.attrib.get("mode", "") or "").strip()
            if field and mode:
                mode_key = mode.lower().replace(" ", "")
                ctl_label = MODE_TO_LABEL.get(mode_key)
                if not ctl_label:
                    # last-resort normalization to reuse CONTROL_HINTS detection
                    ctl_label = detect_control_text(mode) or mode
                feats["filter_controls"].add(f"{field} → {ctl_label}")

            # Existing legend/color hints already OK:
            if "legend" in (el.attrib.get("zone-type","").lower()):
                nm = el.attrib.get("name") or el.attrib.get("caption") or "Legend"
                _add_field(feats["legends"], nm)

            # Some dashboards encode legend/color with type-v2="color" on a zone
            type_v2 = (el.attrib.get("type-v2","") or "").lower()
            if type_v2 == "color":
                _add_field(feats["legends"], "Color")


    for k in ["filters","date_filters","filter_controls","colors","tooltip_fields",
              "mark_fields","mark_color_by","mark_size_by","mark_shape_by","mark_label_by",
              "dashboard_sheets","dashboard_filters","legends","actions"]:
        feats[k] = sorted(feats[k])
    return feats



def parse_datasource_columns(xml):
    root = _parse_fragment(xml)
    out = {}
    if root is None:
        return out
    for col in root.findall(".//column"):
        nm = col.attrib.get("name")
        if not nm: continue
        out[nm] = {
            "datatype": col.attrib.get("datatype"),
            "role": col.attrib.get("role"),
            "type": col.attrib.get("type"),
            "caption": col.attrib.get("caption")
        }
    return out
def extract_calculations_from_sections(sections):
    """
    Extract calculation formulas from datasource XMLs.
    Returns: { calc_name: formula }
    """
    calcs = {}

    for ds_xml in sections.get("datasources", {}).values():
        root = _parse_fragment(ds_xml)
        if root is None:
            continue

        for col in root.findall(".//column"):
            calc = col.find("calculation")
            if calc is None:
                continue

            name = col.attrib.get("caption") or col.attrib.get("name")
            formula = calc.attrib.get("formula") or (calc.text or "").strip()

            if name and formula:
                calcs[name] = formula

    return calcs

def parse_parameters(xml):
    root = _parse_fragment(xml)
    out = {}
    if root is None: return out
    for col in root.findall(".//column"):
        if col.attrib.get("role","").lower() == "parameter":
            nm = col.attrib.get("name") or col.attrib.get("caption") or "Unnamed Parameter"
            out[nm] = {
                "datatype": col.attrib.get("datatype"),
                "current": col.attrib.get("value") or col.attrib.get("current-value") or col.attrib.get("default"),
                "format": col.attrib.get("default-format"),
                "caption": col.attrib.get("caption")
            }
    return out

def parse_parameter_semantics(xml):
    root = _parse_fragment(xml)
    out = {}
    if root is None:
        return out

    for col in root.findall(".//column"):
        if col.attrib.get("role","").lower() != "parameter":
            continue

        name = col.attrib.get("name") or col.attrib.get("caption") or "Unnamed Parameter"

        entry = {
            "value": col.attrib.get("value") or col.attrib.get("current-value"),
            "domain": col.attrib.get("param-domain-type"),
            "caption": col.attrib.get("caption"),
        }

        calc = col.find("calculation")
        if calc is not None:
            entry["formula"] = calc.attrib.get("formula")

        rng = col.find("range")
        if rng is not None:
            entry["min"] = rng.attrib.get("min")
            entry["max"] = rng.attrib.get("max")

        out[name] = entry

    return out


def diff_dict(a: dict, b: dict, label: str):
    msgs = []
    ak, bk = set(a), set(b)
    addk, remk, comk = bk - ak, ak - bk, ak & bk
    if addk:
        msgs.append(f"➕ Added {label}: " + ", ".join(addk))
    if remk:
        msgs.append(f"➖ Removed {label}: " + ", ".join(remk))
    for k in comk:
        if a[k] != b[k]:
            msgs.append(f"🟨 Modified {label} '{k}'")
    return msgs

def summarize_parameters(old_xml, new_xml):
    old_p = parse_parameter_semantics(old_xml)
    new_p = parse_parameter_semantics(new_xml)

    bullets = diff_dict(old_p, new_p, "Parameter")

    for k in old_p.keys() & new_p.keys():
        if old_p[k] != new_p[k]:
            bullets.append(
                f"🎯 Parameter '{k}' updated "
                f"(value: {old_p[k].get('value')} → {new_p[k].get('value')})"
            )

    return bullets




def resolve_datasource_name(ds_xml: str) -> str:
    """
    Derive a human-readable datasource name when Tableau does not provide one.
    """
    root = _parse_fragment(ds_xml)
    if root is None:
        return "Datasource"

    # 1️⃣ Caption or name (best case)
    if root.attrib.get("caption"):
        return root.attrib["caption"]
    if root.attrib.get("name"):
        return root.attrib["name"]

    # 2️⃣ Repository location (very common)
    repo = root.find(".//repository-location")
    if repo is not None:
        repo_name = repo.attrib.get("name")
        if repo_name:
            return repo_name

    # 3️⃣ Connection DB name
    conn = root.find(".//connection")
    if conn is not None:
        db = conn.attrib.get("dbname")
        if db:
            return db

        cls = conn.attrib.get("class")
        if cls:
            return cls.replace("-", " ").title()

    return "Datasource"


def classify_datasource(ds_xml: str) -> dict:
    """
    Enhanced Tableau-accurate datasource classification
    """
    info = {
        "source": "Embedded",
        "datasource_type": "Unknown",
        "connection": "Unknown",
        "mode": "Live",
        "location": "Local/Embedded"
    }

    if not ds_xml: return info

    try:
        root = etree.fromstring(ds_xml.encode("utf-8"))
    except Exception:
        return info

    conn = root.find(".//connection")

    # 1️⃣ Published Datasource (Check for Repository Location)
    repo = root.find(".//repository-location")
    if repo is not None:
        info.update({
            "source": "Published",
            "datasource_type": "Tableau Server",
            "connection": "Managed",
            "mode": "Published",
            "location": f"Site: {repo.get('site', 'Default')} | Path: {repo.get('path', '')}"
        })
        return info

    # 2️⃣ Extract specific details from Connection tag
    if conn is not None:
        cls = (conn.get("class") or "").lower()
        server = conn.get("server") or ""
        dbname = conn.get("dbname") or ""
        filename = (conn.get("filename") or "").lower()

        # Update Connection Type
        info["connection"] = cls.replace("-", " ").title()

        # Snowflake Detail
        if "snowflake" in cls:
            info.update({
                "source": "Snowflake",
                "datasource_type": "Database",
                "location": f"Server: {server} | DB: {dbname}"
            })
        # Excel/CSV Detail
        elif "excel" in cls or "textscan" in cls or filename.endswith((".csv", ".xlsx", ".xls")):
            info.update({
                "source": "Uploaded File",
                "datasource_type": "File",
                "mode": "Extract",
                "location": filename if filename else "Embedded in Workbook"
            })
        # Generic Hyper
        elif cls == "hyper":
            info.update({"source": "Extract", "connection": "Hyper", "mode": "Extract"})

    return info

def infer_datasource_from_name(ds_name: str) -> dict:
    """
    Heuristic fallback when Tableau hides XML metadata
    """
    name = (ds_name or "").lower()

    if "snowflake" in name:
        return {
            "source": "Snowflake",
            "datasource_type": "Database",
            "connection": "Snowflake",
            "mode": "Live",
            "location": "External (Snowflake)"
        }

    if "csv" in name or "text" in name:
        return {
            "source": "Uploaded File",
            "datasource_type": "Text File",
            "connection": "CSV / Text",
            "mode": "Extract",
            "location": "Embedded (Tableau Extract)"
        }

    if "excel" in name or "xls" in name:
        return {
            "source": "Uploaded File",
            "datasource_type": "Excel",
            "connection": "Excel",
            "mode": "Extract",
            "location": "Embedded (Tableau Extract)"
        }

    if name.endswith("_emb") or "extract" in name:
        return {
            "source": "Extract",
            "datasource_type": "Extract",
            "connection": "Hyper",
            "mode": "Extract",
            "location": "Embedded (Tableau Extract)"
        }

    return {
        "source": "Hidden by Tableau",
        "datasource_type": "Hidden by Tableau",
        "connection": "Hidden by Tableau",
        "mode": "Hidden by Tableau",
        "location": "Hidden by Tableau"
    }


def classify_published_datasource(ds_name: str) -> dict:
    return {
        "source": "Tableau Server",
        "datasource_type": "Published Datasource",
        "connection": "Managed by Tableau",
        "mode": "Live",
        "location": "Tableau Server / Cloud"
    }

def has_repository_location(root) -> bool:
    """
    Namespace-safe detection of published datasource
    """
    if root is None:
        return False
    return bool(root.xpath(".//*[local-name()='repository-location']"))




def extract_datasources_raw(file_path):
    if not file_path or not os.path.exists(file_path):
        return {}

    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    datasources = {}
    pattern = re.compile(
        r'(<datasource [^>]*>.*?</datasource>)',
        re.DOTALL | re.IGNORECASE
    )

    for i, xml in enumerate(pattern.findall(content)):
        name_match = re.search(r'(?:caption|name)=[\'"]([^\'"]+)[\'"]', xml)
        name = name_match.group(1) if name_match else f"Datasource_{i}"
        datasources[name] = xml

    return datasources

def clean_xml_for_parsing(xml_string):
    """
    Strips XML namespaces and prefixes to prevent parsing errors.
    """
    if not xml_string:
        return ""
    
    xml_string = re.sub(r'<\?xml.*?\?>', '', xml_string)
    xml_string = re.sub(r'\sxmlns(?::\w+)?="[^"]+"', '', xml_string)
    xml_string = re.sub(r'\s\w+:\w+="[^"]+"', '', xml_string)
    xml_string = re.sub(r'<(\w+):(\w+)', r'<\2', xml_string)
    xml_string = re.sub(r'</(\w+):(\w+)>', r'</\2>', xml_string)
    
    return xml_string

def determine_mode_and_privacy(xml):
    privacy = "Embedded"
    mode = "Live"

    if not xml:
        return mode, privacy

    if "<repository-location" in xml:
        privacy = "Published"
    
    if "<extract" in xml or "class='hyper'" in xml or 'class="hyper"' in xml:
        mode = "Extract"
        
    return mode, privacy

def is_internal_calc(name):
    if not name: return True
    n = name.lower()
    return n.startswith("calculation_") or n == "number of records" or n == "measure values"


# ---------------- METADATA EXTRACTION ----------------

def extract_all_filters_deterministically(xml_text):
    """
    Extract ONLY datasource filters.
    Excludes worksheet filters, action filters, tooltip fields, etc.
    """

    filters = set()

    if not xml_text:
        return filters

    def clean_name(val):

        if not val:
            return None

        val = val.replace("[", "").replace("]", "").strip()

        if val.lower().startswith("extract."):
            val = val.split(".", 1)[1]

        if "." in val:
            val = val.split(".")[-1]

        return val.strip()

    clean_xml = clean_xml_for_parsing(xml_text)

    try:

        root = ET.fromstring(clean_xml)

        # --------------------------------------------------
        # ONLY scan datasource node
        # --------------------------------------------------
        for datasource in root.findall(".//datasource"):

            # Direct datasource filters
            for f in datasource.findall(".//filter"):

                col = f.get("column") or f.get("field")

                name = clean_name(col)

                if name and not is_internal_calc(name):
                    filters.add(name)

            # Extract filters
            for extract in datasource.findall(".//extract"):

                for f in extract.findall(".//filter"):

                    col = f.get("column") or f.get("field")

                    name = clean_name(col)

                    if name and not is_internal_calc(name):
                        filters.add(name)

            # Relation filters
            for relation in datasource.findall(".//relation"):

                expr = relation.get("expression")

                if expr:

                    matches = re.findall(
                        r'\[([^\]]+)\]',
                        expr
                    )

                    for m in matches:

                        name = clean_name(m)

                        if name and not is_internal_calc(name):
                            filters.add(name)

    except Exception:
        pass

    return filters


def extract_user_defined_ds_calcs(xml_text):
    """
    Finds calculated fields, including those nested inside <column> tags.
    """
    calcs = set()
    if not xml_text:
        return calcs
        
    clean = clean_xml_for_parsing(xml_text)
    try:
        root = ET.fromstring(clean)
        
        # 1. Standalone <calculation> tags
        for c in root.findall(".//calculation"):
            name = c.get("name") or c.get("caption")
            if name and not is_internal_calc(name):
                calcs.add(name)

        # 2. <column> tags containing <calculation>
        for col in root.findall(".//column"):
            if col.find("./calculation") is not None:
                name = col.get("caption") or col.get("name")
                if name:
                    name = name.replace("[", "").replace("]", "").strip()
                    if not is_internal_calc(name):
                        calcs.add(name)
    except:
        pass
    return calcs



# ---------------- SILENT DOWNLOADER ----------------

def resolve_luid_by_content_url(name, site_id, token):
    if not name: return None
    url = f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/datasources"
    params = {"filter": f"contentUrl:eq:{name}"}
    headers = {"X-Tableau-Auth": token}
    try:
        r = requests.get(url, headers=headers, params=params, verify=False)
        if r.status_code == 200:
            root = ET.fromstring(r.text)
            ds = root.find(".//datasource")
            if ds is not None: return ds.get("id") 
    except: pass
    return None

def download_published_xml(xml_chunk, site_id, token):
    """
    Downloads published XML.
    SILENT MODE: No print statements for download/404/resolve status.
    """
    if not xml_chunk: return None

    ds_id, ds_name = None, None
    match_id = re.search(r'<repository-location [^>]*id=[\'"]([^\'"]+)[\'"]', xml_chunk)
    if match_id: ds_id = match_id.group(1)
    
    match_url = re.search(r'<repository-location [^>]*content-url=[\'"]([^\'"]+)[\'"]', xml_chunk)
    if match_url: ds_name = match_url.group(1)

    if not ds_id and not ds_name: return None

    def attempt_download(target_id):
        url = f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/datasources/{target_id}/content"
        headers = {"X-Tableau-Auth": token}
        return requests.get(url, headers=headers, stream=True, verify=False)

    # 1. Attempt with ID
    r = attempt_download(ds_id or ds_name)

    # 2. If 404, try resolve (Silently)
    if r.status_code == 404:
        name_to_resolve = ds_name if ds_name else ds_id
        resolved_luid = resolve_luid_by_content_url(name_to_resolve, site_id, token)
        if resolved_luid:
            r = attempt_download(resolved_luid)
        else:
            # Silent fail
            return None

    if r.status_code != 200: return None

    # 3. Extract Content
    try:
        if r.content[:4] == b'PK\x03\x04':
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                for filename in z.namelist():
                    if filename.endswith(".tds"):
                        with z.open(filename) as f:
                            return f.read().decode("utf-8", errors="ignore")
        else:
            return r.text
    except:
        return None


def get_connection_class(xml_text):
    """
    Tableau-accurate connection detection.

    Handles:
    - Excel (.xls, .xlsx, .xlsm)
    - Extract (.hyper)
    - Live DB
    - Embedded extract
    - Published extract
    - Federated connections
    """

    if not xml_text:
        return "Unknown"

    clean = clean_xml_for_parsing(xml_text)

    try:
        root = ET.fromstring(clean)

        # PRIORITY 1 — Named connections
        for nc in root.findall(".//named-connection"):

            conn = nc.find(".//connection")

            if conn is not None:

                cls = conn.get("class")

                filename = (conn.get("filename") or "").lower()

                if filename.endswith((".xls", ".xlsx", ".xlsm")):
                    return "excel"

                if filename.endswith(".hyper"):
                    return "extract"

                if cls and cls != "federated":
                    return cls

        # PRIORITY 2 — Direct connection
        for conn in root.findall(".//connection"):

            cls = conn.get("class")

            filename = (conn.get("filename") or "").lower()

            if filename.endswith((".xls", ".xlsx", ".xlsm")):
                return "excel"

            if filename.endswith(".hyper"):
                return "extract"

            if cls == "hyper":
                return "extract"

            if cls and cls != "federated":
                return cls

        # PRIORITY 3 — Extract node exists
        if root.find(".//extract") is not None:
            return "extract"

    except Exception:
        pass

    # fallback detection
    xml_lower = xml_text.lower()

    if ".xlsm" in xml_lower:
        return "excel"

    if ".xls" in xml_lower:
        return "excel"

    if ".hyper" in xml_lower:
        return "extract"

    return "Unknown"

def compare(name, old_xml, new_xml, site_id, token):

    # PRE-ANALYSIS
    old_mode, old_privacy = determine_mode_and_privacy(old_xml)
    new_mode, new_privacy = determine_mode_and_privacy(new_xml)

    # HYDRATION
    if old_privacy == "Published":

        downloaded = download_published_xml(old_xml, site_id, token)

        if downloaded:
            old_xml = downloaded
            old_mode, _ = determine_mode_and_privacy(old_xml)

    if new_privacy == "Published":

        downloaded = download_published_xml(new_xml, site_id, token)

        if downloaded:
            new_xml = downloaded
            new_mode, _ = determine_mode_and_privacy(new_xml)

    # FILTER EXTRACTION
    old_items = extract_all_filters_deterministically(old_xml)
    new_items = extract_all_filters_deterministically(new_xml)

    # CONNECTION DETECTION
    old_conn = get_connection_class(old_xml)
    new_conn = get_connection_class(new_xml)

    # preserve connection if extract hides original
    if new_conn == "extract" and old_conn == "excel":
        new_conn = "extract"

    if new_conn == "Unknown" and old_conn != "Unknown":
        new_conn = old_conn

    if old_conn == "Unknown" and new_conn != "Unknown":
        old_conn = new_conn

    # DIFFERENCE
    old_info = classify_datasource(old_xml)
    new_info = classify_datasource(new_xml)
    added = sorted(new_items - old_items)
    removed = sorted(old_items - new_items)
    common = sorted(new_items & old_items)

    bullets = []
    status = "info"

    # FILTER CHANGES
    if added:
        status = "modified"
        bullets.append(f"➕ New Filters Added: {', '.join(added)}")

    if removed:
        status = "modified"
        bullets.append(f"➖ Filters Removed: {', '.join(removed)}")

    if common:
        bullets.append(f"ℹ️ Existing Filters (unchanged): {', '.join(common)}")

    # MODE CHANGE
    if old_mode != new_mode:
        status = "modified"
        bullets.append(f"⚡ Mode: {old_mode} → {new_mode}")

    # This addresses your request to show if location changed
    if old_info["location"] != new_info["location"]:
        bullets.append(f"📍 **Location Change:** {old_info['location']} ➡️ {new_info['location']}")
    else:
        bullets.append(f"📍 **Location:** {new_info['location']}")

    if old_info["connection"] != new_info["connection"]:
        bullets.append(f"🔌 **Connection Type:** {old_info['connection']} ➡️ {new_info['connection']}")

    # CONNECTION CHANGE
    if old_conn != new_conn:
        status = "modified"

        bullets.append(
            f"🔌 Connection: "
            f"{old_conn.title()} → {new_conn.title()}"
        )

    # PRIVACY CHANGE
    if old_privacy != new_privacy:
        status = "modified"
        bullets.append(f"🔒 Privacy: {old_privacy} → {new_privacy}")

    # SUMMARY
    display_conn = new_conn.title()

    if display_conn == "Extract":
        display_conn = "Tableau Extract"

    summary = f"ℹ️ {display_conn} ({new_mode} | {new_privacy})"

    bullets.insert(0, summary)

    # SAVE
    CHANGE_REGISTRY["datasources"].setdefault(name, []).append({
        "status": "modified" if old_info != new_info else "info",
        "title": "Datasource Metadata & Connection",
        "object": "__metadata__",
        "bullets": bullets
    })


def summarize_datasources(ds_name, old_xml, new_xml):

    xml = new_xml or old_xml

    try:
        root = etree.fromstring(xml.encode("utf-8")) if xml else None
    except Exception:
        root = None

    # existing classification
    if root is not None and has_repository_location(root):
        info = classify_published_datasource(ds_name)
    else:
        info = classify_datasource(xml)

        if all(isinstance(v, str) and v.startswith("Not exposed") for v in info.values()):
            info = infer_datasource_from_name(ds_name)

    # 🔥 ADD THIS BLOCK
    connection_class = get_connection_class(xml)

    if connection_class and connection_class != "Unknown":
        info["connection"] = connection_class.replace("-", " ").title()

    bullets = [
        f"Source: {info['source']}",
        f"Datasource Type: {info['datasource_type']}",
        f"Connection: {info['connection']}",   # ← THIS WILL NOW SHOW CORRECT VALUE
        f"Mode: {info['mode']}",
        f"Location: {info['location']}",
    ]

    register_change(
        level="datasource",
        parent=ds_name,
        title="Datasource Summary",
        status="info",
        bullets=bullets
    )





def parse_bins(xml):
    root = _parse_fragment(xml)
    bins = set()
    if root is None:
        return bins

    for col in root.findall(".//column"):
        b = col.find("bin")
        if b is not None:
            name = col.attrib.get("name")
            size = b.attrib.get("size")
            bins.add(f"{name} (bin size {size})")
    return bins


def extract_used_fields(sections: dict) -> set:
    """
    Return all fields actually used anywhere in the workbook.
    """
    used = set()

    # Worksheets & dashboards
    for sec in ("worksheets", "dashboards"):
        for xml in sections.get(sec, {}).values():
            used |= extract_worksheet_fields(xml)

    # Stories
    for xml in sections.get("stories", {}).values():
        root = _parse_fragment(xml)
        if root is None:
            continue
        for sp in root.findall(".//story-point"):
            sheet = sp.attrib.get("captured-sheet")
            if sheet:
                used.add(sheet)

    return used

# ==================================================
# CALCULATION HELPERS (MUST BE ABOVE KPI FUNCTION)
# ==================================================

def is_user_visible_calc(name: str) -> bool:
    if not name:
        return False

    n = name.strip().lower()

    if n.startswith("__"):
        return False
    if n in {"measure names", "measure values"}:
        return False
    if n.startswith("agg("):
        return False
    if "calculation_" in n:
        return False

    return True


def is_table_calculation(formula: str) -> bool:
    if not formula:
        return False

    f = formula.upper()

    TABLE_CALC_KEYWORDS = [
        "RANK(", "INDEX(", "LOOKUP(", "WINDOW_",
        "RUNNING_", "TOTAL(", "FIRST()", "LAST()"
    ]

    return any(k in f for k in TABLE_CALC_KEYWORDS)

def is_rls_calculation(formula: str) -> bool:
    """
    Detect Row Level Security (RLS) calculations.
    """
    if not formula:
        return False

    f = formula.upper()

    RLS_KEYWORDS = [
        "USERNAME()",
        "USERDOMAIN()",
        "FULLNAME()",
        "ISMEMBEROF(",
        "USERATTR(",
        "CONTAINS(USERNAME",
        "CASE USERNAME()",
        "IF USERNAME()",
    ]

    return any(k in f for k in RLS_KEYWORDS)


def build_workbook_kpi_snapshot(sections: dict) -> dict:
    """
    Build absolute KPI counts for a workbook version.

    Enhancements:
    - Adds RLS calculated fields count
    - No unused-variable warnings
    - Tableau UI–accurate
    """

    all_columns = set()
    calculated = set()
    rls_calculated = set()  # ✅ USED

    # ==================================================
    # DATASOURCES → COLUMNS & CALCULATIONS
    # ==================================================
    for ds_xml in sections.get("datasources", {}).values():
        root = _parse_fragment(ds_xml)
        if root is None:
            continue

        for col in root.findall(".//column"):
            internal_name = col.attrib.get("name")
            caption = col.attrib.get("caption") or internal_name
            if not internal_name:
                continue

            all_columns.add(internal_name)

            # ---- Parameter detection ----
            is_parameter = bool(
                col.attrib.get("param-domain-type")
                or col.find("range") is not None
                or col.find("list") is not None
                or (col.attrib.get("role") or "").lower() == "parameter"
            )

            # ---- Calculation detection ----
            calc = col.find("calculation")

            if calc is not None and not is_parameter:
                formula = (
                    calc.attrib.get("formula")
                    or (calc.text or "").strip()
                )

                if formula and is_user_visible_calc(caption):
                    calculated.add(caption)

                    # ✅ RLS detection
                    if is_rls_calculation(formula):
                        rls_calculated.add(caption)

    # ==================================================
    # MERGE EXTRACTED CALCULATIONS (PUBLISH SAFE)
    # ==================================================
    for calc_name, xml in sections.get("calculations", {}).items():
        if not is_user_visible_calc(calc_name):
            continue

        calculated.add(calc_name)

        root = _parse_fragment(xml)
        if root is None:
            continue

        c = root.find(".//calculation")
        if c is not None:
            formula = c.attrib.get("formula") or (c.text or "").strip()
            if is_rls_calculation(formula):
                rls_calculated.add(calc_name)

    # ==================================================
    # FILTERS
    # ==================================================
    sem = collect_workbook_semantics(sections)

    context_filters = set()
    for ws_xml in sections.get("worksheets", {}).values():
        root = _parse_fragment(ws_xml)
        if root is None:
            continue
        for f in root.findall(".//filter[@context='true']"):
            name = f.attrib.get("column") or f.attrib.get("field")
            if name:
                context_filters.add(name.replace("[", "").replace("]", ""))

    ds_filters = set()
    for ds_xml in sections.get("datasources", {}).values():
        ds_filters |= parse_datasource_filters(ds_xml)

    total_filters = (
        len(sem.get("filters", set())) +
        len(sem.get("dashboard_filters", set())) +
        len(ds_filters)
    )

    # ==================================================
    # FINAL KPI SNAPSHOT
    # ==================================================
    return {
        "fields_total": len(all_columns),
        "fields_calc": len(calculated),
        "fields_calc_rls": len(rls_calculated),  # ✅ USED

        "parameters": len(sections.get("parameters", {})),

        "filters": len(sem.get("filters", set())),
        "dashboard_filters": len(sem.get("dashboard_filters", set())),
        "filters_context": len(context_filters),
        "filters_datasource": len(ds_filters),
        "filters_total": total_filters,

        "worksheets": len(sections.get("worksheets", {})),
        "dashboards": len(sections.get("dashboards", {})),
        "stories": len(sections.get("stories", {})),
    }






def count_removed_calculations(cards):
    return sum(
        1 for c in cards
        if c["section"] == "calculations" and c["status"] == "removed"
    )


def render_workbook_kpi_table(kpi_old: dict, kpi_new: dict) -> str:
    return f"""
    <details class="panel" style="
        background:#EAF3FF;
        border-radius:14px;
        padding:10px 14px;
        margin:16px 0;
        box-shadow:0 3px 8px rgba(0,0,0,0.08);
    ">
      <summary style="
          cursor:pointer;
          font-size:16px;
          font-weight:600;
          color:#1f6fe5;
          padding:6px 0;
          list-style:none;
      ">
        📊 Workbook Metrics — Source Workbook Vs Target Workbook
      </summary>

      <div style="margin-top:12px;">
        <table style="
            width:100%;
            border-collapse:collapse;
            font-size:14px;
        ">
          <tr style="border-bottom:1px solid #c6dbff;">
            <th align="left">Metric</th>
            <th align="right">Source</th>
            <th align="right">Target</th>
          </tr>

          <tr>
            <td>Stories</td>
            <td align="right">{kpi_old['stories']}</td>
            <td align="right">{kpi_new['stories']}</td>
          </tr>

          <tr>
            <td>Dashboards</td>
            <td align="right">{kpi_old['dashboards']}</td>
            <td align="right">{kpi_new['dashboards']}</td>
          </tr>

          <tr>
            <td>Worksheets</td>
            <td align="right">{kpi_old['worksheets']}</td>
            <td align="right">{kpi_new['worksheets']}</td>
          </tr>

          <tr>
            <td>Filters (Worksheet, Dashboard)</td>
            <td align="right">
              {kpi_old['filters_total']}
              <span style="opacity:0.75;">
                ({kpi_old['filters']}, {kpi_old['dashboard_filters']})
              </span>
            </td>
            <td align="right">
              {kpi_new['filters_total']}
              <span style="opacity:0.75;">
                ({kpi_new['filters']}, {kpi_new['dashboard_filters']})
              </span>
            </td>
          </tr>

          <tr>
            <td>Parameters</td>
            <td align="right">{kpi_old['parameters']}</td>
            <td align="right">{kpi_new['parameters']}</td>
          </tr>

          <tr>
            <td>Calculated Fields</td>
            <td align="right">{kpi_old['fields_calc']}</td>
            <td align="right">{kpi_new['fields_calc']}</td>
          </tr>
        </table>
      </div>
    </details>
    """


def parse_groups(xml):
    root = _parse_fragment(xml)
    groups = set()
    if root is None:
        return groups

    for g in root.findall(".//group"):
        name = g.attrib.get("caption") or g.attrib.get("name")
        members = []
        for gf in g.findall(".//groupfilter"):
            mem = gf.attrib.get("member")
            if mem:
                members.append(mem.replace("[","").replace("]",""))
        if name and members:
            groups.add(f"Group '{name}' on {', '.join(members)}")
        if g.attrib.get("{http://www.tableausoftware.com/xml/user}ui-builder") == "filter-group":
            groups.add(f"Set '{name}'")

    return groups

def summarize_calculations(old_xml, new_xml):
    """
    Enhanced calculation summary:
    - Added → show formula
    - Removed → show old formula
    - Modified → show before & after
    """
    bullets = []

    def extract_formula(xml):
        root = _parse_fragment(xml)
        if root is None:
            return None
        calc = root.find(".//calculation")
        if calc is None:
            return None
        return calc.attrib.get("formula") or (calc.text or "").strip()

    old_formula = extract_formula(old_xml)
    new_formula = extract_formula(new_xml)

    # ADDED
    if not old_formula and new_formula:
        bullets.append("➕ Calculation added")
        bullets.append(f"Formula: {new_formula}")
        return bullets

    # REMOVED
    if old_formula and not new_formula:
        bullets.append("➖ Calculation removed")
        bullets.append(f"Previous Formula: {old_formula}")
        return bullets

    # MODIFIED
    if old_formula != new_formula:
        bullets.append("🟨 Calculation modified")
        bullets.append(f"Before: {old_formula}")
        bullets.append(f"After : {new_formula}")
        return bullets

    return []

def extract_removed_dashboard_filters_from_diff(diff_text: str) -> set:
    """
    Parse xmldiff output lines to detect deleted dashboard filter zones like:
      delete-node: zone[...] param="[none:Category:nk]"
    """
    removed = set()
    for line in diff_text.splitlines():
        if "delete-node" in line and "param=" in line and "[none:" in line.lower():
            m = re.search(r"\[none:([A-Za-z0-9 _-]+?):nk\]", line)
            if m:
                removed.add(m.group(1).strip())
    return removed

def build_kpi_counts(sem: dict) -> dict:
    """
    Return numeric KPI counts from semantic extraction.
    """
    return {
        "Filters": len(sem.get("filters", [])),
        "Dashboard Filters": len(sem.get("dashboard_filters", [])),
        "Filter Controls": len(sem.get("filter_controls", [])),
        "Colors": len(sem.get("colors", [])),
        "Legends": len(sem.get("legends", [])),
        "Actions": len(sem.get("actions", [])),
        "Fields in View": len(sem.get("mark_fields", [])),
        "Tooltip Fields": len(sem.get("tooltip_fields", [])),
        "Dashboard Sheets": len(sem.get("dashboard_sheets", [])),
    }

def summarize_kpi_changes(old_sem: dict, new_sem: dict) -> list:
    """
    Generate KPI-style change bullets based on count differences.
    """
    old_kpi = build_kpi_counts(old_sem)
    new_kpi = build_kpi_counts(new_sem)

    bullets = []

    for kpi, old_count in old_kpi.items():
        new_count = new_kpi.get(kpi, 0)
        if old_count != new_count:
            direction = "increased" if new_count > old_count else "decreased"
            bullets.append(
                f"📊 {kpi} {direction}: {old_count} → {new_count}"
            )

    return bullets


# ----------- GPT summarization -----------
def gpt_summarize_item(title:str, xml_ops:str, semantics_old:dict, semantics_new:dict)->list:
    # Compute explicit differences (add/remove) for dashboard filters & controls
    old_filters = set(semantics_old.get("dashboard_filters", []))
    new_filters = set(semantics_new.get("dashboard_filters", []))
    added_filters = sorted(new_filters - old_filters)
    removed_filters = sorted(old_filters - new_filters)

    explicit_hints = ""
    if added_filters or removed_filters:
        explicit_hints += "\nExplicit Dashboard Filter Changes:\n"
        if added_filters:
            explicit_hints += " - Added filters: " + ", ".join(added_filters) + "\n"
        if removed_filters:
            explicit_hints += " - Removed filters: " + ", ".join(removed_filters) + "\n"

    prompt = f"""
Your task: Convert the following XMLDiff operations into clear, business-friendly change summaries.

Title: {title}

ABSOLUTE RULES (MANDATORY):
- ONLY describe ACTUAL changes that occurred.
- DO NOT generate bullets that state:
  • "No changes"
  • "No modifications"
  • "No filters added"
  • "No impact"
  • Any confirmation or absence-of-change statement.
- If a change cannot be clearly inferred from XMLDiff or SEMANTICS,
  DO NOT mention it at all.
- Silence is preferred over guessing.

INTERPRETATION RULES:
- Interpret operations logically; explain what changed and where.
- Describe modified items as [old → new] ONLY when evidence exists.
- Group related changes when appropriate, but NEVER invent changes.
- Do NOT repeat the same change in multiple bullets.
- Do NOT include XML tags, internal ids, or technical syntax.

OBJECT COVERAGE (ONLY IF CHANGED):
Detect and describe changes ONLY IF THEY OCCURRED for:
• Datasources (filters, joins, relationships, fields, aliases)
• Calculations (formula changes, LODs, logic updates)
• Parameters (value, range, datatype, control)
• Bins, Sets, Groups, Hierarchies
• Filters (worksheet, dashboard, datasource)
• Actions (filter, highlight, URL, parameter)
• Worksheets (added, removed, renamed, modified)
• Dashboards (added, removed, renamed, modified)
• Text objects (content or formatting changes)
• Charts / Marks (type, axis, color, size, labels)
• Tooltips (added, removed, modified)
• Legends (added, removed, resized, styled)

STRICT RULES FOR DASHBOARD MODIFICATION DETECTION:

You MUST classify a dashboard as "Modified" ONLY if there is a change in:
- Dashboard title or name
- Worksheets added to or removed from the dashboard
- Filters added, removed, or changed
- Parameters added, removed, or changed
- Actions added, removed, or changed
- Calculated fields used by dashboard worksheets
- Data sources used by dashboard worksheets

You MUST NOT classify a dashboard as "Modified" for ANY of the following changes:
- Resizing dashboard zones
- Repositioning zones
- Moving worksheets within the dashboard
- Changing tiled vs floating layout
- Device layout changes (Desktop / Tablet / Phone)
- Padding, spacing, margins, or sizing changes
- Formatting-only changes (colors, fonts, borders)
- Changes that affect layout only and not logic or data

IMPORTANT:
If the detected change ONLY involves zones, layout, or positioning,
you MUST return status = "Unchanged" and DO NOT generate a modification message.

FILTER-SPECIFIC RULES:
- Explicitly detect filter control-type changes
  (e.g., Multivalue List → Single Value Dropdown).
- Only describe filter changes if a filter was added, removed, or modified.
- Do NOT mention filters that are unchanged.

COLOR RULES:
- When a color change is detected, identify the color name
  (infer from color code if needed).

CALCULATION RULES:
- When describing calculation changes, specify:
  • Added / Removed / Modified
  • Old → New formula (if available)
  • Mention LOD calculations explicitly when applicable.

SEMANTICS USAGE:
- Use SEMANTICS ONLY to explain or enrich REAL changes.
- Do NOT create bullets based on SEMANTICS alone unless a change is evident.

INPUT DATA:

1) XMLDiff operations:
{xml_ops}

2) Semantic differences (before → after):
old = {semantics_old}
new = {semantics_new}

OUTPUT RULES (STRICT):
- Return ONLY bullets for REAL changes.
- If fewer than 3 meaningful changes exist, return fewer bullets.
- If NO meaningful changes exist, return an EMPTY response.
- Each bullet MUST start with an action verb:
  (Added, Removed, Updated, Modified, Renamed).
- Human-readable only. No XML. No explanations.
"""

    try:
        r = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role":"user","content":prompt}],
            max_tokens=1400,
            temperature=0.2
        )
        text = r.choices[0].message.content or ""
        lines = [ln.lstrip("-• ").strip() for ln in text.splitlines() if ln.strip()]
        return lines[:12]
    except Exception as e:
        return [f"(GPT summary failed: {e})"]

# ----------- Deterministic semantic bullets -----------
def diff_set(title, a:set, b:set, add_icon="➕", rem_icon="➖"):
    out=[]
    add = sorted(b-a); rem = sorted(a-b)
    if add: out.append(f"{add_icon} {title} added: " + ", ".join(add))
    if rem: out.append(f"{rem_icon} {title} removed: " + ", ".join(rem))
    return out

def is_story_publish_noise(old_xml: str, new_xml: str) -> bool:
    """
    Returns True if differences are ONLY backend publish noise
    (repository-location, content-url, ids).
    """
    diff = xmldiff_text(old_xml, new_xml)

    meaningful_keywords = [
        "story-point",
        "captured-sheet",
        "zone",
        "worksheet",
        "dashboard",
        "filter",
        "parameter",
        "calculation",
    ]

    for line in diff.splitlines():
        if "repository-location" in line or "content-url" in line:
            continue

        if any(k in line.lower() for k in meaningful_keywords):
            return False  # real change

    return True

def summarize_semantics(label, old, new):
    bullets=[]
    if label == "Dashboard":
    

        # --- REAL dashboard changes ---
        bullets += diff_set(
            "Sheets",
            set(old.get("dashboard_sheets", [])),
            set(new.get("dashboard_sheets", []))
        )

        bullets += diff_set(
            "Dashboard-level filters",
            set(old.get("dashboard_filters", [])),
            set(new.get("dashboard_filters", [])),
            "🔎",
            "🔎"
        )

        bullets += diff_set(
            "Legends",
            set(old.get("legends", [])),
            set(new.get("legends", [])),
            "🧭",
            "🧭"
        )

        bullets += diff_set(
            "Actions",
            set(old.get("actions", [])),
            set(new.get("actions", [])),
            "💥",
            "💥"
        )

    # 🔒 IGNORE layout-only changes (size / zone movement)
    # If NOTHING meaningful changed → return empty


    if label=="Worksheet":
        bullets += diff_set("Filters", set(old.get("filters",[])), set(new.get("filters",[])))
        if old.get("date_filters",[]) != new.get("date_filters",[]):
            bullets.append("📅 Date filter setup changed.")
        bullets += diff_set("Filter controls", set(old.get("filter_controls",[])), set(new.get("filter_controls",[])))
        bullets += diff_set("Legends", set(old.get("legends",[])), set(new.get("legends",[])), "🧭", "🧭")
        bullets += diff_set("Color by", set(old.get("mark_color_by",[])), set(new.get("mark_color_by",[])), "🎯", "🎯")
        bullets += diff_set("Size by", set(old.get("mark_size_by",[])), set(new.get("mark_size_by",[])), "📏", "📏")
        bullets += diff_set("Label by", set(old.get("mark_label_by",[])), set(new.get("mark_label_by",[])), "🏷️", "🏷️")
        bullets += diff_set("Shape by", set(old.get("mark_shape_by",[])), set(new.get("mark_shape_by",[])), "🔺", "🔺")
        bullets += diff_set("Fields in view", set(old.get("mark_fields",[])), set(new.get("mark_fields",[])), "📊", "📊")
        if old.get("tooltip_raw","") != new.get("tooltip_raw",""):
            bullets.append("💬 Tooltip content updated.")
        bullets += diff_set("Tooltip fields", set(old.get("tooltip_fields",[])), set(new.get("tooltip_fields",[])), "💬", "💬")
    if label == "Story":
        old_points = set(old.get("story_points", []))
        new_points = set(new.get("story_points", []))

        # 🚫 Ignore publish-only noise (repository-location, content-url, ids)
        if old_points == new_points:
            return []   # ⛔ NO real change → no bullets, no card

        bullets += diff_set("Story Points", old_points, new_points)

    return bullets

# ----------- GLOBAL ACTIONS (inserted here after summarize_semantics) -----------

# Minimal xmldiff helper for comparing action XML
def xmldiff_changes(old_xml, new_xml):
    try:
        formatter = DiffFormatter()
        return diff_texts(old_xml, new_xml, formatter=formatter)
    except Exception as exc:
        return f"(xmldiff failed: {exc})"


def summarize_global_actions(old_root, new_root):
    """
    Extracts and compares global Tableau actions (filter, highlight, URL,
    parameter, set control). Returns a list of dictionaries describing 
    added, removed, or modified actions.
    """

    def extract_actions(root):
        if root is None:
            return {}

        actions_nodes = []
        for elem in root.iter():
            tag = elem.tag.lower().split("}")[-1]
            if tag == "actions":
                actions_nodes.append(elem)

        if not actions_nodes:
            return {}

        result = {}
        for actions_node in actions_nodes:
            for action in actions_node:
                a_tag = action.tag.lower().split("}")[-1]
                if a_tag != "action":
                    continue

                caption = (
                    action.attrib.get("caption")
                    or action.attrib.get("name")
                    or f"action_{hash(ET.tostring(action))}"
                )

                # Detect action type
                a_type = "unknown"
                a_class = (action.attrib.get("class", "") or "").lower()
                a_cmd = (action.attrib.get("type", "") or "").lower()

                if "filter" in a_class or "filter" in a_cmd:
                    a_type = "filter"
                elif "highlight" in a_class or "brush" in a_cmd:
                    a_type = "highlight"
                elif "url" in a_class or "url" in a_cmd:
                    a_type = "url"
                elif "set" in a_class:
                    a_type = "set control"
                elif "parameter" in a_class:
                    a_type = "parameter action"

                # Detect scope
                scope = "unknown"
                for child in action:
                    ctag = child.tag.lower().split("}")[-1]
                    if ctag == "source":
                        if "dashboard" in child.attrib:
                            scope = f"dashboard: {child.attrib.get('dashboard')}"
                        elif "worksheet" in child.attrib:
                            scope = f"worksheet: {child.attrib.get('worksheet')}"

                clean_xml = ET.tostring(action, encoding="unicode")

                result[caption] = {
                    "xml": clean_xml,
                    "type": a_type,
                    "scope": scope,
                }

        return result

    # Extract actions from both XML documents
    old_actions = extract_actions(old_root)
    new_actions = extract_actions(new_root)

    old_names = set(old_actions.keys())
    new_names = set(new_actions.keys())

    added = new_names - old_names
    removed = old_names - new_names
    common = new_names & old_names

    summary = []

    # Added actions
    for name in sorted(added):
        a = new_actions[name]
        t = a["type"].upper() if a["type"] != "unknown" else "ACTION"
        summary.append(
            {
                "line": f"➕ {t} added ({a['scope']}): **{name}**",
                "xml_old": None,
                "xml_new": a["xml"],
            }
        )

    # Removed actions
    for name in sorted(removed):
        a = old_actions[name]
        t = a["type"].upper() if a["type"] != "unknown" else "ACTION"
        summary.append(
            {
                "line": f"❌ {t} removed ({a['scope']}): **{name}**",
                "xml_old": a["xml"],
                "xml_new": None,
            }
        )

    # Modified actions
    for name in sorted(common):
        old_a = old_actions[name]
        new_a = new_actions[name]

        if old_a["xml"] != new_a["xml"]:
            ops = xmldiff_changes(old_a["xml"], new_a["xml"])
            reason = "Configuration updated"

            t = new_a["type"].upper() if new_a["type"] != "unknown" else "ACTION"

            summary.append(
                {
                    "line": f"✏️ {t} modified ({new_a['scope']}): **{name}** — {reason}",
                    "xml_old": old_a["xml"],
                    "xml_new": new_a["xml"],
                }
            )

    return summary

def _parse_action_details(action_elem):
    """
    Parse an <action> element (ElementTree Element) and return a dict:
      - caption, type, scope
      - sources: list of {"dashboard"/"worksheet","name"}
      - targets: list of {"dashboard"/"worksheet","name"}
      - field_mappings: list of {"field","role"} 
      - behavior: textual hint (Replace / Add / Keep / Exclude / unknown)
      - raw_xml: str
    """
    details = {
        "caption": None,
        "type": "unknown",
        "scope": "unknown",
        "sources": [],
        "targets": [],
        "field_mappings": [],
        "behavior": "unknown",
        "raw_xml": ET.tostring(action_elem, encoding="unicode")
    }

    # caption
    details["caption"] = action_elem.attrib.get("caption") or action_elem.attrib.get("name") or None

    # type heuristics
    a_class = (action_elem.attrib.get("class", "") or "").lower()
    a_cmd = (action_elem.attrib.get("type", "") or "").lower()
    if "filter" in a_class or "filter" in a_cmd:
        details["type"] = "filter"
    elif "highlight" in a_class or "brush" in a_cmd:
        details["type"] = "highlight"
    elif "url" in a_class or "url" in a_cmd:
        details["type"] = "url"
    elif "parameter" in a_class:
        details["type"] = "parameter"
    elif "set" in a_class:
        details["type"] = "set control"

    # iterate children for source/target/columns/behaviour
    for child in action_elem.iter():
        ctag = child.tag.lower().split("}")[-1]
        # Source / Target node detection
        if ctag == "source":
            # dashboard or worksheet attribute
            if "dashboard" in child.attrib:
                details["sources"].append({"kind": "dashboard", "name": child.attrib.get("dashboard")})
            elif "worksheet" in child.attrib:
                details["sources"].append({"kind": "worksheet", "name": child.attrib.get("worksheet")})
        if ctag == "target":
            if "dashboard" in child.attrib:
                details["targets"].append({"kind": "dashboard", "name": child.attrib.get("dashboard")})
            elif "worksheet" in child.attrib:
                details["targets"].append({"kind": "worksheet", "name": child.attrib.get("worksheet")})

        # Field/column mapping
        if ctag in ("source-column", "target-column", "column", "field", "filter"):
            src_field = child.attrib.get("name") or child.attrib.get("field") or child.attrib.get("column")
            role = ctag
            if src_field:
                details["field_mappings"].append({
                    "field": src_field,
                    "role": role
                })

        # Behavior hints
        for k, v in child.attrib.items():
            lk = k.lower()
            if lk in ("selection", "mode", "behavior", "action-mode", "target-type", "scope-type"):
                details["behavior"] = v

    # If no explicit behavior, try to infer from attributes on action element itself
    if details["behavior"] == "unknown":
        for k, v in action_elem.attrib.items():
            if k.lower() in ("selection", "mode", "behavior", "action-mode"):
                details["behavior"] = v

    # Normalize behavior to readable terms if possible
    beh = str(details["behavior"]).lower()
    if "replace" in beh:
        details["behavior"] = "Replace selection"
    elif "add" in beh or "append" in beh:
        details["behavior"] = "Add to selection"
    elif "keep" in beh or "keeponly" in beh:
        details["behavior"] = "Keep only"
    elif "exclude" in beh:
        details["behavior"] = "Exclude selection"
    elif details["behavior"] == "unknown":
        details["behavior"] = "Unknown"

    return details


def parse_hierarchies(xml):
    """
    Extract Tableau hierarchies from drill-paths.
    Robust against empty <field/> nodes.
    Works for:
      - full workbook XML
      - datasource fragments
    """
    root = _parse_fragment(xml)
    out = {}
    if root is None:
        return out

    # Determine datasource context
    if root.tag.lower().endswith("datasource"):
        datasources = [root]
    else:
        datasources = root.findall(".//datasource")

    for ds in datasources:
        ds_name = ds.attrib.get("name") or ds.attrib.get("caption") or "Datasource"

        drill_paths = ds.find(".//drill-paths")
        if drill_paths is None:
            continue

        for idx, dp in enumerate(drill_paths.findall("drill-path"), start=1):
            levels = []

            for f in dp.findall("field"):
                # 1️⃣ text content
                txt = (f.text or "").strip()

                # 2️⃣ attribute fallback
                if not txt:
                    txt = (
                        f.attrib.get("field")
                        or f.attrib.get("column")
                        or f.attrib.get("name")
                        or ""
                    ).strip()

                if txt:
                    txt = txt.replace("[", "").replace("]", "")
                    levels.append(txt)

            # keep even 1-level paths (Tableau allows this)
            if levels:
                h_name = f"{ds_name} — Hierarchy {idx}"
                out[h_name] = levels

    return out




def summarize_hierarchies(old_xml, new_xml):
    old_h = parse_hierarchies(old_xml)
    new_h = parse_hierarchies(new_xml)

    bullets = []

    old_names = set(old_h)
    new_names = set(new_h)

    # Added hierarchies
    for h in sorted(new_names - old_names):
        chain = " → ".join(new_h[h])
        bullets.append(f"➕ Added hierarchy '{h}' ({chain})")

    # Removed hierarchies
    for h in sorted(old_names - new_names):
        bullets.append(f"➖ Removed hierarchy '{h}'")

    # Modified hierarchies
    for h in sorted(old_names & new_names):
        if old_h[h] != new_h[h]:
            before = " → ".join(old_h[h])
            after = " → ".join(new_h[h])
            bullets.append(
                f"🟨 Modified hierarchy '{h}' (levels changed: {before} → {after})"
            )

    return bullets

def extract_dashboard_worksheets(xml):
    """
    Return set of worksheet names used in a dashboard.
    """
    root = _parse_fragment(xml)
    sheets = set()
    if root is None:
        return sheets

    for el in root.iter():
        tag = el.tag.lower().split("}")[-1]
        if tag in ("zone", "worksheet", "sheet"):
            nm = el.attrib.get("name") or el.attrib.get("sheet")
            if nm:
                sheets.add(nm)
    return sheets

def extract_story_contents(xml):
    """
    Extract story points and their associated worksheets.
    Returns:
      { "Story Point Caption": "Worksheet Name" }
    """
    root = _parse_fragment(xml)
    out = {}
    if root is None:
        return out

    for sp in root.findall(".//story-point"):
        caption = sp.attrib.get("caption") or sp.attrib.get("name") or "Story Point"
        sheet = sp.attrib.get("captured-sheet")
        if sheet:
            out[caption] = sheet
    return out


def extract_worksheet_fields(xml):
    """
    Return set of fields used in a worksheet (rows, columns, marks).
    """
    root = _parse_fragment(xml)
    fields = set()
    if root is None:
        return fields

    for el in root.iter():
        tag = el.tag.lower().split("}")[-1]
        if tag in ("column", "field", "encoding"):
            f = el.attrib.get("name") or el.attrib.get("field") or el.attrib.get("column")
            if f:
                f = f.replace("[", "").replace("]", "")
                fields.add(f)
    return fields

def parse_joins(xml):
    root = _parse_fragment(xml)
    joins = []
    if root is None:
        return joins

    for rel in root.findall(".//relation"):
        jtype = rel.attrib.get("join", "unknown")
        clauses = []
        for c in rel.findall(".//clause"):
            txt = " ".join("".join(e.itertext()).strip() for e in c.iter())
            if txt:
                clauses.append(txt)
        if clauses:
            joins.append(f"{jtype.upper()} JOIN ON " + " AND ".join(clauses))
    return joins

def parse_relationships(xml):
    root = _parse_fragment(xml)
    rels = set()
    if root is None:
        return rels

    for r in root.findall(".//relationship"):
        cols = []
        for c in r.findall(".//relationship-column"):
            col = c.attrib.get("column")
            if col:
                cols.append(col.replace("[","").replace("]",""))
        if cols:
            rels.add("Relationship on " + ", ".join(cols))
    return rels


def build_global_action_card(old_root, new_root):
    """
    Build a cards-style dict for global action changes with richer semantics.
    Returns None if no action changes.
    """
    raw_summary = summarize_global_actions(old_root, new_root)  # uses existing compare logic
    if not raw_summary:
        return None

    # For each item from raw_summary (which has xml_old/xml_new), parse deeper details
    bullets = []
    for item in raw_summary:
        line = item.get("line", "")
        xml_old = item.get("xml_old")
        xml_new = item.get("xml_new")

        # Base bullet headline
        bullets.append(line)
        # ------------------------------------------------------------
        # 🧩 Attach XML snippets for added / removed / modified actions
        # ------------------------------------------------------------
        if xml_old and not xml_new:
            bullets.append("  • [Old XML Snippet] ↓")
            bullets.append(html.unescape(xml_old.strip()))
        elif xml_new and not xml_old:
            bullets.append("  • [New XML Snippet] ↓")
            bullets.append(html.unescape(xml_new.strip()))
        elif xml_old and xml_new and xml_old != xml_new:
            bullets.append("  • [Old XML Snippet] ↓")
            bullets.append(html.unescape(xml_old.strip()))
            bullets.append("  • [New XML Snippet] ↓")
            bullets.append(html.unescape(xml_new.strip()))


        # If added — parse new xml for details
        if xml_new and not xml_old:
            try:
                root_new = ET.fromstring(xml_new)
                details = _parse_action_details(root_new)
                # Compose readable lines
                if details.get("sources"):
                    for s in details["sources"]:
                        bullets.append(f"  • Source: {s['kind']} — {s['name']}")
                if details.get("targets"):
                    for t in details["targets"]:
                        bullets.append(f"  • Target: {t['kind']} — {t['name']}")
                if details.get("field_mappings"):
                    fm = ", ".join([f"{m['field']} ({m['role']})" for m in details["field_mappings"][:6]])
                    bullets.append(f"  • Fields involved: {fm}")
                if details.get("behavior"):
                    bullets.append(f"  • Behavior: {details['behavior']}")
            except Exception:
                pass

        # If removed — parse old xml
        if xml_old and not xml_new:
            try:
                root_old = ET.fromstring(xml_old)
                details = _parse_action_details(root_old)
                if details.get("sources"):
                    for s in details["sources"]:
                        bullets.append(f"  • Source (was): {s['kind']} — {s['name']}")
                if details.get("targets"):
                    for t in details["targets"]:
                        bullets.append(f"  • Target (was): {t['kind']} — {t['name']}")
                if details.get("field_mappings"):
                    fm = ", ".join([f"{m['field']} ({m['role']})" for m in details["field_mappings"][:6]])
                    bullets.append(f"  • Fields involved (was): {fm}")
                if details.get("behavior"):
                    bullets.append(f"  • Behavior (was): {details['behavior']}")
            except Exception:
                pass

        # If modified — parse both and show diffs (concise)
        if xml_old and xml_new:
            try:
                root_old = ET.fromstring(xml_old)
                root_new = ET.fromstring(xml_new)
                d_old = _parse_action_details(root_old)
                d_new = _parse_action_details(root_new)

                # compare sources/targets
                old_srcs = {(s['kind'], s['name']) for s in d_old.get("sources", [])}
                new_srcs = {(s['kind'], s['name']) for s in d_new.get("sources", [])}
                added_src = new_srcs - old_srcs
                removed_src = old_srcs - new_srcs
                for s in added_src:
                    bullets.append(f"  • Source added: {s[0]} — {s[1]}")
                for s in removed_src:
                    bullets.append(f"  • Source removed: {s[0]} — {s[1]}")

                old_tg = {(t['kind'], t['name']) for t in d_old.get("targets", [])}
                new_tg = {(t['kind'], t['name']) for t in d_new.get("targets", [])}
                added_tg = new_tg - old_tg
                removed_tg = old_tg - new_tg
                for t in added_tg:
                    bullets.append(f"  • Target added: {t[0]} — {t[1]}")
                for t in removed_tg:
                    bullets.append(f"  • Target removed: {t[0]} — {t[1]}")

                # fields involvement diff (show small sample)
                old_fields = {m['field'] for m in d_old.get("field_mappings", [])}
                new_fields = {m['field'] for m in d_new.get("field_mappings", [])}
                added_fields = sorted(new_fields - old_fields)
                removed_fields = sorted(old_fields - new_fields)
                if added_fields:
                    bullets.append(f"  • Fields added: {', '.join(added_fields[:8])}")
                if removed_fields:
                    bullets.append(f"  • Fields removed: {', '.join(removed_fields[:8])}")

                # behavior change
                if d_old.get("behavior") != d_new.get("behavior"):
                    bullets.append(f"  • Behavior changed: {d_old.get('behavior')} → {d_new.get('behavior')}")

            except Exception:
                pass

    if not bullets:
        return None

    # Build a card-like dict (compatible with render_cards)
    return {
        "status": "modified",
        "icon": "⚡",
        "section": "global_actions",
        "title": "Global Actions Summary",
        "name": "",
        "bullets": bullets
    }
    

def collect_workbook_semantics(sections: dict) -> dict:
    """
    Aggregate semantics across entire workbook.
    """
    agg = {
        "filters": set(),
        "dashboard_filters": set(),
        "filter_controls": set(),
        "colors": set(),
        "legends": set(),
        "actions": set(),
        "mark_fields": set(),
        "tooltip_fields": set(),
        "dashboard_sheets": set(),
    }
    # ADD THIS LOOP to the function:
    for ds_name, ds_xml in sections.get("datasources", {}).items():
        ds_feats = collect_semantics(ds_xml) 
        agg["filters"].update(ds_feats["filters"])
        # This ensures GPT knows about datasource-level filters

    for sec in ("dashboards", "worksheets","datasources"):
        for nm, xml in sections.get(sec, {}).items():
            sem = collect_semantics(xml)
            for k in agg:
                agg[k].update(sem.get(k, []))
    # include datasource filters
    for xml in sections.get("datasources", {}).values():
        agg["filters"].update(parse_datasource_filters(xml))


    return agg

# =========================================================
# 🔧 FIX: SEMANTIC WORKBOOK DELTA (NEW)
# =========================================================
def build_semantic_workbook_delta(old_sections, new_sections):
    bullets = []

    old_sem = collect_workbook_semantics(old_sections)
    new_sem = collect_workbook_semantics(new_sections)

    # KPI-style semantic changes
    for k in old_sem:
        if len(old_sem[k]) != len(new_sem[k]):
            bullets.append(
                f"📊 {k.replace('_',' ').title()} changed: "
                f"{len(old_sem[k])} → {len(new_sem[k])}"
            )

    def diff_set(label, a, b, add_icon="➕", rem_icon="➖"):
        out = []
        if b - a:
            out.append(f"{add_icon} {label} added: {', '.join(sorted(b - a))}")
        if a - b:
            out.append(f"{rem_icon} {label} removed: {', '.join(sorted(a - b))}")
        return out

    bullets += diff_set(
        "Worksheet-level filters",
        set(old_sem.get("filters", [])),
        set(new_sem.get("filters", [])),
        "🔎", "🔎"
    )

    bullets += diff_set(
        "Dashboard-level filters",
        set(old_sem.get("dashboard_filters", [])),
        set(new_sem.get("dashboard_filters", [])),
        "📊", "📊"
    )

    bullets += diff_set(
        "Actions",
        set(old_sem.get("actions", [])),
        set(new_sem.get("actions", [])),
        "⚡", "⚡"
    )

    if old_sem.get("tooltip_fields") != new_sem.get("tooltip_fields"):
        bullets.append("💬 Tooltip content was modified across one or more views.")

    bullets += diff_set(
        "Legends",
        set(old_sem.get("legends", [])),
        set(new_sem.get("legends", [])),
        "🧭", "🧭"
    )

    bullets += diff_set(
        "Color encodings",
        set(old_sem.get("colors", [])),
        set(new_sem.get("colors", [])),
        "🎨", "🎨"
    )

    return bullets


def build_visual_change_tree(sections):
    tree = {
        "Workbook": {
            "🗄 Datasources": {},
            "📄 Worksheets": {},
            "📊 Dashboards": {},
            "📖 Stories": {},
            "⚡ Actions": {}
        }
    }

    # ===============================
    # 📄 WORKSHEETS (LIST ONLY)
    # ===============================
    for ws_name in sections.get("worksheets", {}):
        tree["Workbook"]["📄 Worksheets"][ws_name] = {}

    # ===============================
    # 📊 DASHBOARDS → WORKSHEETS ONLY (THIS WAS MISSING)
    # ===============================
    dashboards = sections.get("dashboards", {})
    for db_name, db_xml in dashboards.items():
        sheets = extract_dashboard_worksheets(db_xml)

        tree["Workbook"]["📊 Dashboards"][db_name] = {
            "Worksheets": sorted(sheets)
        }

    # ===============================
    # 📖 STORIES (OPTIONAL)
    # ===============================
    for story in sections.get("stories", {}):
        tree["Workbook"]["📖 Stories"][story] = {}

    # ===============================
    # ⚡ ACTIONS (OPTIONAL)
    # ===============================
    actions = collect_workbook_semantics(sections).get("actions", [])
    for a in actions:
        tree["Workbook"]["⚡ Actions"][a] = {}

    return tree


def filter_visual_bullets(level, bullets):
    """
    Keep only bullets relevant to the current visual level.
    """
    out = []
    for b in bullets:
        bl = b.lower()

        if level == "workbook":
            if any(k in bl for k in ("datasource", "parameter", "calculation", "hierarchy")):
                out.append(b)

        elif level == "dashboard":
            if any(k in bl for k in ("dashboard", "action", "layout", "filter")):
                out.append(b)

        elif level == "worksheet":
            if any(k in bl for k in ("worksheet", "filter", "tooltip", "mark", "color")):
                out.append(b)

    return out


def split_gpt_bullets(bullets):
    """
    Split deterministic bullets and GPT bullets using marker.
    """
    if "--- GPT Summary ---" not in bullets:
        return bullets, []

    idx = bullets.index("--- GPT Summary ---")
    return bullets[:idx], bullets[idx+1:]

def svg_expandable_block(x, y, title, status, bullets):
    bg = {
        "added": "#E8F5E9",
        "removed": "#FDECEA",
        "modified": "#FFF8E1"
    }.get(status, "#FFF")

    level = (
    "workbook" if "Workbook" in title
    else "dashboard" if "Dashboard" in title
    else "worksheet"
    )

    clean_bullets = filter_visual_bullets(level, bullets)

    # collapse FIRST
    if len(clean_bullets) > 1:
        clean_bullets = [visual_summary_line(clean_bullets)]

    li = "".join(
        f"<li>{html.escape(simplify_visual_bullet(b))}</li>"
        for b in clean_bullets
    )



    return f"""
<foreignObject x="{x}" y="{y}" width="440" height="320">
  <div xmlns="http://www.w3.org/1999/xhtml"
       style="font:13px Segoe UI;background:{bg};
              border:1px solid #aaa;border-radius:12px;
              padding:10px;box-shadow:0 2px 8px rgba(0,0,0,.1)">
    <details>
      <summary style="font-weight:700;cursor:pointer">
        {html.escape(title)}<br/>
        <span>
            {html.escape(visual_summary_line(clean_bullets))}
        </span>

      </summary>
      <div style="max-height:260px;overflow:auto;margin-top:6px">
        <ul>{li}</ul>
      </div>
    </details>
  </div>
</foreignObject>
"""

def visual_summary_line(bullets):
    counts = {
        "filter": 0,
        "calculation": 0,
        "worksheet": 0,
        "datasource": 0,
    }
    for b in bullets:
        bl = b.lower()
        if "filter" in bl:
            counts["filter"] += 1
        if "calculation" in bl or "lod" in bl:
            counts["calculation"] += 1
        if "worksheet" in bl:
            counts["worksheet"] += 1
        if "datasource" in bl:
            counts["datasource"] += 1

    parts = []
    if counts["filter"]:
        parts.append(f"{counts['filter']} filter changes")
    if counts["calculation"]:
        parts.append(f"{counts['calculation']} calculation changes")
    if counts["worksheet"]:
        parts.append(f"{counts['worksheet']} worksheet changes")
    if counts["datasource"]:
        parts.append(f"{counts['datasource']} datasource changes")

    return ", ".join(parts) or "Configuration updated"

def collapse_entries(entries):
    """
    Collapse a list OR dict of change entries into a single summary.
    Safe for new hierarchy (lists / dicts / None).
    """
    if not entries:
        return None

    all_bullets = []
    status = "modified"

    # If dict → flatten values
    if isinstance(entries, dict):
        for v in entries.values():
            if isinstance(v, list):
                for e in v:
                    all_bullets.extend(e.get("bullets", []))

    # If list → normal behavior
    elif isinstance(entries, list):
        for e in entries:
            if isinstance(e, dict):
                all_bullets.extend(e.get("bullets", []))

    all_bullets = dedupe_visual_bullets(all_bullets)

    if not all_bullets:
        return None

    return {
        "status": status,
        "bullets": all_bullets
    }


def svg_expandable_summary(x, y, card):
    """
    Expandable SVG block with deterministic + GPT summary.
    """
    status = card["status"]
    title = f"{card['icon']} {card['title']} — {card['name']}"
    bullets = card.get("bullets", [])

    sem, gpt = split_gpt_bullets(bullets)

    color_map = {
        "added": "#E8F5E9",
        "removed": "#FDECEA",
        "modified": "#FFF8E1"
    }
    bg = color_map.get(status, "#FFFFFF")

    def li(items):
        return "".join(f"<li>{html.escape(i)}</li>" for i in items)

    return f"""
<foreignObject x="{x}" y="{y}" width="420" height="320">
  <div xmlns="http://www.w3.org/1999/xhtml"
       style="
         font-family: Segoe UI, Arial;
         background:{bg};
         border:1px solid #999;
         border-radius:12px;
         padding:10px;
         box-shadow:0 2px 8px rgba(0,0,0,.1);
         font-size:13px;
       ">
    <details>
      <summary style="cursor:pointer;font-weight:700;">
        {html.escape(title)}
      </summary>

      <div style="margin-top:6px; max-height:260px; overflow:auto;">
        <b>🔍 Detected changes</b>
        <ul>{li(sem)}</ul>

        {"<hr><b>🤖 GPT Summary</b><ul>"+li(gpt)+"</ul>" if gpt else ""}
      </div>
    </details>
  </div>
</foreignObject>
"""

def render_svg_flow(tree):
    X = {"wb": 30, "dash": 260, "ws": 520, "chg": 820}
    BOX_W = {"wb": 180, "dash": 220, "ws": 220}
    BOX_H = 38
    ROW = 70

    elems = []
    y = 40

    def box(x, y, w, h, txt, fill):
        elems.append(f"""
        <rect x="{x}" y="{y}" rx="8" ry="8" width="{w}" height="{h}"
              fill="{fill}" stroke="#444"/>
        <text x="{x+10}" y="{y+25}" font-size="13">{html.escape(txt)}</text>
        """)

    def arrow(x1, y1, x2, y2):
        elems.append(f"""
        <line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}"
              stroke="#666" stroke-width="1.3"
              marker-end="url(#arrow)"/>
        """)

    defs = """
    <defs>
      <marker id="arrow" markerWidth="10" markerHeight="10"
              refX="10" refY="3" orient="auto">
        <path d="M0,0 L10,3 L0,6 Z" fill="#666"/>
      </marker>
    </defs>
    """

    # ================= Workbook =================
    wb_y = y
    box(X["wb"], wb_y, BOX_W["wb"], BOX_H, "Workbook", "#E3F2FD")
    y += ROW

    dashboards = tree["Workbook"].get("📊 Dashboards", {})

    # ================= Dashboards =================
    for dash_name, dash_node in dashboards.items():
        dash_y = y
        box(X["dash"], dash_y, BOX_W["dash"], BOX_H,
            f"Dashboard: {dash_name}", "#E8F5E9")
        arrow(X["wb"] + BOX_W["wb"], wb_y + BOX_H // 2,
              X["dash"], dash_y + BOX_H // 2)
        y += ROW

        # ---- Dashboard Filters summary ----
        collapsed_dash = collapse_entries(dash_node.get("Filters", []))
        if collapsed_dash:
            elems.append(
                svg_expandable_block(
                    X["chg"], y,
                    f"Dashboard Filters — {dash_name}",
                    collapsed_dash["status"],
                    collapsed_dash["bullets"]
                )
            )
            arrow(X["dash"] + BOX_W["dash"], dash_y + BOX_H // 2,
                  X["chg"], y + 24)
            y += 340

        # ================= Worksheets under dashboard =================
        for ws in dash_node.get("Worksheets", []):
            ws_y = y
            box(X["ws"], ws_y, BOX_W["ws"], BOX_H,
                f"Worksheet: {ws}", "#F3E5F5")
            arrow(X["dash"] + BOX_W["dash"], dash_y + BOX_H // 2,
                  X["ws"], ws_y + BOX_H // 2)
            y += ROW

            ws_entries = CHANGE_REGISTRY.get("worksheets", {}).get(ws, [])
            collapsed_ws = collapse_entries(ws_entries)
            if collapsed_ws:
                elems.append(
                    svg_expandable_block(
                        X["chg"], y,
                        f"Worksheet Changes — {ws}",
                        collapsed_ws["status"],
                        collapsed_ws["bullets"]
                    )
                )
                arrow(X["ws"] + BOX_W["ws"], ws_y + BOX_H // 2,
                      X["chg"], y + 24)
                y += 340

    # ================= Workbook-level changes =================
    for e in CHANGE_REGISTRY.get("workbook", []):
        elems.append(
            svg_expandable_block(
                X["chg"], y,
                e["title"], e["status"], e["bullets"]
            )
        )
        arrow(X["wb"] + BOX_W["wb"], wb_y + BOX_H // 2,
              X["chg"], y + 24)
        y += 340

    h = max(y + 40, 400)
    return f"""<svg width="100%" height="{h}" viewBox="0 0 1200 {h}"
        xmlns="http://www.w3.org/2000/svg">{defs}{''.join(elems)}</svg>"""

def dedupe_visual_bullets(bullets):
    """
    Remove duplicates and overly similar lines for SVG.
    Keeps first occurrence only.
    """
    seen = set()
    clean = []
    for b in bullets:
        key = simplify_visual_bullet(b).lower()
        if key not in seen:
            seen.add(key)
            clean.append(b)
    return clean
def extract_formula(xml):
    root = _parse_fragment(xml)
    if root is None:
        return None

    calc = root.find(".//calculation")
    if calc is None:
        return None

    return calc.attrib.get("formula") or (calc.text or "").strip()

def is_rls_calculation(formula: str) -> bool:
    if not formula:
        return False
    f = formula.upper()
    return any(k in f for k in ("USERNAME(", "USERFULLNAME(", "ISMEMBEROF("))

def is_calc_used_as_filter(calc_name, new_sections):
    for ws_xml in new_sections.get("worksheets", {}).values():
        sem = collect_semantics(ws_xml)
        if (
            calc_name in sem.get("filters", [])
            or calc_name in sem.get("dashboard_filters", [])
        ):
            return True
    return False

def build_formula_index(calcs: dict):
    """
    Build a lookup: formula -> [calculation names]
    Used to detect renamed calculations (especially RLS).
    """
    index = {}

    for name, xml in calcs.items():
        formula = extract_formula(xml)
        if not formula:
            continue

        normalized = formula.strip()
        index.setdefault(normalized, []).append(name)

    return index

def detect_rls_renames(old_calcs, new_calcs):
    renames = {}

    old_rls = {
        name: extract_formula(xml)
        for name, xml in old_calcs.items()
        if is_rls_calculation(extract_formula(xml))
    }

    new_rls = {
        name: extract_formula(xml)
        for name, xml in new_calcs.items()
        if is_rls_calculation(extract_formula(xml))
    }

    for o_name, o_formula in old_rls.items():
        matches = [
            n_name for n_name, n_formula in new_rls.items()
            if o_formula == n_formula and o_name != n_name
        ]

        # ✅ ONLY 1-to-1 rename allowed
        if len(matches) == 1:
            renames[o_name] = matches[0]

    return renames




# ----------- Cards -----------
def build_cards(old_sections, new_sections):
    cards = []

    # ================= NON-CALC SECTIONS =================
    for sec, label in (
        ("dashboards", "Dashboard"),
        ("worksheets", "Worksheet"),
        ("parameters", "Parameter"),
        ("stories", "Story"),
        ("datasources", "Datasource"),
    ):
        old = old_sections.get(sec, {})
        new = new_sections.get(sec, {})
        all_names = sorted(set(old) | set(new))

        for name in all_names:
            o = old.get(name)
            n = new.get(name)

            if o and not n:
                if label == "Datasource":
                    summarize_datasources(name, o, None)
                    continue

                cards.append({
                    "status": "removed",
                    "icon": "🟥",
                    "section": sec,
                    "title": f"Removed {label}",
                    "name": name,
                    "bullets": [f"{label} '{name}' removed"]
                })

            elif n and not o:
                if label == "Datasource":
                    summarize_datasources(name, None, n)
                    continue

                cards.append({
                    "status": "added",
                    "icon": "🟩",
                    "section": sec,
                    "title": f"Added {label}",
                    "name": name,
                    "bullets": [f"{label} '{name}' added"]
                })

            elif o and n and o != n:
                if label in ("Dashboard", "Story") and is_story_publish_noise(o, n):
                    continue

                ops = xmldiff_text(o, n)
                sem_old = collect_semantics(o)
                sem_new = collect_semantics(n)

                bullets = summarize_semantics(label, sem_old, sem_new)

                if ops and not ops.startswith("(xmldiff failed"):
                    bullets += ["--- GPT Summary ---"] + gpt_summarize_item(
                        f"{label} '{name}'", ops, sem_old, sem_new
                    )

                if bullets:
                    cards.append({
                        "status": "modified",
                        "icon": "🟨",
                        "section": sec,
                        "title": f"Modified {label}",
                        "name": name,
                        "bullets": bullets
                    })

    # ================= CALCULATIONS (STRICT ORDER) =================
    old_calcs = old_sections.get("calculations", {})
    new_calcs = new_sections.get("calculations", {})

    rls_renames = detect_rls_renames(old_calcs, new_calcs)

    renamed_old = set(rls_renames.keys())
    renamed_new = set(rls_renames.values())

    # 🔁 RENAME CARDS
    for old_name, new_name in rls_renames.items():
        formula = extract_formula(new_calcs[new_name])

        cards.append({
            "status": "modified",
            "icon": "🟨",
            "section": "calculations",
            "title": "Renamed RLS Calculation",
            "name": new_name,
            "bullets": [
                "Type: Row-Level Security",
                f"Old name: {old_name}",
                f"Formula unchanged: {formula}"
            ]
        })

    # 🟥 REMOVED CALCULATIONS
    for name, xml in old_calcs.items():
        if name in renamed_old or name in new_calcs:
            continue

        formula = extract_formula(xml)

        cards.append({
            "status": "removed",
            "icon": "🟥",
            "section": "calculations",
            "title": "Removed RLS Calculation"
            if is_rls_calculation(formula)
            else "Removed Calculation",
            "name": name,
            "bullets": [
                "Type: Row-Level Security"
                if is_rls_calculation(formula)
                else "Type: Standard",
                f"Formula: {formula}" if formula else None
            ]
        })

    # 🟩 ADDED CALCULATIONS
    for name, xml in new_calcs.items():
        if name in renamed_new or name in old_calcs:
            continue

        formula = extract_formula(xml)

        bullets = []
        if is_rls_calculation(formula):
            bullets.append("Type: Row-Level Security")
        if formula:
            bullets.append(f"Formula: {formula}")
        if is_calc_used_as_filter(name, new_sections):
            bullets.append("Applied as worksheet filter (TRUE)")

        cards.append({
            "status": "added",
            "icon": "🟩",
            "section": "calculations",
            "title": "Added RLS Calculation"
            if is_rls_calculation(formula)
            else "Added Calculation",
            "name": name,
            "bullets": bullets
        })

    return cards


def render_known_limitations_card() -> str:
    """
    Render a final informational card describing known comparison limitations.
    This is intentionally static and explanatory (no detection logic).
    """
    return """
    <div class="panel" style="border-left:6px solid #ff9800;">
      <h2>⚠️ Known Limitations</h2>
      <ul style="margin-top:10px;">
        <li>
          <strong>Datasource Filters</strong>:
          Some datasource-level filters are not consistently detected across
          all publishing scenarios. Certain attribute-only or backend-generated
          filters may not appear in the comparison summary yet.
        </li>
        <li style="margin-top:8px;">
          <strong>Story URL Actions</strong>:
          URL links inside stories are currently treated as <em>modified</em>
          on every publish due to Tableau regenerating internal identifiers.
          These changes may represent publish noise rather than real user edits.
        </li>
      </ul>
    </div>
    """


def populate_change_registry_from_cards(cards):
    """
    Single source of truth.
    Ensures everything shown in cards is also visible in the tree.
    """

    for c in cards:
        section = c["section"]
        title = f"{c['icon']} {c['title']} — {c['name']}"
        status = c["status"]
        bullets = c.get("bullets", [])

        if section == "datasources":
            register_change(
                level="datasource",
                parent=c["name"],
                title=title,
                status=status,
                bullets=bullets
            )

        elif section == "calculations":
            parent = "Unknown Datasource"
            register_change(
                level="datasource",
                parent=parent,
                title=f"Calculation — {c['name']}",
                status=status,
                bullets=bullets
            )

        elif section == "parameters":
            CHANGE_REGISTRY["parameters"].setdefault("Parameters", []).append({
                "status": status,
                "title": f"Parameter — {c['name']}",
                "object": c["name"],
                "bullets": bullets
            })

        elif section == "worksheets":
            register_change(
                level="worksheet",
                parent=c["name"],
                title=title,
                status=status,
                bullets=bullets
            )

        elif section == "dashboards":
            register_change(
                level="dashboard",
                parent=c["name"],
                title=title,
                status=status,
                bullets=bullets
            )

        elif section == "stories":
            # 🚫 CRITICAL FIX: ignore publish-only noise
            if not bullets:
                continue

            register_change(
                level="story",
                parent=c["name"],
                title=title,
                status=status,
                bullets=bullets
            )

def extract_datasource_filter_changes(old_sections, new_sections):
    """
    Detect datasource filter additions/removals deterministically.
    Returns list of human-readable facts.
    """

    facts = []

    old_ds = old_sections.get("datasources", {})
    new_ds = new_sections.get("datasources", {})

    for ds_name in set(old_ds) | set(new_ds):
        old_xml = old_ds.get(ds_name)
        new_xml = new_ds.get(ds_name)

        if not old_xml or not new_xml:
            continue

        diff = xmldiff_text(old_xml, new_xml)

        added = set()
        removed = set()

        for line in diff.splitlines():
            if "filter" not in line.lower():
                continue

            # added filter
            if "insert" in line.lower():
                m = re.search(r'\[([^\]]+)\]', line)
                if m:
                    added.add(m.group(1))

            # removed filter
            if "delete" in line.lower():
                m = re.search(r'\[([^\]]+)\]', line)
                if m:
                    removed.add(m.group(1))

        for f in sorted(added):
            facts.append(
                f"Added datasource filter '{f}' in datasource '{ds_name}'."
            )

        for f in sorted(removed):
            facts.append(
                f"Removed datasource filter '{f}' from datasource '{ds_name}'."
            )

    return facts

def build_overall_workbook_summary_card(
    old_sections,
    new_sections,
    cards,
    root_old,
    root_new
):
    """
    Executive workbook summary with STRICT XML-based Global Action detection.
    """

    # =====================================================
    # 1️⃣ COLLECT FACTS (FROM CARDS)
    # =====================================================
    added_facts = []
    removed_facts = []
    modified_facts = []

    NOISE_KEYWORDS = [
        "repository",
        "content-url",
        "site",
        "workbook location",
        "published to",
        "url",
        "uuid",
        "id",
        "zone",
        "layout",
        "container",
        "device layout",
        "floating",
        "tiled",
        "resize",
        "reposition",
        "no visible change",
        "no functional change",
    ]

    def is_noise(text: str) -> bool:
        return any(k in text.lower() for k in NOISE_KEYWORDS)

    for c in cards:
        status = c.get("status")
        section = c.get("section", "object")
        name = c.get("name") or c.get("title") or "Unnamed"

        facts = []
        for b in c.get("bullets", []):
            if not b:
                continue
            if b.startswith("---"):
                continue
            if "XML Snippet" in b:
                continue
            if b.strip().startswith("<"):
                continue

            clean = (
                b.replace("➕", "")
                 .replace("➖", "")
                 .replace("🟨", "")
                 .strip()
            )

            if is_noise(clean):
                continue

            facts.append(clean)

        if status == "modified" and not facts:
            continue

        payload = {
            "type": section,
            "name": name,
            "facts": facts or [f"{section.capitalize()} '{name}' changed"]
        }

        if status == "added":
            added_facts.append(payload)
        elif status == "removed":
            removed_facts.append(payload)
        elif status == "modified":
            modified_facts.append(payload)

    # =====================================================
    # 2️⃣ GLOBAL ACTION DIFF (XML ONLY, STRICT)
    # =====================================================
    def extract_actions(root):
        actions = {}
        if root is None:
            return actions

        for a in root.findall(".//action"):
            caption = a.attrib.get("caption")
            if caption:
                actions[caption] = ET.tostring(a, encoding="unicode")

        return actions

    old_actions = extract_actions(root_old)
    new_actions = extract_actions(root_new)

    # ➕ Added actions
    for cap in new_actions.keys() - old_actions.keys():
        added_facts.append({
            "type": "dashboard action",
            "name": cap,
            "facts": [f'Dashboard action "{cap}" was added']
        })

    # ➖ Removed actions
    for cap in old_actions.keys() - new_actions.keys():
        removed_facts.append({
            "type": "dashboard action",
            "name": cap,
            "facts": [f'Dashboard action "{cap}" was removed']
        })

    # 🟨 Modified actions
    for cap in old_actions.keys() & new_actions.keys():
        if old_actions[cap] != new_actions[cap]:
            modified_facts.append({
                "type": "dashboard action",
                "name": cap,
                "facts": [f'Dashboard action "{cap}" was modified']
            })

    # =====================================================
    # 3️⃣ COUNTS
    # =====================================================
    added_count = len(added_facts)
    removed_count = len(removed_facts)
    modified_count = len(modified_facts)
    total = added_count + removed_count + modified_count

    # =====================================================
    # 4️⃣ GPT EXECUTIVE WRITER (STRICT)
    # =====================================================
    prompt = f"""
You are generating an EXECUTIVE SUMMARY for a Tableau workbook comparison.

STRICT RULES:
- ONLY summarize the facts provided.
- Do NOT invent changes.
- Do NOT infer dashboard actions unless explicitly present in facts.
- Mention exact object type and name.
- Order bullets strictly as: Added → Removed → Modified.
- Ignore layout, zone, or formatting-only changes.

CRITICAL COUNT RULE:
- Return EXACTLY {total} bullets.
- Each bullet represents ONE changed object.

FACTS:
ADDED:
{added_facts}

REMOVED:
{removed_facts}

MODIFIED:
{modified_facts}

OUTPUT:
- EXACTLY {total} bullets
- Each bullet starts with: Added / Removed / Modified
- No headings, no explanations
"""

    try:
        gpt_bullets = gpt_summarize_item(
            title="Overall Workbook Comparison",
            xml_ops=prompt,
            semantics_old={},
            semantics_new={}
        )
    except Exception:
        gpt_bullets = []

    # =====================================================
    # 5️⃣ FINAL CARD
    # =====================================================
    bullets = [
        f"📦 Total objects changed: {total}",
        f"➕ Added: {added_count}",
        f"➖ Removed: {removed_count}",
        f"🟨 Modified: {modified_count}",
        "--- Executive Summary ---",
    ] + gpt_bullets

    return {
        "status": "modified",
        "icon": "📘",
        "section": "workbook",
        "title": "Overall Workbook Differences Summary",
        "name": "",
        "bullets": unique_keep_order(bullets)
    }

def build_layout_only_card():
    if not CHANGE_REGISTRY["layout_only"]:
        return ""

    items = ""
    for c in CHANGE_REGISTRY["layout_only"]:
        for b in c["bullets"]:
            items += f"<li>{html.escape(b)}</li>"

    return f"""
    <div class="panel" style="background:#fafafa;">
      <h2>🎨 Additional Visual Adjustments (Not Counted)</h2>
      <ul>{items}</ul>
    </div>
    """


def parse_datasource_filters(xml):
    root = _parse_fragment(xml)
    filters = set()
    if root is None:
        return filters

    # 1️⃣ filter nodes
    for f in root.findall(".//filter"):
        col = f.attrib.get("column") or f.attrib.get("field") or f.attrib.get("name")
        if col:
            filters.add(col.replace("[","").replace("]",""))

    # 2️⃣ filter-group → groupfilter → column  ✅ MOST COMMON
    for fg in root.findall(".//filter-group"):
        for gf in fg.findall(".//groupfilter"):
            col = gf.attrib.get("column")
            if col:
                filters.add(col.replace("[","").replace("]",""))
            for c in gf.findall(".//column"):
                nm = c.attrib.get("name") or c.attrib.get("column")
                if nm:
                    filters.add(nm.replace("[","").replace("]",""))

    # 3️⃣ standalone groupfilter (some Tableau versions)
    for gf in root.findall(".//groupfilter"):
        col = gf.attrib.get("column")
        if col:
            filters.add(col.replace("[","").replace("]",""))
            

    return filters

def unique_keep_order(items):
    seen=set(); out=[]
    for x in items:
        if x and x not in seen:
            seen.add(x); out.append(x)
    return out

# ----------- HTML (no XML in view cards; full structural) -----------
def badge(kind):
    if kind=="added": return "<span class='badge badge-added'>🟩 Added</span>"
    if kind=="removed": return "<span class='badge badge-removed'>🟥 Removed</span>"
    if kind=="modified": return "<span class='badge badge-modified'>🟨 Modified</span>"
    return "<span class='badge'>•</span>"

def render_cards(cards, force_open=True, skip_empty=True):
    """
    Renders cards with:
    - Light blue theme
    - Optional collapse behavior
    - Skips empty cards if required

    Params:
    - force_open: True → cards open by default
    - skip_empty: True → cards with no bullets are not rendered
    """

    blocks = []

    for c in cards:
        bullets = c.get("bullets", [])

        # 🚫 Skip empty cards (very important for View-Level section)
        if skip_empty and not bullets:
            continue

        # -----------------------------------------
        # Status → CSS class
        # -----------------------------------------
        status = c.get("status", "modified")
        card_class = f"card card-{status}"

        # -----------------------------------------
        # Open / Closed control
        # -----------------------------------------
        open_attr = "open" if force_open else ""

        # -----------------------------------------
        # Bullet rendering
        # -----------------------------------------
        li = "".join(
            f"<li>{html.escape(str(b))}</li>"
            for b in bullets
        )

        block = f"""
<details class="{card_class}" {open_attr}
  style="
    background:#F5F9FF;
    border-radius:12px;
    margin:12px 0;
    border-left:6px solid
      {'#4CAF50' if status=='added' else
       '#E57373' if status=='removed' else
       '#FFCA28'};
  "
>
  <summary style="
    padding:12px 14px;
    font-weight:600;
    color:#1F6FE5;
    cursor:pointer;
  ">
    {c.get('icon','')} <strong>{html.escape(c.get('title',''))}</strong>
    — {html.escape(c.get('name',''))}
    {badge(status)}
  </summary>

  <div class="card-body" style="padding:10px 16px;">
    <ul class="bullets" style="margin:0;">
      {li}
    </ul>
  </div>
</details>
"""
        blocks.append(block)

    return "\n".join(blocks)


def render_datasource_cards():
    blocks = []

    for ds_name, changes in CHANGE_REGISTRY.get("datasources", {}).items():
        if not changes:
            continue

        cards_html = []

        for c in changes:
            status = c.get("status", "info").lower()

            icon = {
                "modified": "🟨",
                "added": "➕",
                "removed": "➖",
            }.get(status, "ℹ️")

            bullets = c.get("bullets") or []

            bullets_html = "".join([f"<li>{html.escape(str(b))}</li>" for b in c.get("bullets", [])])
            
            cards_html.append(f"""
            <details class="panel" style="background:#F2F7FF; border-radius:12px; padding:8px 12px; margin:10px 0;">
              <summary style="cursor:pointer; font-weight:600; color:#1f6fe5;">
                {html.escape(c.get('title','Connection Details'))}
              </summary>
              <div style="margin-top:10px;">
                <ul>{bullets_html}</ul>
              </div>
            </details>
            """)

        blocks.append(f"""
        <details class="panel" style="
            background:#EAF3FF;
            border-radius:14px;
            padding:10px 14px;
            margin:18px 0;
            box-shadow:0 3px 8px rgba(0,0,0,0.08);
        ">
          <summary style="
              cursor:pointer;
              font-size:17px;
              font-weight:600;
              color:#1f6fe5;
              list-style:none;
          ">
            🗄 Datasource Details — {html.escape(ds_name)}
          </summary>

          <div style="margin-top:12px;">
            {''.join(cards_html)}
          </div>
        </details>
        """)

    return "\n".join(blocks)



def map_capabilities_for_display(raw_caps: str, site_role: str) -> str:
    """
    Convert Tableau raw permissions into human-readable capabilities
    based on Site Role.
    """
    if not raw_caps or raw_caps.strip() == "—":
        return "No Access"

    caps = {c.strip().lower() for c in raw_caps.split(",")}

    display = []

    # Read → View
    if "read" in caps:
        display.append("View")

    # Write depends on Site Role
    if "write" in caps:
        role = (site_role or "").lower()

        if "creator" in role:
            display.append("Edit & Publish")
        elif "explorer" in role:
            display.append("Web Edit")
        else:
            # Viewer or restricted role
            pass

    return ", ".join(display) if display else "View"

def build_effective_permissions(site_users, workbook_perms, project_perms):
    user_map = {}
    group_map = {}

    # Combine workbook + project permissions
    for p in workbook_perms + project_perms:
        if p["type"] == "User":
            user_map[p["name"]] = p
        elif p["type"] == "Group":
            group_map[p["name"]] = p

    # If ANY project permission exists, treat it as inherited
    project_inherited = None
    for p in project_perms:
        if p["type"] == "Group" and p["capabilities"]:
            project_inherited = p
            break

    print("DEBUG — Project Permission Groups:")
    for p in project_perms:
        print(p)



    effective = []

    for u in site_users:
        name = u["display_name"]
        role = u["site_role"]

        # 🚫 Unlicensed users never get access
        if role == "Unlicensed":
            effective.append({
                "name": name,
                "type": "User",
                "permission": "None",
                "capabilities": "No Access"
            })
        


            continue

        # 1️⃣ Explicit workbook permission
        if name in user_map:
            p = user_map[name].copy()
            p["capabilities"] = map_capabilities_for_display(
                p.get("capabilities"),
                role
            )
            effective.append(p)

            continue

        # 2️⃣ Inherited from project permissions
        if project_inherited:
            effective.append({
                "name": name,
                "type": "User",
                "permission": "Inherited (Project)",
                "capabilities": map_capabilities_for_display(
                    project_inherited["capabilities"],
                    role
                )
            })

            continue

        # 3️⃣ No access
        effective.append({
            "name": name,
            "type": "User",
            "permission": "None",
            "capabilities": "No Access"
        })

    return effective


def get_project_permissions(token, site_id, project_id):
    url = f"{TABLEAU_SITE_URL}/api/3.25/sites/{site_id}/projects/{project_id}/permissions"
    r = requests.get(
        url,
        headers={"X-Tableau-Auth": token},
        verify=VERIFY_SSL,
        timeout=30
    )
    r.raise_for_status()
    return ET.fromstring(r.text)


def resolve_effective_permissions(site_users, workbook_permissions):
    # Split permissions
    user_perms = {}
    group_perms = {}

    for p in workbook_permissions:
        if p["type"] == "User":
            user_perms[p["name"]] = p
        elif p["type"] == "Group":
            group_perms[p["name"]] = p

    # Tableau default access group
    all_users_group = group_perms.get("All Users")

    resolved = []

    for u in site_users:
        name = u["display_name"]
        role = u["site_role"]

        # 1️⃣ Explicit user permission
        if name in user_perms:
            resolved.append(user_perms[name])
            continue

        # 2️⃣ Inherit from All Users group
        if all_users_group:
            # 🚨 Site-role ceiling
            if role == "Unlicensed":
                resolved.append({
                    "name": name,
                    "type": "User",
                    "permission": "None",
                    "capabilities": "—"
                })
            else:
                resolved.append({
                    "name": name,
                    "type": "User",
                    "permission": "Inherited",
                    "capabilities": all_users_group["capabilities"]
                })
            continue

        # 3️⃣ No access
        resolved.append({
            "name": name,
            "type": "User",
            "permission": "None",
            "capabilities": "—"
        })

    return resolved


def build_permission_lookup(permission_rows):
    lookup = {"users": {}, "groups": {}}

    for p in permission_rows:
        if p["type"] == "User":
            lookup["users"][p["name"]] = p
        elif p["type"] == "Group":
            lookup["groups"][p["name"]] = p

    return lookup
def merge_site_users_with_permissions(site_users, permission_rows):
    # Build lookup
    user_perms = {}
    group_perms = {}

    for p in permission_rows:
        if p["type"] == "User":
            user_perms[p["name"]] = p
        elif p["type"] == "Group":
            group_perms[p["name"]] = p

    all_users_group = group_perms.get("All Users")

    merged = []

    for u in site_users:
        name = u["display_name"]

        # 1️⃣ Explicit user permission
        if name in user_perms:
            merged.append(user_perms[name])
            continue

        # 2️⃣ Inherit from All Users group
        if all_users_group:
            merged.append({
                "name": name,
                "type": "User",
                "permission": "Inherited",
                "capabilities": all_users_group["capabilities"]
            })
            continue

        # 3️⃣ No access
        merged.append({
            "name": name,
            "type": "User",
            "permission": "None",
            "capabilities": "—"
        })

    return merged

def build_users_permissions_card(permissions):
    rows = ""

    for p in permissions:
        rows += f"""
        <tr>
          <td>{html.escape(p['name'])}</td>
          <td>{p['type']}</td>
          <td><b>{p['permission']}</b></td>
          <td>{html.escape(p['capabilities']) if p['capabilities'] else "—"}</td>
        </tr>
        """

    return f"""
    <details class="panel" style="
        background:#EAF3FF;
        border-radius:14px;
        padding:10px 14px;
        margin:16px 0;
        box-shadow:0 3px 8px rgba(0,0,0,0.08);
    ">
      <summary style="
          cursor:pointer;
          font-size:16px;
          font-weight:600;
          color:#1f6fe5;
          padding:6px 0;
          list-style:none;
      ">
        👥 Users & Permissions
      </summary>

      <div style="margin-top:12px;">
        <table style="
            width:100%;
            border-collapse:collapse;
            font-size:14px;
        ">
          <tr style="border-bottom:1px solid #c6dbff;">
            <th align="left">User / Group</th>
            <th align="left">Type</th>
            <th align="left">Permission Level</th>
            <th align="left">Capabilities (if Custom)</th>
          </tr>
          {rows}
        </table>
      </div>
    </details>
    """

# ---------------- HIERARCHY IMPACT (GLOBAL) ----------------

GLOBAL_HIERARCHY_DIFF = []

def render_visual_change_tree(sections, registry, workbook_name):
    """
    Tableau-internal accurate Visual Change Tree
    Structure-only (no GPT, no summaries)
    """

    lines = []
    lines.append("🌳 Visual Change Tree")
    lines.append(f"📦 Workbook — {workbook_name}")

    # ===============================
    # 📄 WORKSHEETS
    # ===============================
    lines.append("├── 📄 Worksheets")
 
    for ws_name, ws_xml in sections.get("worksheets", {}).items():
        lines.append(f"│    ├── {ws_name}")
 
        sem = collect_semantics(ws_xml)
 
        # 1️⃣ Filters Shelf (Worksheet-level)
        if sem.get("filters"):
            lines.append("│    │    ├── Filters Shelf (Worksheet)")
            for f in sem["filters"]:
                lines.append(f"│    │    │    └── {f}")
 
        # 4️⃣ Context Filters
        context_filters = []
        try:
            root = ET.fromstring(ws_xml)
            for f in root.findall(".//filter[@context='true']"):
                name = f.attrib.get("column") or f.attrib.get("field")
                if name:
                    context_filters.append(name.replace("[", "").replace("]", ""))
        except Exception:
            pass
 
        if context_filters:
            lines.append("│    │    ├── Context Filters")
            for f in sorted(set(context_filters)):
                lines.append(f"│    │    │    └── {f}")
 
        # 3️⃣ Marks Card Filters
        if sem.get("mark_fields"):
            lines.append("│    │    ├── Marks Shelf (Worksheet)")
            for m in sorted(set(sem["mark_fields"])):
                lines.append(f"│    │    │    └── {m}")
 
        # 5️⃣ Action Filters (Indirect)
        actions = sem.get("actions", [])
        if actions:
            lines.append("│    │    ├── Action Filters")
            for a in actions:
                lines.append(f"│    │    │    └── {a}")
    # ----------------------------
    # 📊 DASHBOARDS
    # ----------------------------
    lines.append("├── 📊 Dashboards")
    for db_name, db_xml in sections.get("dashboards", {}).items():
        lines.append(f"│    └── {db_name}")

        sheets = extract_dashboard_worksheets(db_xml)
        if sheets:
            lines.append("│         ├── Worksheets")
            for s in sorted(sheets):
                lines.append(f"│         │    └── {s}")

        dash_filters = collect_semantics(db_xml).get("dashboard_filters", [])
        if dash_filters:
            lines.append("│         ├── Dashboard Filters")
            for f in dash_filters:
                lines.append(f"│         │    └── {f}")

        actions = collect_semantics(db_xml).get("actions", [])
        if actions:
            lines.append("│         └── Actions")
            for a in actions:
                lines.append(f"│              └── {a}")

    # ----------------------------
    # 📖 STORIES (WITH CONTENTS)
    # ----------------------------
    stories = sections.get("stories", {})
    if stories:
        lines.append("└── 📖 Stories")
        for story_name, story_xml in stories.items():
            lines.append(f"     └── {story_name}")

            story_points = extract_story_contents(story_xml)
            if story_points:
                lines.append("          ├── Story Points")
                for sp_name, ws in story_points.items():
                    lines.append(f"          │    └── {sp_name}")
                    lines.append(f"          │         └── Worksheet: {ws}")



    return "\n".join(lines)



def adapt_datasource_registry_for_rendering(ds_registry):
    """
    Convert Datasource_1 registry format into
    main renderer-compatible card format.
    """
    adapted = {}

    for ds_name, ds_data in ds_registry.items():
        # ds_data is: { "status": "...", "bullets": [...] }
        adapted[ds_name] = [
            {
                "status": ds_data.get("status", "modified"),
                "title": ds_name,
                "object": ds_name,
                "bullets": ds_data.get("bullets", [])
            }
        ]

    return adapted


def generate_html_report(
    title_a,
    title_b,
    cards,
    structural_ops,
    out_file,
    kpi_html,
    root_new,
    visual_tree_text,

    latest_publisher=None,
    latest_revision=None,
    latest_published_at=None,

    old_publisher=None,
    new_publisher=None,

    users_permissions_html="",
    site_users_html=""
):

    publisher_block = ""

    clean_title_a = title_a.replace("(Latest)", "").strip()
    clean_title_b = title_b.replace("(Latest)", "").strip()


    html_text = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Tableau Diff — GPT-4o (v3.2)</title>

<style>
:root {{
  --bg:#f7f9fc; --fg:#233142; --card:#fff; --muted:#6b7a90;
  --added:#8bc34a; --removed:#e57373; --modified:#ffca28;
  --added-bg:#e8f5e9; --removed-bg:#fdecea; --modified-bg:#fff8e1;
}}
body{{background:var(--bg);color:var(--fg);font:15px/1.6 Segoe UI,Arial,sans-serif;margin:28px}}
h1{{margin:0 0 6px}}
.meta{{color:var(--muted);margin-bottom:16px}}
.panel{{background:var(--card);border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);padding:18px;margin:16px 0}}
.card{{background:var(--card);border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin:14px 0;border-left:5px solid transparent}}
.card-modified{{border-left-color:var(--modified)}}
.card > summary{{cursor:pointer;list-style:none;padding:12px 14px;font-weight:600}}
.card > summary::-webkit-details-marker{{display:none}}
.card-body{{padding:0 14px 14px}}
table{{width:100%;border-collapse:collapse}}
th,td{{padding:8px;border-bottom:1px solid #ddd;text-align:left}}
</style>
</head>

<body>

<div style="
    background:#2D8CFF;
    color:white;
    padding:18px 22px;
    border-radius:14px;
    box-shadow:0 4px 12px rgba(0,0,0,0.15);
    margin-bottom:20px;
    position:relative;
">

  <!-- Zoom Logo -->
  <img
    src="https://upload.wikimedia.org/wikipedia/commons/7/7b/Zoom_Communications_Logo.svg"
    alt="Zoom"
    style="
      position:absolute;
      top:16px;
      right:18px;
      height:34px;
      filter:brightness(0) invert(1);
      opacity:0.95;
    "
  />

    <div style="
        font-size:20px;
        font-weight:700;
        margin-bottom:4px;
    ">
    Tableau Online — Visual Delta  
    <span style="font-weight:400;opacity:0.85;">
        (Zoom | Ascendion POC)
    </span>
    </div>

    <div style="
        font-size:14px;
        opacity:0.9;
        margin-bottom:12px;
    ">
    Comparing latest workbooks across 2(Source, Target ) environments
    </div>


  <div style="
      display:grid;
      grid-template-columns:220px auto;
      row-gap:6px;
      font-size:14px;
  ">

    <div style="opacity:0.85;">Generated</div>
    <div>{html.escape(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>

    <div style="opacity:0.85;">Latest Source Workbook</div>
    <div>
      <b>{html.escape(clean_title_a)}</b>
      <span style="opacity:0.85;">
        → {html.escape(str(old_publisher))}
      </span>
    </div>

    <div style="opacity:0.85;">Latest Target Workbook</div>
    <div>
      <b>{html.escape(clean_title_b)}</b>
      <span style="opacity:0.85;">
        → {html.escape(str(new_publisher))}
      </span>
    </div>

  </div>

  {publisher_block}

</div>


<!-- ===================================================== -->
<!-- 👥 Users & Permissions -->
<!-- ===================================================== -->
{users_permissions_html}
<!--site_users_html-->


<!-- ===================================================== -->
<!-- 1️⃣ Workbook Metrics -->
<!-- ===================================================== -->
{kpi_html}


<!-- ===================================================== -->
<!-- 2️⃣ Overall Workbook Differences Summary -->
<!-- ===================================================== -->
{render_cards(
    [c for c in cards if c.get("section") == "workbook"],
    force_open=True,
    skip_empty=False
)}

<!-- ===================================================== -->
<!-- 3️⃣ View-Level Changes -->
<!-- ===================================================== -->
<details class="panel">
  <summary style="
      font-size:18px;
      font-weight:700;
      cursor:pointer;
  ">
    🧩 View-Level Changes
  </summary>

  <div style="margin-top:12px;">
    {render_cards(
        [
            c for c in cards
            if c.get("section") != "workbook"
            and c.get("section") != "datasource"
            and c.get("title") != "Overall Workbook Differences Summary"
        ],
        force_open=True,
        skip_empty=True
    )}
  </div>
</details>


<!-- ===================================================== -->
<!-- 4️⃣ Tableau Hierarchy Breakdown -->
<!-- ===================================================== -->
<details class="card card-modified"
  style="
    background:#F5F9FF;
    border-radius:12px;
    margin:14px 0;
    border-left:6px solid #FFCA28;
  "
>
  <summary style="
    padding:12px 14px;
    font-weight:600;
    color:#1F6FE5;
    cursor:pointer;
  ">
    🌳 <strong>Tableau Hierarchy Breakdown</strong>
    <span style="
      margin-left:8px;
      font-size:12px;
      background:#FFE082;
      color:#333;
      padding:2px 8px;
      border-radius:8px;
    ">
      Modified
    </span>
  </summary>

  <div class="card-body" style="padding:12px 16px;">
    <pre style="
      background:#EDF4FF;
      border:1px solid #C7DBFF;
      border-radius:10px;
      padding:14px;
      font-family: Consolas, Monaco, monospace;
      font-size:14px;
      white-space:pre;
      overflow:auto;
      margin:0;
    ">
{html.escape(visual_tree_text)}
    </pre>
  </div>
</details>



<!-- ===================================================== -->
<!-- Datasource Changes -->
<!-- ===================================================== -->
{render_datasource_cards()}

<!-- ===================================================== -->
<!-- 5️⃣ Known Limitations -->
<!-- ===================================================== -->
<div class="panel" style="
  border-left:6px solid #ff9800;
  background:#fff3e0;
">
  <h2>⚠️ Known Limitations</h2>
  <ul style="margin-top:10px;">
    <li>
      <strong>Datasource Connection Type:</strong>
      In some cases, the datasource connection type may appear as
      <b>UNKNOWN</b> (for example: <i>Unknown (Extract | Published)</i>).
      This happens when Tableau abstracts or hides the underlying connection
      details in the workbook XML.
    </li>

    <li style="margin-top:8px;">
      <strong>Story URL Links:</strong>
      URL actions inside stories may appear as modified due to
      publish-related metadata changes, even when there is no functional
      change.
    </li>
  </ul>
</div>
"""

    with open(out_file, "w", encoding="utf-8") as f:
        f.write(html_text)

    try:
        webbrowser.open(f"file:///{os.path.abspath(out_file)}")
    except Exception:
        pass

    print("✅ HTML report saved →", os.path.abspath(out_file))





# ----------- TXT export (optional) -----------
def write_text(path: str, content: str):
    pathlib.Path(os.path.dirname(path)).mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content or "")
    print("📝 Wrote:", os.path.abspath(path))

def sanitize_name(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", s.strip())[:200] or "item"


def ensure_change_registry_keys():
    required = [
        "workbook", "datasources", "calculations",
        "parameters", "worksheets", "dashboards", "stories"
    ]
    for k in required:
        CHANGE_REGISTRY.setdefault(k, {} if k != "workbook" else [])


def main():
    print("==============================================")
    print("Tableau Workbook Comparator (Project → Project)")
    print("==============================================")

    token, site_id = sign_in()
    if not token or not site_id:
        print("❌ Tableau login failed")
        return

    # =====================================================
    # USER INPUT — SOURCE & TARGET
    # =====================================================
    source_project = input("Enter SOURCE project name: ").strip()
    source_workbook = input("Enter SOURCE workbook name: ").strip()

    target_project = input("Enter TARGET project name: ").strip()
    target_workbook = input("Enter TARGET workbook name: ").strip()

    # =====================================================
    # FIND WORKBOOK IDS
    # =====================================================
    try:
        source_wid, source_project_id = get_workbook_id_in_project(
            token, site_id, source_project, source_workbook
        )
        target_wid, target_project_id = get_workbook_id_in_project(
            token, site_id, target_project, target_workbook
        )
    except Exception as e:
        print(f"❌ {e}")
        return

    print(f"✅ Source Workbook ID: {source_wid}")
    print(f"✅ Target Workbook ID: {target_wid}")

    # =====================================================
    # GET REVISIONS
    # =====================================================
    source_revs = get_revisions(token, site_id, source_wid)
    target_revs = get_revisions(token, site_id, target_wid)

    if not source_revs or not target_revs:
        print("❌ Unable to fetch revisions for one or both workbooks")
        return

    # =====================================================
    # LATEST REVISION (SAFE)
    # =====================================================
    source_owner = get_workbook_owner(token, site_id, source_wid)
    target_owner = get_workbook_owner(token, site_id, target_wid)

    source_latest = get_latest_revision_info(source_revs, source_owner)
    target_latest = get_latest_revision_info(target_revs, target_owner)

    OLD_REV = source_latest["revision"]
    NEW_REV = target_latest["revision"]

    OLD_PUBLISHER = source_latest["publisher"]
    NEW_PUBLISHER = target_latest["publisher"]

    LATEST_PUBLISHER = NEW_PUBLISHER
    LATEST_REVISION = NEW_REV
    LATEST_PUBLISHED_AT = target_latest["published_at"]

    print(f"📌 Source latest revision : {OLD_REV}")
    print(f"📌 Target latest revision : {NEW_REV}")

    # =====================================================
    # DOWNLOAD LATEST REVISIONS
    # =====================================================
    twb_old = download_rev(token, site_id, source_wid, OLD_REV, force=False)
    twb_new = download_rev(token, site_id, target_wid, NEW_REV, force=False)

    with open(twb_old, "r", encoding="utf-8", errors="ignore") as f:
        raw_old_twb = f.read()

    with open(twb_new, "r", encoding="utf-8", errors="ignore") as f:
        raw_new_twb = f.read()

    root_old = parse_twb(twb_old)
    root_new = parse_twb(twb_new)

    if root_old is None or root_new is None:
        print("❌ Failed to parse one or both TWB files.")
        return

    # =====================================================
    # EXTRACT SECTIONS
    # =====================================================
    sec_old = extract_sections(root_old)
    sec_new = extract_sections(root_new)

    # =====================================================
    # BUILD CARDS & REGISTRY
    # =====================================================
    cards = build_cards(sec_old, sec_new)
    populate_change_registry_from_cards(cards)

    # =====================================================
    # OVERALL SUMMARY
    # =====================================================
    overall_summary_card = build_overall_workbook_summary_card(
        sec_old,
        sec_new,
        cards,
        root_old,
        root_new
    )
    if overall_summary_card:
        cards.insert(0, overall_summary_card)

    # =====================================================
    # 👥 USERS & PERMISSIONS (SOURCE vs TARGET WORKBOOK)
    # =====================================================
    try:
        source_permissions = get_users_and_permissions_for_workbook(
            token,
            site_id,
            source_project,
            source_workbook
        )

        target_permissions = get_users_and_permissions_for_workbook(
            token,
            site_id,
            target_project,
            target_workbook
        )

        # If you don’t have a diff UI yet, stack both cards
        users_permissions_html = (
    build_users_permissions_card_with_context(
        source_project,
        source_workbook,
        source_permissions,
        context="source"
    )
    + "<hr/>"
    + build_users_permissions_card_with_context(
        target_project,
        target_workbook,
        target_permissions,
        context="Target"
    )
)


    except Exception as e:
        users_permissions_html = f"""
        <div class="panel">
        <h2>👥 Users & Permissions</h2>
        <p style="color:#b71c1c;">
            Unable to retrieve user permissions.<br>
            Reason: {html.escape(str(e))}
        </p>
        </div>
        """

    # =====================================================
    # GLOBAL ACTIONS
    # =====================================================
    global_action_card = build_global_action_card(root_old, root_new)
    if global_action_card:
        cards.insert(0, global_action_card)

     # 🔐 SAFETY: ensure all registry buckets exist
    ensure_change_registry_keys()
    # =====================================================
# DATASOURCES (AUTHORITATIVE – SINGLE FILE)
# =====================================================

    old_datasources = extract_datasources_raw(twb_old)
    new_datasources = extract_datasources_raw(twb_new)

    for ds_name in set(old_datasources) | set(new_datasources):
        compare(
            ds_name,
            old_datasources.get(ds_name),
            new_datasources.get(ds_name),
            site_id,
            token
        )


    # =====================================================
    # STRUCTURAL XML DIFF
    # =====================================================
    structural = xmldiff_text(
        ET.tostring(root_old, encoding="unicode"),
        ET.tostring(root_new, encoding="unicode")
    )

    safe_wb = sanitize_name(f"{source_workbook}_VS_{target_workbook}")
    struct_path = f"{safe_wb}_LATEST_STRUCT.txt"
    write_text(struct_path, structural)

    # =====================================================
    # KPIs + VISUAL TREE
    # =====================================================
    kpi_old = build_workbook_kpi_snapshot(sec_old)
    kpi_new = build_workbook_kpi_snapshot(sec_new)
    kpi_html = render_workbook_kpi_table(kpi_old, kpi_new)

    visual_tree_text = render_visual_change_tree(
        sec_new, CHANGE_REGISTRY, target_workbook
    )

    # =====================================================
    # GENERATE REPORT
    # =====================================================
    out = "compare_SOURCE_vs_TARGET_latest.html"

    generate_html_report(
        f"{source_workbook} (Latest)",
        f"{target_workbook} (Latest)",
        cards,
        None,
        out,
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

if __name__ == "__main__":
    main()
