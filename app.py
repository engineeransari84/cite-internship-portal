"""
CITE Internship Record Portal — Streamlit edition
==================================================
Centre for Internship, Training and Employment (CITE)

WHAT THIS IS
------------
A Streamlit rebuild of the original single-file HTML portal. The Excel
workbook itself IS the database: everything the app reads and writes lives
in one .xlsx file (two sheets: "Internships" and "Settings") inside a data
folder you choose in the sidebar. You can open that workbook directly in
Excel to inspect or bulk-edit it — just close it in Excel before editing
in the app again, since the app can't write to a file Excel has locked.

ACCESS MODEL
------------
Everyone who can open the app can browse records (read-only). Editing
(add / edit / delete / import / settings) requires the admin password,
set the first time in Settings. Default admin password: cite-admin
(change it immediately — see the Settings dialog, admin only).

Run with:  streamlit run app.py
"""

import csv
import hashlib
import io
import os
import re
import sys
from datetime import date, datetime

import altair as alt
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CONFIG
# ==============================================================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
# When frozen by PyInstaller, app.py itself lives inside the bundled
# _internal folder — anchor "next to the app" paths (remembered data folder,
# default data folder) to the .exe's own folder instead, so they're a normal,
# discoverable, writable location rather than tucked inside the bundle.
APP_ROOT = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else APP_DIR
POINTER_FILE = os.path.join(APP_ROOT, ".cite_last_datadir.txt")
WORKBOOK_NAME = "CITE-internships.xlsx"

CONFIG = {
    "owner": {
        "name": "[Your Name]",
        "title": "Convener, CITE (Centre for Internship, Training and Employment)",
        "inst": "[Your Institution]",
        "email": "[your-email@institution.edu]",
        "phone": "",
    },
    "default_admin_password": "cite-admin",
}

COLS = [  # (canonical header, internal field name) — order matters for the sheet
    ("Student Name", "name"),
    ("Roll Number", "roll"),
    ("Batch", "batch"),
    ("CNIC", "cnic"),
    ("Semester", "semester"),
    ("Organization", "org"),
    ("Start Date (YYYY-MM-DD)", "start"),
    ("End Date (YYYY-MM-DD)", "end"),
    ("Weeks (if no dates)", "manualWeeks"),
    ("Certificate On File (Yes/No)", "certOnFile"),
    ("Certificate Link", "certLink"),
]
SHEET_COLS = [("ID", "id")] + COLS
EXPORT_EXTRA = ["Internship Year", "Duration (weeks)", "Status"]
SETTINGS_KEYS = [
    "owner_name", "owner_title", "owner_inst", "owner_email", "owner_phone",
    "admin_hash", "reveal_cnic",
]

STATUS_LABEL = {"green": "6+ wks", "amber": "4-5 wks", "red": "<4 wks", "ongoing": "Ongoing"}
STATUS_EMOJI = {"green": "🟢", "amber": "🟡", "red": "🔴", "ongoing": "⚪"}
STATUS_TX = {"green": "#166b3c", "amber": "#8a5e00", "red": "#9a2b28", "ongoing": "#37507a"}
STATUS_BG = {"green": "#e7f4ec", "amber": "#fbf1d5", "red": "#fbe6e4", "ongoing": "#e9eef6"}


def sha256(s: str) -> str:
    return hashlib.sha256((s or "").encode("utf-8")).hexdigest()


# ==============================================================================
# 2. DATA FOLDER — remembered pointer + env override
# ==============================================================================
def load_last_folder():
    try:
        with open(POINTER_FILE, "r", encoding="utf-8") as f:
            p = f.read().strip()
            return p or None
    except Exception:
        return None


def save_last_folder(p):
    try:
        with open(POINTER_FILE, "w", encoding="utf-8") as f:
            f.write(p)
    except Exception:
        pass


def default_data_folder():
    return (
        os.environ.get("CITE_DATA_DIR")
        or load_last_folder()
        or os.path.join(APP_ROOT, "data")
    )


# ==============================================================================
# 3. WORKBOOK I/O — the .xlsx file IS the database
# ==============================================================================
def default_settings():
    return {
        "owner": dict(CONFIG["owner"]),
        "adminHash": sha256(CONFIG["default_admin_password"]),
        "revealCnic": False,
    }


def uid():
    import uuid
    return "r" + uuid.uuid4().hex[:12]


def to_iso(v):
    """Turn many possible date inputs into an ISO yyyy-mm-dd string (or '')."""
    if v is None or v == "":
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", s)
    if m:
        return f"{int(m.group(3)):04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    try:
        d = pd.to_datetime(s)
        if pd.isna(d):
            return ""
        return d.strftime("%Y-%m-%d")
    except Exception:
        return ""


def to_manual_weeks(v):
    """Parse a 'weeks (if no dates)' value; '' if blank/invalid/non-positive."""
    try:
        n = float(str(v).strip())
    except (TypeError, ValueError):
        return ""
    return n if n > 0 else ""


def cell_to_field(field, value):
    if value is None:
        return False if field == "certOnFile" else ""
    if field in ("start", "end"):
        return to_iso(value)
    if field == "certOnFile":
        if isinstance(value, bool):
            return value
        return bool(re.match(r"^(y|yes|true|1|on file)", str(value).strip(), re.I))
    if field == "manualWeeks":
        return to_manual_weeks(value)
    return str(value).strip()


def read_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    records = []
    if "Internships" in wb.sheetnames:
        ws = wb["Internships"]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            header = [str(h or "").strip() for h in rows[0]]
            idx = {h: i for i, h in enumerate(header)}
            for raw in rows[1:]:
                if raw is None or all((c is None or str(c).strip() == "") for c in raw):
                    continue
                rec = {}
                for hname, field in SHEET_COLS:
                    i = idx.get(hname)
                    v = raw[i] if (i is not None and i < len(raw)) else None
                    rec[field] = cell_to_field(field, v)
                if not rec.get("id"):
                    rec["id"] = uid()
                records.append(rec)

    settings = default_settings()
    if "Settings" in wb.sheetnames:
        ws = wb["Settings"]
        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < 2:
                continue
            k, v = row[0], row[1]
            if k not in SETTINGS_KEYS:
                continue
            if k == "reveal_cnic":
                settings["revealCnic"] = str(v).strip().lower() in ("true", "1", "yes")
            elif k == "admin_hash":
                if v:
                    settings["adminHash"] = str(v)
            elif k and k.startswith("owner_"):
                settings["owner"][k.replace("owner_", "")] = "" if v is None else str(v)
    return records, settings


def write_workbook(path, records, settings):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Internships"
    headers = [h for h, _ in SHEET_COLS]
    ws.append(headers)
    for r in records:
        row = []
        for _, f in SHEET_COLS:
            v = r.get(f, "")
            if f == "certOnFile":
                v = "Yes" if v else "No"
            row.append(v)
        ws.append(row)
    for i in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20

    ws2 = wb.create_sheet("Settings")
    ws2.append(["Key", "Value"])
    owner = settings.get("owner", {})
    kv = [
        ("owner_name", owner.get("name", "")),
        ("owner_title", owner.get("title", "")),
        ("owner_inst", owner.get("inst", "")),
        ("owner_email", owner.get("email", "")),
        ("owner_phone", owner.get("phone", "")),
        ("admin_hash", settings.get("adminHash", "")),
        ("reveal_cnic", "true" if settings.get("revealCnic") else "false"),
    ]
    for k, v in kv:
        ws2.append([k, v])
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 70

    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    wb.save(tmp)
    os.replace(tmp, path)


def ensure_workbook(path):
    if not os.path.exists(path):
        write_workbook(path, [], default_settings())


def sync_from_disk(force=False):
    path = st.session_state.data_path
    try:
        ensure_workbook(path)
        mtime = os.path.getmtime(path)
    except Exception as e:
        st.session_state.load_error = str(e)
        return
    st.session_state.load_error = None
    if force or st.session_state.get("loaded_mtime") != mtime:
        try:
            records, settings = read_workbook(path)
        except Exception as e:
            st.session_state.load_error = f"Could not read workbook: {e}"
            return
        st.session_state.records = records
        st.session_state.settings = settings
        st.session_state.loaded_mtime = mtime


def persist_local():
    try:
        write_workbook(st.session_state.data_path, st.session_state.records, st.session_state.settings)
        st.session_state.loaded_mtime = os.path.getmtime(st.session_state.data_path)
        return True
    except PermissionError:
        st.error("Could not save — the workbook seems to be open in Excel or another program. Close it there and try again.")
        return False
    except Exception as e:
        st.error(f"Could not save changes: {e}")
        return False


# ==============================================================================
# 4. DATA HELPERS — duration, status, CNIC, keys (ported 1:1 from the HTML app)
# ==============================================================================
def weeks_of(rec):
    """Effective duration in weeks. Dates win when both are present (most
    precise); otherwise a directly-entered week count (manualWeeks) is used,
    so records can be logged by weeks alone with no start/end date."""
    if rec.get("start") and rec.get("end"):
        try:
            a = datetime.strptime(rec["start"], "%Y-%m-%d").date()
            b = datetime.strptime(rec["end"], "%Y-%m-%d").date()
            if b >= a:
                days = (b - a).days + 1
                return round((days / 7) * 10) / 10
        except Exception:
            pass
    mw = to_manual_weeks(rec.get("manualWeeks"))
    if mw != "":
        return round(mw * 10) / 10
    return None


def status_of(rec):
    """Status from the effective duration. 'ongoing' means we simply have no
    duration yet (no end date and no entered weeks)."""
    w = weeks_of(rec)
    if w is None:
        return "ongoing"
    if w < 4:
        return "red"
    if w < 6:
        return "amber"
    return "green"


def cnic_digits(v):
    return re.sub(r"\D", "", str(v or ""))


def cnic_display(v, reveal):
    d = cnic_digits(v)
    if reveal:
        if len(d) == 13:
            return f"{d[0:5]}-{d[5:12]}-{d[12:]}"
        return d or "—"
    if len(d) <= 4:
        return d or "—"
    return "•••••-••••" + d[-4:-1] + "-" + d[-1:]


def batch_of(rec):
    b = str(rec.get("batch") or "").strip()
    if not b:
        m = re.match(r"^\s*(\d{2}\s*[A-Za-z]{1,5})", str(rec.get("roll") or ""))
        if m:
            b = m.group(1)
    return re.sub(r"\s+", "", b).upper()


def year_of(rec):
    iso = rec.get("start") or rec.get("end") or ""
    return iso[:4] if re.match(r"^\d{4}", iso) else ""


def rec_key(rec):
    """Dedup key: same student + same start date + same semester = same record.
    Semester is included so a student can have several date-less records — one
    per semester — without them merging into one."""
    return (
        f"{str(rec.get('roll') or '').strip().upper()}|{cnic_digits(rec.get('cnic'))}"
        f"|{rec.get('start') or ''}|{str(rec.get('semester') or '').strip().lower()}"
    )


def student_key(rec):
    return f"{str(rec.get('roll') or '').strip().upper()}|{cnic_digits(rec.get('cnic'))}"


def fmt_date(iso):
    if not iso:
        return "—"
    try:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
    except Exception:
        return iso
    return d.strftime("%d %b %Y")


def validate(rec):
    if not rec.get("name"):
        return "Student name is required."
    if not rec.get("roll"):
        return "Roll number is required."
    if not cnic_digits(rec.get("cnic")):
        return "CNIC is required."
    has_weeks = to_manual_weeks(rec.get("manualWeeks")) != ""
    if not rec.get("start") and not has_weeks:
        return "Enter a start date, or a number of weeks if you don't have dates."
    if rec.get("end") and rec.get("start") and rec["end"] < rec["start"]:
        return "End date cannot be before the start date."
    return None


def student_status(total_weeks, has_ongoing):
    if total_weeks >= 6:
        return "green"
    if total_weeks >= 4:
        return "amber"
    if total_weeks > 0:
        return "red"
    return "ongoing" if has_ongoing else "red"


def aggregate_students(recs):
    m = {}
    for r in recs:
        k = student_key(r)
        a = m.setdefault(k, {
            "key": k, "name": r.get("name") or "", "roll": r.get("roll") or "",
            "cnic": r.get("cnic") or "", "batch": batch_of(r),
            "count": 0, "totalWeeks": 0.0, "hasOngoing": False,
            "years": set(), "sems": set(),
        })
        a["count"] += 1
        w = weeks_of(r)
        if w is not None:
            a["totalWeeks"] += w
        else:
            a["hasOngoing"] = True
        y = year_of(r)
        if y:
            a["years"].add(y)
        if r.get("semester"):
            a["sems"].add(r["semester"])
        if r.get("name") and not a["name"]:
            a["name"] = r["name"]
        if not a["batch"]:
            a["batch"] = batch_of(r)
    out = []
    for a in m.values():
        a["totalWeeks"] = round(a["totalWeeks"] * 10) / 10
        a["status"] = student_status(a["totalWeeks"], a["hasOngoing"])
        a["years"] = sorted(a["years"])
        out.append(a)
    return out


def badge_html(status, extra=""):
    return (
        f'<span style="background:{STATUS_BG[status]};color:{STATUS_TX[status]};'
        f'padding:3px 9px;border-radius:999px;font-size:12.5px;font-weight:600;'
        f'white-space:nowrap">{STATUS_LABEL[status]}</span>{extra}'
    )


# ==============================================================================
# 5. IMPORT (merge, not overwrite) — mirrors the HTML app's merge logic
# ==============================================================================
def map_header(h):
    s = str(h or "").strip().lower()
    if "name" in s:
        return "name"
    if "roll" in s:
        return "roll"
    if re.search(r"batch|session|cohort", s):
        return "batch"
    if "cnic" in s:
        return "cnic"
    if "sem" in s:
        return "semester"
    if re.search(r"organi|company|firm", s):
        return "org"
    if "start" in s:
        return "start"
    if re.search(r"end|finish|to date", s):
        return "end"
    if re.search(r"week|duration", s):
        return "manualWeeks"
    if re.search(r"cert.*link|link", s):
        return "certLink"
    if "cert" in s:
        return "certOnFile"
    return None


def find_header_row(rows):
    best, best_score = 0, 0
    for i in range(min(len(rows), 10)):
        score = sum(1 for c in (rows[i] or []) if map_header(c))
        if score > best_score:
            best_score, best = score, i
    return best if best_score >= 2 else 0


def parse_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.read()
    if name.endswith(".csv"):
        text = data.decode("utf-8-sig", errors="replace")
        return [row for row in csv.reader(io.StringIO(text))]
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb[wb.sheetnames[0]]
    return [list(row) for row in ws.iter_rows(values_only=True)]


def merge_rows(rows, records):
    h = find_header_row(rows)
    if len(rows) < h + 2:
        return records, 0, 0, 0, [(0, "The file has no data rows below the header.")], []
    headers = [map_header(c) for c in rows[h]]
    existing = {rec_key(r): r for r in records}
    added = updated = dup = 0
    invalid = []
    touched = []
    for i in range(h + 1, len(rows)):
        raw = rows[i]
        if raw is None or all((c is None or str(c).strip() == "") for c in raw):
            continue
        rec = {"certOnFile": False, "certLink": ""}
        for idx, f in enumerate(headers):
            if not f:
                continue
            v = raw[idx] if idx < len(raw) else ""
            if f in ("start", "end"):
                v = to_iso(v)
            elif f == "certOnFile":
                v = bool(re.match(r"^(y|yes|true|1|on file)", str(v).strip(), re.I))
            elif f == "manualWeeks":
                v = to_manual_weeks(v)
            else:
                v = "" if v is None else str(v).strip()
            rec[f] = v
        err = validate(rec)
        if err:
            invalid.append((i + 1, err))
            continue
        key = rec_key(rec)
        if key in existing:
            cur = existing[key]
            changed = any((cur.get(f) or "") != (rec.get(f) or "") for f in
                          ["name", "batch", "semester", "org", "end", "manualWeeks", "certOnFile", "certLink"])
            if changed:
                cur.update(rec)
                updated += 1
                touched.append(cur)
            else:
                dup += 1
        else:
            nr = {"id": uid(), **rec}
            records.append(nr)
            existing[key] = nr
            added += 1
            touched.append(nr)
    return records, added, updated, dup, invalid, touched


# ==============================================================================
# 6. EXPORT + TEMPLATE
# ==============================================================================
def records_to_matrix(records):
    header = [h for h, _ in COLS] + EXPORT_EXTRA
    rows = [header]
    for r in records:
        row = []
        for h, f in COLS:
            if f == "certOnFile":
                row.append("Yes" if r.get("certOnFile") else "No")
            elif f == "batch":
                row.append(batch_of(r))
            else:
                row.append(r.get(f) or "")
        w = weeks_of(r)
        row += [year_of(r), (f"{w:.1f}" if w is not None else ""), STATUS_LABEL[status_of(r)]]
        rows.append(row)
    return rows


def template_matrix():
    example = ["Ali Raza", "21CE-045", "21CE", "41303-1234567-1", "5th",
               "Descon Engineering", "2025-06-02", "2025-07-13", "", "Yes", ""]
    return [[h for h, _ in COLS], example]


def matrix_to_csv_bytes(matrix):
    buf = io.StringIO()
    csv.writer(buf).writerows(matrix)
    return buf.getvalue().encode("utf-8-sig")


def matrix_to_xlsx_bytes(matrix, sheet_name="Internships"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in matrix:
        ws.append(row)
    if matrix:
        for i in range(1, len(matrix[0]) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 20
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================================================================
# 6b. CLOUD SYNC — optional shared database via Supabase REST
# When SUPABASE_URL / SUPABASE_ANON_KEY are set in Streamlit secrets, every
# session reads & writes the SAME records, so data entered by one teammate
# shows up for the others. Falls back to the local Excel workbook (section 3)
# when no cloud config is present — that's what the desktop .exe build uses.
# ==============================================================================
def cloud_config():
    try:
        url = str(st.secrets.get("SUPABASE_URL", "") or "").strip().rstrip("/")
        key = str(st.secrets.get("SUPABASE_ANON_KEY", "") or "").strip()
    except Exception:
        return None
    if not url or not key:
        return None
    return {
        "url": url,
        "key": key,
        "table": str(st.secrets.get("SUPABASE_TABLE", "internships") or "internships").strip(),
        "settings_table": str(st.secrets.get("SUPABASE_SETTINGS_TABLE", "app_settings") or "app_settings").strip(),
    }


def _cloud_headers(cfg, extra=None):
    h = {"apikey": cfg["key"], "Authorization": f"Bearer {cfg['key']}", "Content-Type": "application/json"}
    if extra:
        h.update(extra)
    return h


def rec_to_row(r):
    mw = to_manual_weeks(r.get("manualWeeks"))
    return {
        "id": r.get("id") or uid(),
        "name": r.get("name") or "", "roll": r.get("roll") or "",
        "batch": batch_of(r), "cnic": r.get("cnic") or "",
        "semester": r.get("semester") or "", "org": r.get("org") or "",
        "start_date": r.get("start") or "", "end_date": r.get("end") or "",
        "manual_weeks": (mw if mw != "" else None),
        "cert_on_file": bool(r.get("certOnFile")), "cert_link": r.get("certLink") or "",
    }


def row_to_rec(x):
    mw = x.get("manual_weeks")
    return {
        "id": x.get("id") or uid(), "name": x.get("name") or "", "roll": x.get("roll") or "",
        "batch": x.get("batch") or "", "cnic": x.get("cnic") or "",
        "semester": x.get("semester") or "", "org": x.get("org") or "",
        "start": x.get("start_date") or "", "end": x.get("end_date") or "",
        "manualWeeks": ("" if mw is None else mw),
        "certOnFile": bool(x.get("cert_on_file")), "certLink": x.get("cert_link") or "",
    }


def settings_to_row(settings):
    owner = settings.get("owner", {})
    return {
        "id": "default",
        "owner_name": owner.get("name", ""), "owner_title": owner.get("title", ""),
        "owner_inst": owner.get("inst", ""), "owner_email": owner.get("email", ""),
        "owner_phone": owner.get("phone", ""),
        "admin_hash": settings.get("adminHash", ""),
        "reveal_cnic": bool(settings.get("revealCnic")),
    }


def row_to_settings(x):
    s = default_settings()
    s["owner"] = {
        "name": x.get("owner_name") or "", "title": x.get("owner_title") or "",
        "inst": x.get("owner_inst") or "", "email": x.get("owner_email") or "",
        "phone": x.get("owner_phone") or "",
    }
    if x.get("admin_hash"):
        s["adminHash"] = x["admin_hash"]
    s["revealCnic"] = bool(x.get("reveal_cnic"))
    return s


def cloud_fetch_all(cfg):
    try:
        r = requests.get(f"{cfg['url']}/rest/v1/{cfg['table']}?select=*", headers=_cloud_headers(cfg), timeout=10)
        r.raise_for_status()
        return r.json()
    except Exception:
        return None


def cloud_upsert_rows(cfg, rows):
    if not rows:
        return True
    try:
        r = requests.post(
            f"{cfg['url']}/rest/v1/{cfg['table']}",
            headers=_cloud_headers(cfg, {"Prefer": "resolution=merge-duplicates,return=minimal"}),
            json=rows, timeout=10,
        )
        return r.ok
    except Exception:
        return False


def cloud_delete_row(cfg, rec_id):
    try:
        r = requests.delete(
            f"{cfg['url']}/rest/v1/{cfg['table']}?id=eq.{rec_id}",
            headers=_cloud_headers(cfg, {"Prefer": "return=minimal"}), timeout=10,
        )
        return r.ok or r.status_code == 404
    except Exception:
        return False


def cloud_delete_all(cfg):
    try:
        r = requests.delete(
            f"{cfg['url']}/rest/v1/{cfg['table']}?id=not.is.null",
            headers=_cloud_headers(cfg, {"Prefer": "return=minimal"}), timeout=10,
        )
        return r.ok
    except Exception:
        return False


def cloud_fetch_settings(cfg):
    try:
        r = requests.get(
            f"{cfg['url']}/rest/v1/{cfg['settings_table']}?id=eq.default&select=*",
            headers=_cloud_headers(cfg), timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
        return rows[0] if rows else None
    except Exception:
        return None


def cloud_upsert_settings(cfg, row):
    try:
        r = requests.post(
            f"{cfg['url']}/rest/v1/{cfg['settings_table']}",
            headers=_cloud_headers(cfg, {"Prefer": "resolution=merge-duplicates,return=minimal"}),
            json=[row], timeout=10,
        )
        return r.ok
    except Exception:
        return False


def sync_from_cloud():
    cfg = st.session_state.cloud_cfg
    rows = cloud_fetch_all(cfg)
    if rows is None:
        st.session_state.load_error = (
            "Could not reach the shared cloud database — check the Supabase URL/key in "
            "this app's secrets, or your connection. Showing the last-loaded data."
        )
        st.session_state.setdefault("records", [])
        st.session_state.setdefault("settings", default_settings())
        return
    st.session_state.load_error = None
    st.session_state.records = [row_to_rec(x) for x in rows]
    srow = cloud_fetch_settings(cfg)
    st.session_state.settings = row_to_settings(srow) if srow else default_settings()


def persist(touched=None, deleted_ids=None, settings_changed=False, wipe=False):
    """Save pending changes. Local mode rewrites the whole workbook (cheap,
    already-tested path). Cloud mode pushes only what actually changed."""
    if not USE_CLOUD:
        return persist_local()
    cfg = st.session_state.cloud_cfg
    ok = True
    if wipe:
        ok = cloud_delete_all(cfg) and ok
    if deleted_ids:
        for did in deleted_ids:
            ok = cloud_delete_row(cfg, did) and ok
    if touched:
        ok = cloud_upsert_rows(cfg, [rec_to_row(r) for r in touched]) and ok
    if settings_changed:
        ok = cloud_upsert_settings(cfg, settings_to_row(st.session_state.settings)) and ok
    if not ok:
        st.error("Could not save to the shared cloud database — check your connection and try again.")
        return False
    return True


# ==============================================================================
# 7. SESSION STATE + PAGE SETUP
# ==============================================================================
st.set_page_config(page_title="CITE · Internship Records", page_icon="🎓", layout="wide")

CLOUD_CFG = cloud_config()
USE_CLOUD = CLOUD_CFG is not None

_defaults = {
    "data_folder": default_data_folder(),
    "is_admin": False,
    "scope_mode": "all",
    "scope_value": None,
    "table_mode": "record",
    "chart_group": "semester",
    "editing_id": None,
    "show_editor": False,
    "show_settings": False,
    "confirm_delete_id": None,
    "profile_target": None,
    "import_summary": None,
    "loaded_mtime": None,
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v
st.session_state.cloud_cfg = CLOUD_CFG

if USE_CLOUD:
    sync_from_cloud()
else:
    st.session_state.data_path = os.path.join(st.session_state.data_folder, WORKBOOK_NAME)
    sync_from_disk()
if "records" not in st.session_state:
    st.session_state.records = []
if "settings" not in st.session_state:
    st.session_state.settings = default_settings()


def scoped_records():
    recs = st.session_state.records
    if st.session_state.scope_mode == "batch" and st.session_state.scope_value:
        return [r for r in recs if batch_of(r) == st.session_state.scope_value]
    if st.session_state.scope_mode == "year" and st.session_state.scope_value:
        return [r for r in recs if year_of(r) == st.session_state.scope_value]
    return list(recs)


def distinct_batches():
    return sorted({batch_of(r) for r in st.session_state.records if batch_of(r)})


def distinct_years():
    return sorted({year_of(r) for r in st.session_state.records if year_of(r)})


# ==============================================================================
# 8. DIALOGS
# ==============================================================================
@st.dialog("Internship record", width="large")
def record_dialog(edit_id=None):
    existing = next((r for r in st.session_state.records if r["id"] == edit_id), None) if edit_id else None
    st.caption("Edit internship record" if existing else "Add internship record")

    c1, c2 = st.columns(2)
    name = c1.text_input("Student name *", value=existing.get("name", "") if existing else "")
    roll = c2.text_input("Roll number *", value=existing.get("roll", "") if existing else "", key="e_roll")
    derived = batch_of({"roll": roll})
    batch_hint = f'Leave blank to use "{derived}" (from the roll number).' if derived else "Leave blank to derive from the roll number prefix."
    batch = c1.text_input("Batch", value=existing.get("batch", "") if existing else "", help=batch_hint, placeholder="e.g. 21CE")
    cnic = c2.text_input("CNIC * (13 digits)", value=existing.get("cnic", "") if existing else "", placeholder="e.g. 41303-1234567-1")
    semester = c1.text_input("Semester", value=existing.get("semester", "") if existing else "", placeholder="e.g. 5th")
    org = c2.text_input("Organization", value=existing.get("org", "") if existing else "")

    def _parse(iso):
        try:
            return datetime.strptime(iso, "%Y-%m-%d").date()
        except Exception:
            return None

    start_default = _parse(existing.get("start")) if existing else None
    ongoing_default = bool(existing) and bool(existing.get("start")) and not existing.get("end")
    ongoing = st.checkbox("Ongoing — no end date yet", value=ongoing_default)

    c3, c4 = st.columns(2)
    start = c3.date_input("Start date", value=start_default, format="YYYY-MM-DD")
    end = None
    if not ongoing:
        end_default = _parse(existing.get("end")) if existing else None
        end = c4.date_input("End date", value=end_default, format="YYYY-MM-DD")
    else:
        c4.text_input("End date", value="(ongoing)", disabled=True)

    weeks_default = existing.get("manualWeeks") if existing else ""
    weeks_default = float(weeks_default) if weeks_default not in ("", None) else 0.0
    manual_weeks = st.number_input(
        "Weeks (if no dates)", min_value=0.0, step=0.5, value=weeks_default,
        help="Don't have exact dates? Log the internship by a week count instead — leave Start date blank.",
    )

    c5, c6 = st.columns(2)
    cert = c5.selectbox("Certificate on file", ["No", "Yes"], index=1 if (existing and existing.get("certOnFile")) else 0)
    certlink = c6.text_input("Certificate link (optional)", value=existing.get("certLink", "") if existing else "", placeholder="Drive / OneDrive URL")

    st.caption("Duration and status colour are calculated automatically from the dates (or the week count, if no dates).")

    b1, b2 = st.columns(2)
    if b1.button("Cancel", use_container_width=True):
        st.session_state.show_editor = False
        st.rerun()
    if b2.button("Save record", type="primary", use_container_width=True):
        rec = {
            "name": name.strip(), "roll": roll.strip(),
            "batch": re.sub(r"\s+", "", batch.strip()).upper(),
            "cnic": cnic.strip(), "semester": semester.strip(), "org": org.strip(),
            "start": to_iso(start) if start else "", "end": "" if (ongoing or not end) else to_iso(end),
            "manualWeeks": manual_weeks if manual_weeks > 0 else "",
            "certOnFile": cert == "Yes", "certLink": certlink.strip(),
        }
        err = validate(rec)
        if err:
            st.error(err)
        else:
            if existing:
                existing.update(rec)
                saved = existing
            else:
                saved = {"id": uid(), **rec}
                st.session_state.records.append(saved)
            if persist(touched=[saved]):
                st.session_state.show_editor = False
                st.rerun()


@st.dialog("Student profile", width="large")
def profile_dialog(roll, cnic_d):
    recs = [r for r in st.session_state.records
            if str(r.get("roll", "")).strip().upper() == roll.strip().upper()
            and cnic_digits(r.get("cnic")) == cnic_d]
    recs = sorted(recs, key=lambda r: r.get("start") or "")
    if not recs:
        st.write("No records found for this student.")
        if st.button("Close"):
            st.session_state.profile_target = None
            st.rerun()
        return

    s = recs[0]
    total_weeks = round(sum((weeks_of(r) or 0) for r in recs) * 10) / 10
    has_ongoing = any(weeks_of(r) is None for r in recs)
    overall = student_status(total_weeks, has_ongoing)
    batch = batch_of(s)
    years = sorted({year_of(r) for r in recs if year_of(r)})

    st.markdown(f"### {s.get('name') or '—'}  `{s.get('roll') or ''}`")
    st.markdown(badge_html(overall, " overall"), unsafe_allow_html=True)
    meta = f"CNIC {cnic_display(s.get('cnic'), st.session_state.settings.get('revealCnic'))}"
    if batch:
        meta += f" · Batch {batch}"
    if years:
        meta += f" · Internship year{'s' if len(years) > 1 else ''} {', '.join(years)}"
    st.caption(meta)

    m1, m2, m3 = st.columns(3)
    m1.metric("Internships", len(recs))
    m2.metric("Total weeks", f"{total_weeks:.1f}" + (" +ongoing" if has_ongoing else ""))
    m3.metric("Semesters covered", len({r.get("semester") for r in recs if r.get("semester")}))
    st.caption("Total weeks add up every internship below — including ones imported from different sheets.")

    for r in recs:
        w = weeks_of(r)
        stt = status_of(r)
        with st.container(border=True):
            t1, t2 = st.columns([3, 2])
            t1.markdown(f"**{r.get('org') or 'Organization not set'}**")
            t2.markdown(badge_html(stt, f" {w:.1f} wks" if w is not None else ""), unsafe_allow_html=True)
            st.caption(f"Semester: {r.get('semester') or '—'}   |   "
                       f"Dates: {fmt_date(r.get('start'))} – {fmt_date(r.get('end')) if r.get('end') else 'ongoing'}")
            if r.get("certLink"):
                st.markdown(f"[View certificate ↗]({r['certLink']})")
            elif r.get("certOnFile"):
                st.caption("Certificate on file")
            else:
                st.caption("Certificate not on file")

    if st.button("Close", key="profile_close"):
        st.session_state.profile_target = None
        st.rerun()


@st.dialog("Settings", width="large")
def settings_dialog():
    s = st.session_state.settings
    st.markdown("##### Owner & contact (shown in footer)")
    c1, c2 = st.columns(2)
    name = c1.text_input("Name", value=s["owner"].get("name", ""))
    title = c2.text_input("Title", value=s["owner"].get("title", ""))
    inst = c1.text_input("Institution", value=s["owner"].get("inst", ""))
    email = c2.text_input("Contact email", value=s["owner"].get("email", ""))
    phone = c1.text_input("Phone", value=s["owner"].get("phone", ""))

    st.markdown("##### Privacy")
    reveal = st.checkbox("Show full CNIC numbers (default: masked to last 4 digits)", value=bool(s.get("revealCnic")))

    st.markdown("##### Change admin password")
    c3, c4 = st.columns(2)
    p1 = c3.text_input("New password", type="password", key="s_pass1")
    p2 = c4.text_input("Confirm password", type="password", key="s_pass2")

    st.markdown("---")
    st.markdown(":red[**Danger zone**]")
    st.caption("This clears data stored in the workbook. Export a backup first.")
    confirm_wipe = st.checkbox("I understand — erase ALL records in this workbook")
    d1, d2 = st.columns(2)
    if d1.button("Erase all records", disabled=not confirm_wipe):
        st.session_state.records = []
        if persist(wipe=True):
            st.success("All records erased.")
            st.rerun()
    if d2.button("Save settings", type="primary"):
        ok = True
        if p1 or p2:
            if len(p1) < 4:
                st.error("Password must be at least 4 characters.")
                ok = False
            elif p1 != p2:
                st.error("Passwords do not match.")
                ok = False
            else:
                s["adminHash"] = sha256(p1)
        if ok:
            s["owner"] = {
                "name": name.strip(), "title": title.strip(), "inst": inst.strip(),
                "email": email.strip(), "phone": phone.strip(),
            }
            s["revealCnic"] = reveal
            if persist(settings_changed=True):
                st.session_state.show_settings = False
                st.success("Settings saved.")
                st.rerun()

    if st.button("Close", key="settings_close"):
        st.session_state.show_settings = False
        st.rerun()


# ==============================================================================
# 9. SIDEBAR — data location + access
# ==============================================================================
with st.sidebar:
    st.markdown("## CITE")
    st.caption("Internship Records — Streamlit edition")

    st.markdown("### Data location")
    if USE_CLOUD:
        st.success(f"☁️ Connected — shared cloud database (`{st.session_state.cloud_cfg['table']}`)")
        st.caption("Everyone with this app's URL reads and writes the same live records — no per-device setup.")
        if st.button("Refresh now", use_container_width=True):
            sync_from_cloud()
            st.rerun()
        if st.session_state.get("load_error"):
            st.error(st.session_state.load_error)
    else:
        st.caption("The workbook below is the live database — the app reads and writes it directly.")
        folder_input = st.text_input("Data folder", value=st.session_state.data_folder)
        if st.button("Use this folder", use_container_width=True):
            st.session_state.data_folder = folder_input.strip() or default_data_folder()
            save_last_folder(st.session_state.data_folder)
            st.session_state.data_path = os.path.join(st.session_state.data_folder, WORKBOOK_NAME)
            sync_from_disk(force=True)
            st.rerun()
        st.code(st.session_state.data_path, language=None)
        if st.session_state.get("load_error"):
            st.error(st.session_state.load_error)

    st.markdown("### Access")
    if st.session_state.is_admin:
        st.success("Admin mode — editing enabled")
        a1, a2 = st.columns(2)
        if a1.button("Settings", use_container_width=True):
            st.session_state.show_settings = True
            st.rerun()
        if a2.button("Log out", use_container_width=True):
            st.session_state.is_admin = False
            st.rerun()
    else:
        st.info("Read-only mode")
        with st.form("login_form", clear_on_submit=True):
            pw = st.text_input("Admin password", type="password")
            submitted = st.form_submit_button("Log in", use_container_width=True)
        if submitted:
            if sha256(pw) == st.session_state.settings.get("adminHash"):
                st.session_state.is_admin = True
                if st.session_state.settings.get("adminHash") == sha256(CONFIG["default_admin_password"]):
                    st.session_state.show_default_pw_warning = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        if st.session_state.pop("show_default_pw_warning", False):
            pass  # warning shown after rerun below via main body

    if st.session_state.is_admin:
        st.markdown("### Import (merge, not overwrite)")
        with st.form("import_form", clear_on_submit=True):
            uploaded = st.file_uploader("Excel or CSV file", type=["xlsx", "xls", "csv"])
            do_import = st.form_submit_button("Import", use_container_width=True)
        if do_import and uploaded is not None:
            try:
                rows = parse_uploaded_file(uploaded)
                records, added, updated, dup, invalid, touched = merge_rows(rows, st.session_state.records)
                st.session_state.records = records
                if persist(touched=touched):
                    st.session_state.import_summary = {
                        "added": added, "updated": updated, "dup": dup,
                        "invalid": invalid, "total": len(records),
                    }
                    st.rerun()
            except Exception as e:
                st.error(f"Could not read that file: {e}")

# ==============================================================================
# 10. MAIN — top bar
# ==============================================================================
top1, top2 = st.columns([4, 2])
with top1:
    st.markdown("# CITE · Internship Records")
with top2:
    st.markdown(
        f"<div style='text-align:right;padding-top:18px'>"
        f"<span style='background:{'#2f7d4f' if st.session_state.is_admin else '#eef1f6'};"
        f"color:{'#eafaf0' if st.session_state.is_admin else '#5f6b7c'};"
        f"padding:4px 12px;border-radius:999px;font-size:12.5px;font-weight:600'>"
        f"{'Admin' if st.session_state.is_admin else 'Read-only'}</span></div>",
        unsafe_allow_html=True,
    )
if st.session_state.is_admin and st.session_state.settings.get("adminHash") == sha256(CONFIG["default_admin_password"]):
    st.warning("You're using the default admin password (`cite-admin`). Open **Settings** in the sidebar to change it.")

if st.session_state.get("import_summary"):
    summ = st.session_state.import_summary
    st.info(
        f"Import complete — the workbook now holds **{summ['total']}** record(s). "
        f"Added {summ['added']}, updated {summ['updated']}, duplicates skipped {summ['dup']}, "
        f"invalid {len(summ['invalid'])}."
    )
    if summ["invalid"]:
        with st.expander(f"{len(summ['invalid'])} row(s) skipped — fix and re-import"):
            for row_no, err in summ["invalid"][:30]:
                st.write(f"Row {row_no}: {err}")
    if st.button("Dismiss import summary"):
        st.session_state.import_summary = None
        st.rerun()

# ---- scope bar: All / By batch / By internship year -------------------------
st.markdown("#### Viewing")
mode_labels = {"all": "All records", "batch": "By batch", "year": "By internship year"}
mode_choice = st.radio(
    "Scope", list(mode_labels.values()),
    index=list(mode_labels.keys()).index(st.session_state.scope_mode),
    horizontal=True, label_visibility="collapsed",
)
new_mode = [k for k, v in mode_labels.items() if v == mode_choice][0]
if new_mode != st.session_state.scope_mode:
    st.session_state.scope_mode = new_mode
    st.session_state.scope_value = None
    st.rerun()

if st.session_state.scope_mode != "all":
    noun = "batch" if st.session_state.scope_mode == "batch" else "internship year"
    vals = distinct_batches() if st.session_state.scope_mode == "batch" else distinct_years()
    if not vals:
        st.caption(f"No {noun} values yet — import or add records first.")
    else:
        options = [f"All {noun}s"] + vals
        cur = st.session_state.scope_value if st.session_state.scope_value in vals else options[0]
        picked = st.selectbox(f"Focus on one {noun}", options, index=options.index(cur))
        new_val = None if picked == options[0] else picked
        if new_val != st.session_state.scope_value:
            st.session_state.scope_value = new_val
            st.rerun()

scoped = scoped_records()

# ---- overview -----------------------------------------------------------------
students = len({student_key(r) for r in scoped})
counts = {"green": 0, "amber": 0, "red": 0, "ongoing": 0}
for r in scoped:
    counts[status_of(r)] += 1
o1, o2, o3, o4, o5, o6 = st.columns(6)
o1.metric("Students", students)
o2.metric("Internship records", len(scoped))
o3.metric("🟢 Green · 6+ wks", counts["green"])
o4.metric("🟡 Amber · 4-5 wks", counts["amber"])
o5.metric("🔴 Red · under 4 wks", counts["red"])
o6.metric("⚪ Ongoing", counts["ongoing"])

# ---- chart ----------------------------------------------------------------
st.markdown("#### Records by group — status mix")
group_labels = {"semester": "semester", "batch": "batch", "year": "internship year"}
group_choice = st.selectbox(
    "Group chart by", list(group_labels.values()),
    index=list(group_labels.keys()).index(st.session_state.chart_group),
)
st.session_state.chart_group = [k for k, v in group_labels.items() if v == group_choice][0]


def group_of(r):
    if st.session_state.chart_group == "batch":
        return batch_of(r) or "—"
    if st.session_state.chart_group == "year":
        return year_of(r) or "—"
    return r.get("semester") or "—"


if scoped:
    chart_df = pd.DataFrame([{"group": group_of(r), "status": status_of(r)} for r in scoped])
    counts_df = chart_df.groupby(["group", "status"]).size().reset_index(name="count")
    order = ["green", "amber", "red", "ongoing"]
    chart = alt.Chart(counts_df).mark_bar().encode(
        y=alt.Y("group:N", sort="-x", title=None),
        x=alt.X("count:Q", title="Records", stack="zero"),
        color=alt.Color("status:N", scale=alt.Scale(domain=order, range=[STATUS_TX[s] for s in order]),
                         legend=alt.Legend(title="Status")),
        order=alt.Order("status:N"),
        tooltip=["group", "status", "count"],
    ).properties(height=max(120, 32 * counts_df["group"].nunique() + 40))
    st.altair_chart(chart, use_container_width=True)
else:
    st.caption("No records to chart yet.")

# ---- pie charts: finished vs remaining, status mix, by organization --------
if scoped:
    agg_scoped = aggregate_students(scoped)
    finished = sum(1 for a in agg_scoped if a["status"] == "green")
    not_finished = sum(1 for a in agg_scoped if a["status"] in ("amber", "red"))
    stu_ongoing = sum(1 for a in agg_scoped if a["status"] == "ongoing")

    org_counts = {}
    for r in scoped:
        o = (r.get("org") or "").strip() or "—"
        org_counts[o] = org_counts.get(o, 0) + 1
    org_items = sorted(org_counts.items(), key=lambda kv: -kv[1])
    if len(org_items) > 6:
        top, rest = org_items[:6], org_items[6:]
        org_items = top + [("Others", sum(v for _, v in rest))]

    def donut(labels_values, colors, key):
        df = pd.DataFrame(labels_values, columns=["label", "value"])
        df = df[df["value"] > 0]
        if df.empty:
            st.caption("No data yet.")
            return
        chart = alt.Chart(df).mark_arc(innerRadius=45, outerRadius=75).encode(
            theta=alt.Theta("value:Q"),
            color=alt.Color("label:N", scale=alt.Scale(domain=df["label"].tolist(), range=colors[:len(df)]),
                             legend=alt.Legend(title=None, orient="right")),
            tooltip=["label", "value"],
        ).properties(height=170)
        st.altair_chart(chart, use_container_width=True, key=key)

    students_reg = len(agg_scoped)
    orgs_n = len({o for o in org_counts if o != "—"})
    st.caption(f"{students_reg} student{'s' if students_reg != 1 else ''} registered · {orgs_n} organization{'s' if orgs_n != 1 else ''}")
    p1, p2, p3 = st.columns(3)
    with p1:
        st.markdown("**Students — finished vs. remaining**")
        donut(
            [("Finished (6+ wks)", finished), ("Yet to finish", not_finished), ("Ongoing / no duration", stu_ongoing)],
            [STATUS_TX["green"], STATUS_TX["amber"], STATUS_TX["ongoing"]], "pie_students",
        )
    with p2:
        st.markdown("**Internship records by status**")
        rc = {"green": 0, "amber": 0, "red": 0, "ongoing": 0}
        for r in scoped:
            rc[status_of(r)] += 1
        donut(
            [("Green (6+ wks)", rc["green"]), ("Amber (4-5 wks)", rc["amber"]),
             ("Red (<4 wks)", rc["red"]), ("Ongoing", rc["ongoing"])],
            [STATUS_TX["green"], STATUS_TX["amber"], STATUS_TX["red"], STATUS_TX["ongoing"]], "pie_status",
        )
    with p3:
        st.markdown("**Records by organization**")
        palette = ["#4f6b93", "#1f9d57", "#c98a00", "#d1443f", "#7a5ea8", "#2c8f9e", "#8793a3"]
        donut(org_items, palette, "pie_orgs")

# ---- toolbar: search / filters / view mode ---------------------------------
st.markdown("#### Records")
tb1, tb2, tb3, tb4 = st.columns([3, 2, 2, 2])
search = tb1.text_input("Search name or roll number…", label_visibility="collapsed", placeholder="Search name or roll number…")
sems = sorted({r.get("semester") for r in scoped if r.get("semester")})
orgs = sorted({r.get("org") for r in scoped if r.get("org")})
f_sem = tb2.selectbox("Semester", ["All semesters"] + sems, label_visibility="collapsed")
f_org = tb3.selectbox("Organization", ["All organizations"] + orgs, label_visibility="collapsed")
f_status = tb4.selectbox(
    "Status", ["All statuses", "green", "amber", "red", "ongoing"],
    format_func=lambda v: v if v in ("All statuses",) else f"{STATUS_EMOJI[v]} {STATUS_LABEL[v]}",
    label_visibility="collapsed",
)
f_sem = "" if f_sem == "All semesters" else f_sem
f_org = "" if f_org == "All organizations" else f_org
f_status = "" if f_status == "All statuses" else f_status

view_choice = st.radio(
    "View", ["By record", "By student"],
    index=0 if st.session_state.table_mode == "record" else 1,
    horizontal=True,
)
st.session_state.table_mode = "record" if view_choice == "By record" else "student"

# ---- export / template / add -----------------------------------------------
stamp = date.today().isoformat()
ex1, ex2, ex3, ex4 = st.columns(4)
ex1.download_button(
    "Export CSV", data=matrix_to_csv_bytes(records_to_matrix(st.session_state.records)),
    file_name=f"CITE-internships-{stamp}.csv", mime="text/csv", use_container_width=True,
)
ex2.download_button(
    "Export Excel", data=matrix_to_xlsx_bytes(records_to_matrix(st.session_state.records)),
    file_name=f"CITE-internships-{stamp}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
)
if st.session_state.is_admin:
    ex3.download_button(
        "Template", data=matrix_to_xlsx_bytes(template_matrix(), "Internships"),
        file_name="CITE-internship-template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
    )
    if ex4.button("+ Add record", type="primary", use_container_width=True):
        st.session_state.editing_id = None
        st.session_state.show_editor = True
        st.rerun()
else:
    ex3.caption("Log in as admin (sidebar) to add, edit, import, or change settings.")

st.caption("To print, use your browser's Print (Ctrl/Cmd+P) — choose 'Save as PDF' for a PDF copy.")

# ==============================================================================
# 11. TABLE
# ==============================================================================
if st.session_state.table_mode == "record":
    rows = []
    for r in scoped:
        rr = dict(r)
        rr["batch"] = batch_of(r)
        rr["weeks"] = weeks_of(r)
        rr["status"] = status_of(r)
        rows.append(rr)
    q = (search or "").strip().lower()
    if q:
        rows = [r for r in rows if q in (r.get("name") or "").lower() or q in (r.get("roll") or "").lower()]
    if f_sem:
        rows = [r for r in rows if (r.get("semester") or "") == f_sem]
    if f_org:
        rows = [r for r in rows if (r.get("org") or "") == f_org]
    if f_status:
        rows = [r for r in rows if r.get("status") == f_status]

    if not rows:
        st.info(
            "No records match your filters." if scoped else
            ("No records in this view — switch the scope above." if st.session_state.records else
             ("Import an Excel/CSV file or add a record to get started." if st.session_state.is_admin else
              "Ask a CITE admin to import the internship data."))
        )
        selected_row = None
    else:
        reveal = st.session_state.settings.get("revealCnic")
        display_df = pd.DataFrame([{
            "Student": r.get("name") or "—",
            "Roll no.": r.get("roll") or "—",
            "Batch": r.get("batch") or "—",
            "CNIC": cnic_display(r.get("cnic"), reveal),
            "Sem.": r.get("semester") or "—",
            "Organization": r.get("org") or "—",
            "Start": fmt_date(r.get("start")),
            "End": fmt_date(r.get("end")) if r.get("end") else "Ongoing",
            "Weeks": (f"{r['weeks']:.1f}" if r.get("weeks") is not None else "—"),
            "Status": f"{STATUS_EMOJI[r['status']]} {STATUS_LABEL[r['status']]}",
            "Cert.": ("View ↗" if r.get("certLink") else ("On file" if r.get("certOnFile") else "—")),
        } for r in rows])

        event = st.dataframe(
            display_df, use_container_width=True, hide_index=True,
            on_select="rerun", selection_mode="single-row", key="record_table",
        )
        sel = event.selection.rows if event and event.selection else []
        selected_row = rows[sel[0]] if sel else None

    if selected_row:
        rc1, rc2, rc3, rc4 = st.columns([1, 1, 1, 3])
        if rc1.button("View profile", use_container_width=True):
            st.session_state.profile_target = (selected_row["roll"], cnic_digits(selected_row["cnic"]))
            st.rerun()
        if st.session_state.is_admin:
            if rc2.button("Edit", use_container_width=True):
                st.session_state.editing_id = selected_row["id"]
                st.session_state.show_editor = True
                st.rerun()
            if rc3.button("Delete", use_container_width=True):
                st.session_state.confirm_delete_id = selected_row["id"]
                st.rerun()
    elif rows:
        st.caption("Select a row above to view a profile" + (", edit, or delete." if st.session_state.is_admin else "."))

else:  # By student
    agg = aggregate_students(scoped)
    q = (search or "").strip().lower()
    if q:
        agg = [a for a in agg if q in (a.get("name") or "").lower() or q in (a.get("roll") or "").lower()]
    if f_status:
        agg = [a for a in agg if a["status"] == f_status]

    if not agg:
        st.info(
            "No students in this view." if st.session_state.records else
            ("Import an Excel/CSV file or add a record to get started." if st.session_state.is_admin else
             "Ask a CITE admin to import the internship data.")
        )
        selected_student = None
    else:
        display_df = pd.DataFrame([{
            "Student": a["name"] or "—",
            "Roll no.": a["roll"] or "—",
            "Batch": a["batch"] or "—",
            "Internships": a["count"],
            "Total weeks": f"{a['totalWeeks']:.1f}" + (" +ongoing" if a["hasOngoing"] else ""),
            "Overall": f"{STATUS_EMOJI[a['status']]} {STATUS_LABEL[a['status']]}",
            "Internship years": ", ".join(a["years"]) if a["years"] else "—",
        } for a in agg])
        event = st.dataframe(
            display_df, use_container_width=True, hide_index=True,
            on_select="rerun", selection_mode="single-row", key="student_table",
        )
        sel = event.selection.rows if event and event.selection else []
        selected_student = agg[sel[0]] if sel else None

    if selected_student:
        if st.button("View profile", use_container_width=False):
            st.session_state.profile_target = (selected_student["roll"], cnic_digits(selected_student["cnic"]))
            st.rerun()
    elif agg:
        st.caption("Select a row above to view that student's full profile.")

# ---- delete confirmation ----------------------------------------------------
if st.session_state.get("confirm_delete_id"):
    rid = st.session_state.confirm_delete_id
    rec = next((r for r in st.session_state.records if r["id"] == rid), None)
    if rec:
        st.warning(f"Delete this internship record for **{rec.get('name') or 'this student'}**? This cannot be undone.")
        cd1, cd2 = st.columns(2)
        if cd1.button("Confirm delete", type="primary"):
            st.session_state.records = [r for r in st.session_state.records if r["id"] != rid]
            if persist(deleted_ids=[rid]):
                st.session_state.confirm_delete_id = None
                st.rerun()
        if cd2.button("Cancel delete"):
            st.session_state.confirm_delete_id = None
            st.rerun()
    else:
        st.session_state.confirm_delete_id = None

# ---- open dialogs (must be called, not just defined, to render) -------------
if st.session_state.show_editor:
    record_dialog(st.session_state.editing_id)
if st.session_state.show_settings:
    settings_dialog()
if st.session_state.profile_target:
    profile_dialog(*st.session_state.profile_target)

# ==============================================================================
# 12. FOOTER
# ==============================================================================
st.divider()
owner = st.session_state.settings.get("owner", {})
foot = f"Developed and maintained by **{owner.get('name') or '—'}**"
if owner.get("title"):
    foot += f", {owner['title']}"
if owner.get("inst"):
    foot += f" · {owner['inst']}"
st.caption(foot + ".")
contact_bits = []
if owner.get("email"):
    contact_bits.append(f"**{owner['email']}**")
if owner.get("phone"):
    contact_bits.append(f"**{owner['phone']}**")
st.caption("For licensing or usage inquiries, contact " + (" · ".join(contact_bits) if contact_bits else "**—**") + ".")
st.caption(f"© {date.today().year} {owner.get('name') or 'CITE'}. All rights reserved.")
