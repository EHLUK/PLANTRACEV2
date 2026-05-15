"""
P6 XER Project Manager Planning Tool
=====================================
A Streamlit app for interrogating Primavera P6 XER schedules
without needing to open P6. Designed for Project Managers.
"""

import io
import re
import math
import warnings
from collections import defaultdict, deque
from datetime import datetime, timedelta

import networkx as nx
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# -----------------------------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="PlanTrace",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# CUSTOM CSS  --  PlanTrace v3 Professional Control Centre
# Brand: Navy=#071827  Amber=#F5A623  Background=#F3F5F7
# -----------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

/* ========== HIDE STREAMLIT CHROME ========== */
#MainMenu,footer,header{visibility:hidden}
[data-testid="stToolbar"]{display:none!important}
[data-testid="stDecoration"]{display:none!important}
[data-testid="stStatusWidget"]{display:none!important}
.stDeployButton{display:none!important}
section[data-testid="stSidebarNav"]{display:none!important}

/* ========== ROOT ========== */
html,body,.stApp{background:#F3F5F7!important;font-family:'Inter','Segoe UI',Arial,sans-serif!important}
[data-testid="stAppViewContainer"]{background:#F3F5F7!important}
[data-testid="block-container"]{
    padding-top:0!important;padding-bottom:32px!important;
    padding-left:28px!important;padding-right:28px!important;
    max-width:100%!important}

/* ========== SIDEBAR ========== */
[data-testid="stSidebar"]{
    background:#071827!important;
    border-right:1px solid #0B2438;
    min-width:220px!important;max-width:236px!important}
[data-testid="stSidebar"] *{color:#4B6275!important;font-family:'Inter','Segoe UI',Arial,sans-serif!important}
[data-testid="stSidebar"] input{background:#0B2438!important;color:#8FA3B1!important;
    border:1px solid #102A43!important;border-radius:5px!important;font-size:12px!important}
[data-testid="stSidebar"] label{color:#3A5060!important;font-size:11px!important}
[data-testid="stSidebar"] .stSlider div[data-testid="stSliderThumb"]{background:#F5A623!important}
[data-testid="stSidebar"] .stSlider [data-testid="stSliderTrack"]{background:#0B2438!important}

/* ---- Sidebar nav buttons ---- */
[data-testid="stSidebar"] .stButton>button{
    background:transparent!important;color:#4B6275!important;
    border:none!important;border-radius:0!important;
    text-align:left!important;font-size:13px!important;font-weight:500!important;
    padding:8px 14px!important;width:100%!important;
    box-shadow:none!important;letter-spacing:0.1px;
    transition:background 0.1s ease,color 0.1s ease}
[data-testid="stSidebar"] .stButton>button:hover{
    background:#0B2438!important;color:#CBD5E1!important;
    transform:none!important;box-shadow:none!important}

/* ========== PAGE CONTENT TABS ========== */
.stTabs [data-baseweb="tab-list"]{
    background:#FFFFFF;gap:0;padding:0 12px;
    border-bottom:2px solid #E5E7EB;
    box-shadow:0 1px 3px rgba(7,24,39,0.05)}
.stTabs [data-baseweb="tab"]{
    font-size:12px;font-weight:600;color:#6B7280;
    padding:10px 16px;border-bottom:2px solid transparent;
    margin-bottom:-2px;background:transparent;border-radius:0;
    letter-spacing:0.3px;text-transform:uppercase}
.stTabs [aria-selected="true"]{
    color:#071827!important;border-bottom:2px solid #F5A623!important;
    background:transparent!important}
.stTabs [data-baseweb="tab"]:hover{color:#071827!important;background:#F8FAFC!important}
.stTabs [data-baseweb="tab-panel"]{padding:20px 0 0 0}

/* ========== METRIC CONTAINERS ========== */
div[data-testid="metric-container"]{
    background:#FFFFFF;border-radius:8px;padding:14px 16px;
    border:1px solid #E5E7EB;box-shadow:0 1px 3px rgba(7,24,39,0.06)}

/* ========== DATAFRAMES ========== */
.stDataFrame{border-radius:6px;border:1px solid #E5E7EB;overflow:hidden;
             box-shadow:0 1px 4px rgba(7,24,39,0.06)}
.stDataFrame thead tr th{
    background:#071827!important;color:#F3F5F7!important;
    font-size:10px!important;font-weight:700!important;
    padding:10px 14px!important;letter-spacing:0.8px!important;text-transform:uppercase!important}
.stDataFrame tbody tr:nth-child(even){background:#F8FAFC}
.stDataFrame tbody tr:hover{background:#EFF6FF!important}
.stDataFrame tbody tr td{font-size:13px!important;padding:8px 14px!important;color:#111827!important}

/* ========== MAIN AREA BUTTONS ========== */
[data-testid="stAppViewContainer"] .stButton>button{
    background:#071827;color:white;border:none;border-radius:6px;
    font-weight:600;font-size:13px;padding:8px 18px;
    transition:all 0.15s ease;letter-spacing:0.2px}
[data-testid="stAppViewContainer"] .stButton>button:hover{
    background:#0B2438;box-shadow:0 4px 12px rgba(7,24,39,0.25);transform:translateY(-1px)}
[data-testid="stAppViewContainer"] .stButton>button[kind="primary"]{
    background:#F5A623;color:#071827}
[data-testid="stAppViewContainer"] .stButton>button[kind="primary"]:hover{
    background:#e09520;box-shadow:0 4px 12px rgba(245,166,35,0.35)}

/* ========== DOWNLOAD BUTTONS ========== */
.stDownloadButton>button{
    background:#FFFFFF!important;color:#071827!important;
    border:1.5px solid #071827!important;border-radius:6px!important;
    font-weight:600!important;font-size:12px!important;padding:7px 16px!important;
    transition:all 0.15s ease!important;letter-spacing:0.2px!important}
.stDownloadButton>button:hover{
    background:#071827!important;color:white!important;transform:translateY(-1px)!important}

/* ========== INPUTS ========== */
.stSelectbox>div>div,.stMultiSelect>div>div{
    border-radius:6px;border:1.5px solid #E5E7EB;font-size:13px;background:#FFFFFF}
.stTextInput>div>div>input{
    border-radius:6px;border:1.5px solid #E5E7EB;
    font-size:13px;padding:8px 12px;background:#FFFFFF;color:#111827}
.stTextInput>div>div>input:focus{
    border-color:#F5A623;box-shadow:0 0 0 3px rgba(245,166,35,0.15);outline:none}
.stNumberInput>div>div>input{
    border-radius:6px;border:1.5px solid #E5E7EB;font-size:13px;padding:8px 12px}

/* ========== EXPANDER ========== */
.streamlit-expanderHeader{
    background:#FFFFFF;border:1px solid #E5E7EB;border-radius:6px;
    font-weight:600;font-size:13px;color:#111827;padding:10px 14px}
.streamlit-expanderContent{
    border:1px solid #E5E7EB;border-top:none;border-radius:0 0 6px 6px;background:#FFFFFF}

/* ========== ALERTS ========== */
div[data-testid="stAlert"]{border-radius:6px;font-size:13px}
.stAlert{border-radius:6px}

/* ========== HEADINGS ========== */
h1{display:none!important}
h2{font-size:18px!important;font-weight:700!important;color:#071827!important;
   margin:20px 0 10px 0!important;letter-spacing:-0.3px}
h3{font-size:14px!important;font-weight:600!important;color:#0B2438!important;
   margin:14px 0 8px 0!important}
h4,h5{font-size:13px!important;font-weight:600!important;color:#374151!important}
p,li{font-size:13px;color:#374151;line-height:1.6}
.stCaption,.stCaption p{font-size:12px!important;color:#9CA3AF!important}

/* ========== DIVIDER ========== */
hr{border:none;border-top:1px solid #E5E7EB;margin:16px 0}

/* ========== STATUS CHIPS ========== */
.chip{display:inline-flex;align-items:center;gap:4px;
      padding:2px 9px;border-radius:3px;font-size:10px;font-weight:700;letter-spacing:0.5px;
      text-transform:uppercase}
.chip-red   {background:#FEE2E2;color:#991B1B}
.chip-amber {background:#FEF3C7;color:#92400E}
.chip-green {background:#DCFCE7;color:#166534}
.chip-blue  {background:#DBEAFE;color:#1E40AF}
.chip-grey  {background:#F3F4F6;color:#4B5563}
.chip-teal  {background:#CCFBF1;color:#0F766E}

/* ========== KPI CARDS ========== */
.kpi-card{background:#FFFFFF;border-radius:8px;padding:16px 18px;
          border:1px solid #E5E7EB;box-shadow:0 1px 4px rgba(7,24,39,0.06);height:100%}
.kpi-label{font-size:10px;font-weight:700;color:#6B7280;
           letter-spacing:1.2px;text-transform:uppercase;margin-bottom:8px}
.kpi-value{font-size:30px;font-weight:800;color:#071827;line-height:1;margin-bottom:4px;
           font-variant-numeric:tabular-nums}
.kpi-sub  {font-size:11px;color:#6B7280;margin-top:4px}

/* ========== SECTION LABELS ========== */
.section-label{font-size:10px;font-weight:700;color:#6B7280;
               letter-spacing:1.5px;text-transform:uppercase;margin-bottom:10px}

/* ========== PAGE HEADER / CONTROL BAR ========== */
.pt-page-header{
    background:linear-gradient(180deg,#071827 0%,#0B2438 100%);
    border-radius:0;padding:16px 28px;margin-bottom:20px;
    margin-left:-28px;margin-right:-28px;margin-top:0;
    border-bottom:2px solid #F5A623}

/* ========== ATTENTION PANEL ========== */
.attention-item{background:#FFFFFF;border-radius:6px;padding:11px 14px;
                border-left:3px solid #DC2626;margin-bottom:6px;
                box-shadow:0 1px 3px rgba(7,24,39,0.05)}
.attention-item-amber{border-left-color:#F59E0B}
.attention-item-blue {border-left-color:#3B82F6}

/* ========== EMPTY STATE ========== */
.empty-state{background:#FFFFFF;border-radius:8px;padding:40px 32px;
             text-align:center;border:1px dashed #D1D5DB}

/* ========== DATA QUALITY CARD ========== */
.dq-row{display:flex;align-items:center;justify-content:space-between;
        padding:7px 0;border-bottom:1px solid #F3F4F6;font-size:12px}
.dq-row:last-child{border-bottom:none}

/* ========== MODULE CARDS (landing page) ========== */
.module-card{background:#FFFFFF;border-radius:8px;padding:20px;
             border:1px solid #E5E7EB;box-shadow:0 1px 4px rgba(7,24,39,0.06);
             border-top:3px solid #F5A623;height:100%}
.module-label{font-size:9px;font-weight:800;color:#F5A623;
              letter-spacing:2px;text-transform:uppercase;margin-bottom:6px}
.module-title{font-size:15px;font-weight:700;color:#071827;margin-bottom:6px;line-height:1.2}
.module-desc{font-size:12px;color:#6B7280;line-height:1.55;margin-bottom:12px}
.module-outputs{font-size:11px;color:#374151;padding-left:0;list-style:none}
.module-outputs li{padding:2px 0;display:flex;align-items:center;gap:5px}
.module-outputs li::before{content:"";display:inline-block;width:4px;height:4px;
                            border-radius:50%;background:#F5A623;flex-shrink:0}

/* ========== STATUS STRIP ========== */
.status-strip{background:#0B2438;padding:8px 28px;margin-left:-28px;margin-right:-28px;
              display:flex;gap:0;border-bottom:1px solid #102A43}
.status-item{display:flex;align-items:center;gap:8px;padding:0 20px 0 0;
             border-right:1px solid #102A43;margin-right:20px}
.status-item:last-child{border-right:none;margin-right:0}
.status-key{font-size:10px;color:#4B6275;text-transform:uppercase;letter-spacing:0.8px;font-weight:600}
.status-val{font-size:11px;color:#8FA3B1;font-weight:600;margin-left:4px}
.status-dot-ok{width:5px;height:5px;border-radius:50%;background:#16A34A;display:inline-block}
.status-dot-warn{width:5px;height:5px;border-radius:50%;background:#374151;display:inline-block}
</style>
""", unsafe_allow_html=True)






# -----------------------------------------------------------------------------
# XER PARSING  (xerparser + manual fallback)
# -----------------------------------------------------------------------------

def parse_xer_fallback(raw_text: str) -> dict:
    """
    Manual fallback parser that reads XER table format:
    %T TABLE_NAME  /  %F col1 col2 ...  /  %R val1 val2 ...
    Returns dict of {table_name: list_of_dicts}
    """
    tables = {}
    current_table = None
    current_fields = []

    for line in raw_text.splitlines():
        line = line.rstrip("\r")
        if line.startswith("%T\t"):
            current_table = line[3:].strip()
            current_fields = []
            tables[current_table] = []
        elif line.startswith("%F\t") and current_table:
            current_fields = line[3:].split("\t")
        elif line.startswith("%R\t") and current_table and current_fields:
            values = line[3:].split("\t")
            # Pad values if shorter than fields
            while len(values) < len(current_fields):
                values.append("")
            row = {current_fields[i]: values[i] for i in range(len(current_fields))}
            tables[current_table].append(row)

    return tables


def hours_to_days(hours, hours_per_day=8.0):
    """Convert hours to working days."""
    if hours is None:
        return None
    try:
        return round(float(hours) / hours_per_day, 1)
    except (TypeError, ValueError):
        return None


def safe_float(val, default=None):
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def safe_date(val):
    if val is None or str(val).strip() in ("", "None"):
        return None
    if isinstance(val, datetime):
        return val
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt)
        except ValueError:
            pass
    return None


def parse_xer(file_bytes: bytes):
    """
    Parse an XER file. Uses xerparser library first; falls back to manual parsing.
    Returns a dict with keys: tasks_df, relationships_df, wbs_df, resources_df,
    task_resources_df, project_info, calendars_df, parse_method
    """
    # Try to decode the file
    for codec in ("cp1252", "utf-8", "latin-1"):
        try:
            raw_text = file_bytes.decode(codec)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("Cannot decode XER file. Please check the file encoding.")

    result = {
        "tasks_df": pd.DataFrame(),
        "relationships_df": pd.DataFrame(),
        "wbs_df": pd.DataFrame(),
        "resources_df": pd.DataFrame(),
        "task_resources_df": pd.DataFrame(),
        "project_info": {},
        "calendars_df": pd.DataFrame(),
        "parse_method": "unknown",
    }

    # -- Try xerparser library -------------------------------------------------
    try:
        from xerparser.src.xer import Xer
        xer = Xer(raw_text)

        # Project info
        proj = None
        if xer.projects:
            proj_id = next(iter(xer.projects))
            proj = xer.projects[proj_id]
            result["project_info"] = {
                "name": getattr(proj, "name", ""),
                "data_date": getattr(proj, "last_recalc_date", None),
                "project_id": proj_id,
                "plan_start": getattr(proj, "plan_start_date", None),
                "scd_end": getattr(proj, "scd_end_date", None),
            }

        # Tasks DataFrame
        rows = []
        for uid, task in xer.tasks.items():
            tf = task.total_float_hr_cnt
            ff = task.free_float_hr_cnt
            # Effective start/finish (actual if done, early if not)
            eff_start = task.act_start_date or task.early_start_date or task.target_start_date
            eff_finish = task.act_end_date or task.early_end_date or task.target_end_date

            # WBS path
            wbs_node = xer.wbs_nodes.get(task.wbs_id)
            wbs_path = ""
            if wbs_node:
                parts = []
                n = wbs_node
                while n:
                    parts.append(getattr(n, "name", ""))
                    n = getattr(n, "parent", None)
                wbs_path = " > ".join(reversed(parts))

            # Calendar name
            cal = xer.calendars.get(task.clndr_id)
            cal_name = getattr(cal, "name", "") if cal else ""

            rows.append({
                "task_id": uid,
                "task_code": task.task_code,
                "task_name": task.name,
                "wbs_id": task.wbs_id,
                "wbs_path": wbs_path,
                "status": task.status.value if task.status else "",
                "task_type": task.type.value if task.type else "",
                "calendar": cal_name,
                "early_start": task.early_start_date,
                "early_finish": task.early_end_date,
                "late_start": task.late_start_date,
                "late_finish": task.late_end_date,
                "act_start": task.act_start_date,
                "act_finish": task.act_end_date,
                "target_start": task.target_start_date,
                "target_finish": task.target_end_date,
                "eff_start": eff_start,
                "eff_finish": eff_finish,
                "orig_dur_days": hours_to_days(task.target_drtn_hr_cnt),
                "rem_dur_days": hours_to_days(task.remain_drtn_hr_cnt),
                "total_float_days": hours_to_days(tf),
                "free_float_days": hours_to_days(ff),
                "total_float_hrs": tf,
                "is_longest_path": task.is_longest_path,
                "cstr_type": task.cstr_type,
                "cstr_date": task.cstr_date,
                "cstr_type2": task.cstr_type2,
                "cstr_date2": task.cstr_date2,
                "phys_pct": round(task.phys_complete_pct * 100, 1),
                "float_path": task.float_path,
            })

        result["tasks_df"] = pd.DataFrame(rows)

        # Relationships DataFrame
        rel_rows = []
        for uid, rel in xer.relationships.items():
            rel_rows.append({
                "pred_id": uid,
                "pred_task_id": rel.predecessor.uid if rel.predecessor else "",
                "pred_task_code": rel.predecessor.task_code if rel.predecessor else "",
                "pred_task_name": rel.predecessor.name if rel.predecessor else "",
                "succ_task_id": rel.successor.uid if rel.successor else "",
                "succ_task_code": rel.successor.task_code if rel.successor else "",
                "succ_task_name": rel.successor.name if rel.successor else "",
                "rel_type": rel.link,
                "lag_days": rel.lag,
                "lag_hrs": rel.lag_hr_cnt,
            })
        result["relationships_df"] = pd.DataFrame(rel_rows)

        # WBS DataFrame
        wbs_rows = []
        for uid, wbs in xer.wbs_nodes.items():
            wbs_rows.append({
                "wbs_id": uid,
                "wbs_code": getattr(wbs, "short_name", ""),
                "wbs_name": getattr(wbs, "name", ""),
                "parent_wbs_id": getattr(wbs, "parent_wbs_id", ""),
                "proj_id": getattr(wbs, "proj_id", ""),
            })
        result["wbs_df"] = pd.DataFrame(wbs_rows)

        # Resources & task resources
        if xer.resources:
            res_rows = []
            for uid, r in xer.resources.items():
                res_rows.append({
                    "rsrc_id": uid,
                    "rsrc_name": getattr(r, "name", ""),
                    "rsrc_short": getattr(r, "rsrc_short_name", ""),
                    "rsrc_type": getattr(r, "rsrc_type", ""),
                })
            result["resources_df"] = pd.DataFrame(res_rows)

        # Task resources (loading)
        taskrsrc_rows = []
        for uid, task in xer.tasks.items():
            for tr in getattr(task, "resources", []):
                taskrsrc_rows.append({
                    "task_id": uid,
                    "task_code": task.task_code,
                    "rsrc_id": getattr(tr, "rsrc_id", ""),
                    "target_qty": safe_float(getattr(tr, "target_qty", 0), 0),
                    "remain_qty": safe_float(getattr(tr, "remain_qty", 0), 0),
                    "act_reg_qty": safe_float(getattr(tr, "act_reg_qty", 0), 0),
                    "target_start": safe_date(getattr(tr, "target_start_date", None)),
                    "target_finish": safe_date(getattr(tr, "target_end_date", None)),
                })
        result["task_resources_df"] = pd.DataFrame(taskrsrc_rows)

        result["parse_method"] = "xerparser"
        return result

    except Exception as e:
        st.warning(f"xerparser failed ({e}), using fallback parser...")

    # -- Manual fallback -------------------------------------------------------
    try:
        tables = parse_xer_fallback(raw_text)
        return _build_from_raw_tables(tables)
    except Exception as e2:
        raise ValueError(f"Both parsers failed. Last error: {e2}")


def _build_from_raw_tables(tables: dict) -> dict:
    """Build result dict from raw parsed tables (fallback)."""
    result = {
        "tasks_df": pd.DataFrame(),
        "relationships_df": pd.DataFrame(),
        "wbs_df": pd.DataFrame(),
        "resources_df": pd.DataFrame(),
        "task_resources_df": pd.DataFrame(),
        "project_info": {},
        "calendars_df": pd.DataFrame(),
        "parse_method": "manual_fallback",
    }

    # Project info
    if "PROJECT" in tables and tables["PROJECT"]:
        proj = tables["PROJECT"][0]
        result["project_info"] = {
            "name": proj.get("proj_short_name", proj.get("proj_id", "")),
            "data_date": safe_date(proj.get("last_recalc_date")),
            "plan_start": safe_date(proj.get("plan_start_date")),
            "scd_end": safe_date(proj.get("scd_end_date")),
        }

    # Tasks
    if "TASK" in tables:
        df = pd.DataFrame(tables["TASK"])
        # Normalise date columns
        for col in ["early_start_date", "early_end_date", "late_start_date",
                    "late_end_date", "act_start_date", "act_end_date",
                    "target_start_date", "target_end_date", "cstr_date", "cstr_date2"]:
            if col in df.columns:
                df[col] = df[col].apply(safe_date)
        # Float
        for col in ["total_float_hr_cnt", "free_float_hr_cnt",
                    "target_drtn_hr_cnt", "remain_drtn_hr_cnt"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # Build normalised columns
        df["eff_start"] = df.get("act_start_date", df.get("early_start_date"))
        df["eff_finish"] = df.get("act_end_date", df.get("early_end_date"))
        df["total_float_days"] = df.get("total_float_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)
        df["free_float_days"] = df.get("free_float_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)
        df["orig_dur_days"] = df.get("target_drtn_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)
        df["rem_dur_days"] = df.get("remain_drtn_hr_cnt", pd.Series(dtype=float)).apply(hours_to_days)

        # Rename for consistency
        rename = {
            "task_id": "task_id", "task_code": "task_code",
            "task_name": "task_name", "wbs_id": "wbs_id",
            "status_code": "status", "task_type": "task_type",
            "early_start_date": "early_start", "early_end_date": "early_finish",
            "late_start_date": "late_start", "late_end_date": "late_finish",
            "act_start_date": "act_start", "act_end_date": "act_finish",
            "target_start_date": "target_start", "target_end_date": "target_finish",
            "cstr_type": "cstr_type", "cstr_date": "cstr_date",
            "cstr_type2": "cstr_type2", "cstr_date2": "cstr_date2",
            "driving_path_flag": "is_longest_path_flag",
            "phys_complete_pct": "phys_pct",
        }
        df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
        if "is_longest_path_flag" in df.columns:
            df["is_longest_path"] = df["is_longest_path_flag"] == "Y"
        df["wbs_path"] = df.get("wbs_id", "")
        result["tasks_df"] = df

    # Relationships
    if "TASKPRED" in tables:
        df = pd.DataFrame(tables["TASKPRED"])
        if "lag_hr_cnt" in df.columns:
            df["lag_days"] = pd.to_numeric(df["lag_hr_cnt"], errors="coerce").apply(hours_to_days)
        rename_r = {"pred_type": "rel_type", "task_id": "succ_task_id",
                    "pred_task_id": "pred_task_id"}
        df = df.rename(columns={k: v for k, v in rename_r.items() if k in df.columns})
        result["relationships_df"] = df

    # WBS
    if "PROJWBS" in tables:
        df = pd.DataFrame(tables["PROJWBS"])
        rename_w = {"wbs_id": "wbs_id", "wbs_short_name": "wbs_code",
                    "wbs_name": "wbs_name", "parent_wbs_id": "parent_wbs_id"}
        df = df.rename(columns={k: v for k, v in rename_w.items() if k in df.columns})
        result["wbs_df"] = df

    # Resources
    if "RSRC" in tables:
        df = pd.DataFrame(tables["RSRC"])
        rename_rs = {"rsrc_id": "rsrc_id", "rsrc_name": "rsrc_name",
                     "rsrc_short_name": "rsrc_short"}
        df = df.rename(columns={k: v for k, v in rename_rs.items() if k in df.columns})
        result["resources_df"] = df

    # Task resources
    if "TASKRSRC" in tables:
        df = pd.DataFrame(tables["TASKRSRC"])
        for col in ["target_qty", "remain_qty", "act_reg_qty"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in ["target_start_date", "target_end_date"]:
            if col in df.columns:
                df[col] = df[col].apply(safe_date)
                df = df.rename(columns={col: col.replace("_date", "")})
        result["task_resources_df"] = df

    return result


# -----------------------------------------------------------------------------
# GRAPH BUILDING
# -----------------------------------------------------------------------------

def build_graph(tasks_df: pd.DataFrame, rels_df: pd.DataFrame) -> nx.DiGraph:
    """Build a networkx directed graph from tasks and relationships."""
    G = nx.DiGraph()
    for _, row in tasks_df.iterrows():
        G.add_node(row["task_id"], **row.to_dict())
    for _, row in rels_df.iterrows():
        if row.get("pred_task_id") and row.get("succ_task_id"):
            G.add_edge(
                row["pred_task_id"],
                row["succ_task_id"],
                rel_type=row.get("rel_type", "FS"),
                lag_days=row.get("lag_days", 0),
            )
    return G


# -----------------------------------------------------------------------------
# CRITICAL PATH HELPERS
# -----------------------------------------------------------------------------

def get_critical_threshold(tasks_df: pd.DataFrame, near_crit_days: float = 10.0):
    """Classify activities as critical / near-critical / float."""
    df = tasks_df.copy()
    df["is_critical"] = df["total_float_days"].apply(
        lambda f: f is not None and f <= 0
    )
    df["is_near_critical"] = df["total_float_days"].apply(
        lambda f: f is not None and 0 < f <= near_crit_days
    )
    return df


def float_status_badge(f):
    if f is None:
        return "-"
    elif f <= 0:
        return "🔴 Critical"
    elif f <= 10:
        return "🟡 Near-Critical"
    else:
        return "🟢 Float"


# -----------------------------------------------------------------------------
# LOGIC TRACE HELPERS
# -----------------------------------------------------------------------------

def trace_predecessors(G: nx.DiGraph, task_id: str, max_depth=100) -> list:
    """BFS backwards through predecessors. Returns list of (task_id, depth)."""
    visited = {}
    queue = deque([(task_id, 0)])
    result = []
    while queue:
        node, depth = queue.popleft()
        if node in visited or depth > max_depth:
            continue
        visited[node] = depth
        if node != task_id:
            result.append((node, depth))
        for pred in G.predecessors(node):
            if pred not in visited:
                queue.append((pred, depth + 1))
    return result


def trace_successors(G: nx.DiGraph, task_id: str, max_depth=100) -> list:
    """BFS forwards through successors."""
    visited = {}
    queue = deque([(task_id, 0)])
    result = []
    while queue:
        node, depth = queue.popleft()
        if node in visited or depth > max_depth:
            continue
        visited[node] = depth
        if node != task_id:
            result.append((node, depth))
        for succ in G.successors(node):
            if succ not in visited:
                queue.append((succ, depth + 1))
    return result


def driving_path_to_activity(
    G: nx.DiGraph,
    tasks_df: pd.DataFrame,
    rels_df: pd.DataFrame,
    target_id: str,
) -> list:
    """
    Identify the most likely driving predecessor chain into a target activity.

    Driving predecessor selection priority (in order):
      1. Lowest total float  (most constrained activity wins)
      2. On P6 longest-path / driving flag where available
      3. Latest early-finish date  (latest predecessor is usually the driver)
      4. Highest lag on the connecting relationship (more constraining)

    Returns ordered list of task_ids, from chain start -> target.
    """
    task_lookup = tasks_df.set_index("task_id").to_dict("index") if not tasks_df.empty else {}

    def _score(pred_id, succ_id):
        t  = task_lookup.get(pred_id, {})
        tf = safe_float(t.get("total_float_days"), 9999)
        finish = t.get("eff_finish")
        if finish is not None:
            try:
                finish_score = -(finish.timestamp() / 86400)
            except Exception:
                finish_score = 0
        else:
            finish_score = 0
        driving_bonus = 0 if t.get("is_longest_path", False) else 1
        lag_score = 0
        if not rels_df.empty:
            rel = rels_df[
                (rels_df.get("pred_task_id", pd.Series(dtype=str)) == pred_id) &
                (rels_df.get("succ_task_id", pd.Series(dtype=str)) == succ_id)
            ]
            if not rel.empty and "lag_days" in rel.columns:
                lag_val = safe_float(rel["lag_days"].iloc[0], 0)
                lag_score = -lag_val
        return (tf, driving_bonus, finish_score, lag_score)

    path    = [target_id]
    visited = {target_id}
    current = target_id

    for _ in range(500):
        preds = list(G.predecessors(current))
        if not preds:
            break
        unvisited = [p for p in preds if p not in visited]
        if not unvisited:
            break
        best = min(unvisited, key=lambda p: _score(p, current))
        path.insert(0, best)
        visited.add(best)
        current = best

    return path


def _all_pred_paths(
    G: nx.DiGraph,
    tasks_df: pd.DataFrame,
    target_id: str,
    max_paths: int = 8,
) -> list:
    """
    Find up to max_paths predecessor chains into target_id.
    Each chain is a list of task_ids ordered start -> target.
    Only returns chains that begin at an activity with no predecessors.
    """
    task_lookup = tasks_df.set_index("task_id").to_dict("index") if not tasks_df.empty else {}

    def _float(tid):
        return safe_float(task_lookup.get(tid, {}).get("total_float_days"), 9999)

    found_paths = []

    def dfs(node, current_path, visited_set):
        if len(found_paths) >= max_paths:
            return
        preds = [p for p in G.predecessors(node) if p not in visited_set]
        if not preds:
            found_paths.append(list(reversed(current_path)))
            return
        for pred in sorted(preds, key=_float)[:4]:
            dfs(pred, current_path + [pred], visited_set | {pred})

    dfs(target_id, [target_id], {target_id})
    return found_paths


# -----------------------------------------------------------------------------
# EXPORT HELPERS
# -----------------------------------------------------------------------------

def style_header_row(ws, row_idx, fill_color="1e3a5f", font_color="FFFFFF"):
    fill = PatternFill("solid", start_color=fill_color, fgColor=fill_color)
    font = Font(bold=True, color=font_color)
    for cell in ws[row_idx]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def df_to_sheet(ws, df, sheet_title=None):
    """Write a DataFrame to an openpyxl worksheet with formatting."""
    if sheet_title:
        ws.title = sheet_title[:31]
    ws.append(list(df.columns))
    style_header_row(ws, 1)
    for r in df.itertuples(index=False):
        ws.append(list(r))
    # Auto-width
    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 50)


def export_df_to_excel(sheets: dict) -> bytes:
    """sheets = {sheet_name: dataframe}. Returns Excel bytes."""
    wb = Workbook()
    first = True
    for name, df in sheets.items():
        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()
        df_to_sheet(ws, df, name)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def format_date(d):
    if d is None:
        return "-"
    try:
        return d.strftime("%d %b %Y")
    except Exception:
        return str(d)


# -----------------------------------------------------------------------------
# PAGE: PROJECT SUMMARY
# -----------------------------------------------------------------------------

def page_project_summary(data: dict, near_crit_days: float):
    st.title("📊 Project Summary")

    proj = data["project_info"]
    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities found in this file.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # Header metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Project Name", proj.get("name", "Unknown"))
    c2.metric("Data Date", format_date(proj.get("data_date")))
    c3.metric("Parse Method", data.get("parse_method", "-"))
    c4.metric("Activities", len(tasks))

    st.divider()
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("🔴 Critical", int(tasks["is_critical"].sum()))
    c2.metric("🟡 Near-Critical", int(tasks["is_near_critical"].sum()))
    neg_float = tasks["total_float_days"].apply(lambda f: f is not None and f < 0).sum()
    c3.metric("⚠️ Negative Float", int(neg_float))
    c4.metric("🔗 Relationships", len(rels))

    # Open-ended
    if not rels.empty and "pred_task_id" in rels.columns:
        tasks_with_pred = set(rels["succ_task_id"].dropna())
        tasks_with_succ = set(rels["pred_task_id"].dropna())
        task_ids = set(tasks["task_id"])
        no_pred = len(task_ids - tasks_with_pred)
        no_succ = len(task_ids - tasks_with_succ)
        c5.metric("Open-Ended Activities", no_pred + no_succ)
    else:
        c5.metric("Open-Ended Activities", "-")

    # Date range
    valid_starts = tasks["eff_start"].dropna()
    valid_finishes = tasks["eff_finish"].dropna()
    if not valid_starts.empty and not valid_finishes.empty:
        earliest = min(valid_starts)
        latest = max(valid_finishes)
        st.info(f"**Schedule Span:** {format_date(earliest)} -> {format_date(latest)}")

    # Constraint count
    constrained = tasks["cstr_type"].apply(lambda x: bool(x) and str(x).strip() not in ("", "None")).sum() if "cstr_type" in tasks.columns else 0
    st.info(f"**Constrained Activities:** {int(constrained)}")

    st.divider()

    # Charts
    tab1, tab2, tab3 = st.tabs(["Float Distribution", "Activities by WBS", "Status Breakdown"])

    with tab1:
        float_vals = tasks["total_float_days"].dropna()
        if not float_vals.empty:
            fig = px.histogram(
                float_vals, nbins=40, title="Total Float Distribution (Days)",
                labels={"value": "Float (days)", "count": "Activities"},
                color_discrete_sequence=["#2563eb"],
            )
            fig.add_vline(x=0, line_dash="dash", line_color="red", annotation_text="Critical")
            fig.add_vline(x=near_crit_days, line_dash="dot", line_color="orange",
                          annotation_text=f"Near-Critical ({near_crit_days}d)")
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        if "wbs_path" in tasks.columns:
            # Show top-level WBS only
            tasks["wbs_top"] = tasks["wbs_path"].apply(
                lambda x: str(x).split(" > ")[0] if pd.notna(x) and x else "Unknown"
            )
            wbs_counts = tasks.groupby("wbs_top").size().reset_index(name="count")
            wbs_counts = wbs_counts.sort_values("count", ascending=False).head(20)
            fig = px.bar(wbs_counts, x="count", y="wbs_top", orientation="h",
                         title="Activities by Top-Level WBS",
                         color_discrete_sequence=["#1e3a5f"])
            fig.update_layout(yaxis_title="", xaxis_title="Activity Count")
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        if "status" in tasks.columns:
            status_counts = tasks["status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            fig = px.pie(status_counts, values="Count", names="Status",
                         title="Activity Status Breakdown",
                         color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig, use_container_width=True)

    # Summary table
    st.subheader("Activity Summary Table")
    display_cols = ["task_code", "task_name", "wbs_path", "eff_start", "eff_finish",
                    "total_float_days", "status", "is_critical"]
    avail = [c for c in display_cols if c in tasks.columns]
    st.dataframe(tasks[avail].head(100), use_container_width=True)


# -----------------------------------------------------------------------------
# PAGE: ACTIVITY SEARCH
# -----------------------------------------------------------------------------

def _col(df: pd.DataFrame, col: str, default="-"):
    """Safely get a column value; return default if column missing or null."""
    if col not in df.index:
        return default
    val = df.get(col, default)
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return default
    return val


def _float_color(f) -> str:
    """Return a hex colour string for a float value."""
    if f is None:
        return "#6b7280"
    try:
        f = float(f)
    except (TypeError, ValueError):
        return "#6b7280"
    if f < 0:
        return "#991b1b"   # dark red  - negative float
    if f == 0:
        return "#dc2626"   # red       - critical
    if f <= 10:
        return "#d97706"   # amber     - near-critical
    return "#15803d"       # green     - has float


def _status_label(status: str) -> str:
    """Convert P6 status code to a readable label."""
    mapping = {
        "TK_NotStart": "Not Started",
        "TK_Active":   "In Progress",
        "TK_Complete": "Complete",
        "Not Started": "Not Started",
        "In Progress": "In Progress",
        "Complete":    "Complete",
    }
    return mapping.get(str(status).strip(), str(status).strip() or "-")


def _status_colour(status: str) -> str:
    s = _status_label(status)
    if s == "Complete":    return "#15803d"
    if s == "In Progress": return "#2563eb"
    return "#6b7280"


def page_activity_search(data: dict, near_crit_days: float):
    """
    Activity Search page.
    Lets the user filter activities by ID, name, WBS, date range and
    critical status, then shows a full detail panel for the selected activity.
    """
    st.title("🔍 Activity Search")
    st.caption("Search and filter activities, then click any row to view its full detail.")

    tasks = data["tasks_df"]
    rels  = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities found. Please upload an XER file first.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # -------------------------------------------------------------------------
    # SEARCH / FILTER PANEL
    # -------------------------------------------------------------------------
    with st.expander("🔎  Search & Filter", expanded=True):
        r1c1, r1c2, r1c3 = st.columns(3)
        search_code = r1c1.text_input("Activity ID", placeholder="e.g. A1000")
        search_name = r1c2.text_input("Activity Name", placeholder="partial match")
        search_wbs  = r1c3.text_input("WBS", placeholder="partial match")

        r2c1, r2c2, r2c3 = st.columns(3)

        # Critical / float filter
        crit_filter = r2c1.selectbox(
            "Float Status",
            ["All", "Critical (float <= 0)", "Near-Critical", "Positive Float", "Negative Float"],
        )

        # Status filter - only show if column exists
        if "status" in tasks.columns:
            status_opts = ["All"] + sorted(
                [s for s in tasks["status"].dropna().unique() if str(s).strip()]
            )
        else:
            status_opts = ["All"]
        status_filter = r2c2.selectbox("Activity Status", status_opts)

        # Activity type filter
        if "task_type" in tasks.columns:
            type_opts = ["All"] + sorted(
                [t for t in tasks["task_type"].dropna().unique() if str(t).strip()]
            )
        else:
            type_opts = ["All"]
        type_filter = r2c3.selectbox("Activity Type", type_opts)

        # Date range - only show if date columns exist and have data
        date_from = date_to = None
        valid_starts  = tasks["eff_start"].dropna()  if "eff_start"  in tasks.columns else pd.Series(dtype=object)
        valid_finishes = tasks["eff_finish"].dropna() if "eff_finish" in tasks.columns else pd.Series(dtype=object)

        if not valid_starts.empty and not valid_finishes.empty:
            try:
                min_d = min(valid_starts).date()
                max_d = max(valid_finishes).date()
                dc1, dc2 = st.columns(2)
                date_from = dc1.date_input("Finish on or After", value=min_d,
                                           min_value=min_d, max_value=max_d)
                date_to   = dc2.date_input("Finish on or Before", value=max_d,
                                           min_value=min_d, max_value=max_d)
            except Exception:
                pass  # silently skip date filter if dates are unusable

    # -------------------------------------------------------------------------
    # APPLY FILTERS
    # -------------------------------------------------------------------------
    filtered = tasks.copy()

    if search_code.strip():
        if "task_code" in filtered.columns:
            filtered = filtered[
                filtered["task_code"].astype(str).str.contains(
                    search_code.strip(), case=False, na=False
                )
            ]
    if search_name.strip():
        if "task_name" in filtered.columns:
            filtered = filtered[
                filtered["task_name"].astype(str).str.contains(
                    search_name.strip(), case=False, na=False
                )
            ]
    if search_wbs.strip():
        if "wbs_path" in filtered.columns:
            filtered = filtered[
                filtered["wbs_path"].astype(str).str.contains(
                    search_wbs.strip(), case=False, na=False
                )
            ]

    if crit_filter == "Critical (float <= 0)" and "is_critical" in filtered.columns:
        filtered = filtered[filtered["is_critical"] == True]
    elif crit_filter == "Near-Critical" and "is_near_critical" in filtered.columns:
        filtered = filtered[filtered["is_near_critical"] == True]
    elif crit_filter == "Positive Float" and "total_float_days" in filtered.columns:
        filtered = filtered[filtered["total_float_days"].apply(
            lambda f: f is not None and safe_float(f, 1) > 0
        )]
    elif crit_filter == "Negative Float" and "total_float_days" in filtered.columns:
        filtered = filtered[filtered["total_float_days"].apply(
            lambda f: f is not None and safe_float(f, 0) < 0
        )]

    if status_filter != "All" and "status" in filtered.columns:
        filtered = filtered[filtered["status"] == status_filter]

    if type_filter != "All" and "task_type" in filtered.columns:
        filtered = filtered[filtered["task_type"] == type_filter]

    if date_from is not None and "eff_finish" in filtered.columns:
        filtered = filtered[
            filtered["eff_finish"].apply(
                lambda d: d is not None and hasattr(d, "date") and d.date() >= date_from
            )
        ]
    if date_to is not None and "eff_finish" in filtered.columns:
        filtered = filtered[
            filtered["eff_finish"].apply(
                lambda d: d is not None and hasattr(d, "date") and d.date() <= date_to
            )
        ]

    # -------------------------------------------------------------------------
    # RESULTS TABLE
    # -------------------------------------------------------------------------
    n_found = len(filtered)
    n_total = len(tasks)

    if n_found == 0:
        st.warning("No activities match your filters. Try broadening your search.")
        return

    # Build a clean display version of the table
    TABLE_COLS = {
        "task_code":        "Activity ID",
        "task_name":        "Activity Name",
        "wbs_path":         "WBS",
        "eff_start":        "Start",
        "eff_finish":       "Finish",
        "orig_dur_days":    "Orig Dur (d)",
        "rem_dur_days":     "Rem Dur (d)",
        "total_float_days": "Float (d)",
        "status":           "Status",
        "task_type":        "Type",
        "is_critical":      "Critical",
    }

    present_cols = {k: v for k, v in TABLE_COLS.items() if k in filtered.columns}
    display_df = filtered[list(present_cols.keys())].copy()
    display_df = display_df.rename(columns=present_cols)

    # Format date columns for readability
    for col in ["Start", "Finish"]:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(format_date)

    # Format critical flag
    if "Critical" in display_df.columns:
        display_df["Critical"] = display_df["Critical"].apply(
            lambda x: "Yes" if x else ""
        )

    # Friendly status labels
    if "Status" in display_df.columns:
        display_df["Status"] = display_df["Status"].apply(_status_label)

    st.markdown(
        f"<p style='color:#6b7280;font-size:13px;'>"
        f"Showing <strong>{n_found}</strong> of <strong>{n_total}</strong> activities"
        f"</p>",
        unsafe_allow_html=True,
    )

    st.dataframe(
        display_df,
        use_container_width=True,
        height=min(400, 45 + n_found * 35),
        hide_index=True,
    )

    # -------------------------------------------------------------------------
    # ACTIVITY SELECTOR  -  pick from the filtered results
    # -------------------------------------------------------------------------
    st.divider()

    # Build selector labels
    def make_label(r):
        code = str(r.get("task_code", "?"))
        name = str(r.get("task_name", "?"))
        return f"{code}  --  {name}"

    act_labels = filtered.apply(make_label, axis=1).tolist()

    selected_label = st.selectbox(
        "Select an activity to view full details",
        options=act_labels,
        index=0,
    )

    sel_idx = act_labels.index(selected_label)
    row = filtered.iloc[sel_idx]

    # -------------------------------------------------------------------------
    # SELECTED ACTIVITY BANNER
    # -------------------------------------------------------------------------
    tf_val  = row.get("total_float_days") if "total_float_days" in row.index else None
    tf_num  = safe_float(tf_val, None)
    f_color = _float_color(tf_num)

    status_raw = row.get("status", "") if "status" in row.index else ""
    s_label    = _status_label(str(status_raw))
    s_color    = _status_colour(str(status_raw))

    is_crit = bool(row.get("is_critical", False)) if "is_critical" in row.index else False
    crit_banner = (
        '<span style="background:#dc2626;color:white;padding:3px 10px;'
        'border-radius:12px;font-size:12px;font-weight:700;margin-left:10px;">CRITICAL</span>'
        if is_crit else ""
    )

    task_code = str(row.get("task_code", "-")) if "task_code" in row.index else "-"
    task_name = str(row.get("task_name", "-")) if "task_name" in row.index else "-"

    st.markdown(
        f"""
        <div style="background:#1e3a5f;color:white;border-radius:10px;
                    padding:18px 24px;margin-bottom:16px;">
            <div style="font-size:13px;color:#93c5fd;font-weight:600;
                        letter-spacing:1px;text-transform:uppercase;">
                Selected Activity
            </div>
            <div style="font-size:22px;font-weight:700;margin-top:4px;">
                {task_code}{crit_banner}
            </div>
            <div style="font-size:16px;color:#bfdbfe;margin-top:4px;">
                {task_name}
            </div>
            <div style="margin-top:10px;">
                <span style="background:{s_color};color:white;padding:3px 10px;
                             border-radius:12px;font-size:12px;font-weight:600;">
                    {s_label}
                </span>
                <span style="background:{f_color};color:white;padding:3px 10px;
                             border-radius:12px;font-size:12px;font-weight:600;
                             margin-left:8px;">
                    Float: {tf_num if tf_num is not None else "-"} days
                </span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # DETAIL PANEL  -  two column layout
    # -------------------------------------------------------------------------
    def field(label: str, value, suffix: str = "") -> str:
        """Render a labelled field as HTML. Handles missing/None gracefully."""
        if value is None or str(value).strip() in ("", "None", "nan", "-"):
            disp = '<span style="color:#9ca3af;">Not available</span>'
        else:
            disp = f'<span style="font-weight:600;color:#111827;">{value}{suffix}</span>'
        return (
            f'<div style="padding:6px 0;border-bottom:1px solid #f3f4f6;">'
            f'<span style="color:#6b7280;font-size:12px;text-transform:uppercase;'
            f'letter-spacing:0.5px;">{label}</span><br>{disp}</div>'
        )

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("##### Identity & Classification")
        st.markdown(
            field("Activity ID",    _col(row, "task_code")) +
            field("Activity Name",  _col(row, "task_name")) +
            field("WBS",            _col(row, "wbs_path")) +
            field("Activity Type",  _col(row, "task_type")) +
            field("Calendar",       _col(row, "calendar")) +
            field("% Complete",     _col(row, "phys_pct"), suffix="%"),
            unsafe_allow_html=True,
        )

    with col_right:
        st.markdown("##### Schedule & Float")
        early_start  = format_date(row.get("early_start")  if "early_start"  in row.index else None)
        early_finish = format_date(row.get("early_finish") if "early_finish" in row.index else None)
        late_start   = format_date(row.get("late_start")   if "late_start"   in row.index else None)
        late_finish  = format_date(row.get("late_finish")  if "late_finish"  in row.index else None)
        act_start    = format_date(row.get("act_start")    if "act_start"    in row.index else None)
        act_finish   = format_date(row.get("act_finish")   if "act_finish"   in row.index else None)
        orig_dur     = _col(row, "orig_dur_days")
        rem_dur      = _col(row, "rem_dur_days")
        ff_val       = _col(row, "free_float_days")

        st.markdown(
            field("Early Start",         early_start) +
            field("Early Finish",        early_finish) +
            field("Late Start",          late_start) +
            field("Late Finish",         late_finish) +
            field("Actual Start",        act_start) +
            field("Actual Finish",       act_finish) +
            field("Original Duration",   orig_dur,  suffix=" days") +
            field("Remaining Duration",  rem_dur,   suffix=" days") +
            field("Total Float",         tf_num,    suffix=" days") +
            field("Free Float",          ff_val,    suffix=" days"),
            unsafe_allow_html=True,
        )

    # Constraint row - only show when a constraint is set
    if "cstr_type" in row.index:
        cstr = row.get("cstr_type", "")
        if cstr and str(cstr).strip() not in ("", "None", "nan"):
            cstr_date = format_date(row.get("cstr_date") if "cstr_date" in row.index else None)
            st.markdown(
                f'<div style="background:#fef3c7;border-left:4px solid #f59e0b;'
                f'border-radius:6px;padding:10px 16px;margin-top:12px;">'
                f'<strong>Constraint:</strong> {cstr} &nbsp;|&nbsp; '
                f'<strong>Constraint Date:</strong> {cstr_date}</div>',
                unsafe_allow_html=True,
            )

    # -------------------------------------------------------------------------
    # PREDECESSORS & SUCCESSORS
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("##### Logic")

    pred_col, succ_col = st.columns(2)

    if not rels.empty and "task_id" in row.index:
        task_id = row["task_id"]

        # --- Predecessors ---
        with pred_col:
            st.markdown("**Predecessors**  *(what drives this activity)*")
            if "succ_task_id" in rels.columns:
                preds = rels[rels["succ_task_id"] == task_id].copy()
            else:
                preds = pd.DataFrame()

            if preds.empty:
                st.info("No predecessors -- this is an open start.")
            else:
                pred_display_cols = {
                    "pred_task_code": "Activity ID",
                    "pred_task_name": "Activity Name",
                    "rel_type":       "Link",
                    "lag_days":       "Lag (d)",
                }
                pred_show = {k: v for k, v in pred_display_cols.items()
                             if k in preds.columns}
                pred_df = preds[list(pred_show.keys())].rename(columns=pred_show)
                st.dataframe(pred_df, use_container_width=True, hide_index=True)

        # --- Successors ---
        with succ_col:
            st.markdown("**Successors**  *(what this activity drives)*")
            if "pred_task_id" in rels.columns:
                succs = rels[rels["pred_task_id"] == task_id].copy()
            else:
                succs = pd.DataFrame()

            if succs.empty:
                st.info("No successors -- this is an open finish.")
            else:
                succ_display_cols = {
                    "succ_task_code": "Activity ID",
                    "succ_task_name": "Activity Name",
                    "rel_type":       "Link",
                    "lag_days":       "Lag (d)",
                }
                succ_show = {k: v for k, v in succ_display_cols.items()
                             if k in succs.columns}
                succ_df = succs[list(succ_show.keys())].rename(columns=succ_show)
                st.dataframe(succ_df, use_container_width=True, hide_index=True)
    else:
        with pred_col:
            st.info("Relationship data not available.")
        with succ_col:
            st.info("Relationship data not available.")

    # -------------------------------------------------------------------------
    # EXPORT THIS ACTIVITY
    # -------------------------------------------------------------------------
    st.markdown("---")

    # Build a single-row detail export
    export_fields = {
        "Activity ID":        _col(row, "task_code"),
        "Activity Name":      _col(row, "task_name"),
        "WBS":                _col(row, "wbs_path"),
        "Activity Type":      _col(row, "task_type"),
        "Calendar":           _col(row, "calendar"),
        "Status":             _status_label(str(_col(row, "status", ""))),
        "% Complete":         _col(row, "phys_pct"),
        "Early Start":        early_start,
        "Early Finish":       early_finish,
        "Late Start":         late_start,
        "Late Finish":        late_finish,
        "Actual Start":       act_start,
        "Actual Finish":      act_finish,
        "Original Duration":  _col(row, "orig_dur_days"),
        "Remaining Duration": _col(row, "rem_dur_days"),
        "Total Float (days)": tf_num,
        "Free Float (days)":  _col(row, "free_float_days"),
        "Critical":           "Yes" if is_crit else "No",
        "Constraint":         _col(row, "cstr_type"),
        "Constraint Date":    format_date(row.get("cstr_date") if "cstr_date" in row.index else None),
    }
    detail_df = pd.DataFrame([export_fields])

    # Predecessors / successors for export
    export_sheets = {"Activity Detail": detail_df}

    if not rels.empty and "task_id" in row.index:
        task_id = row["task_id"]
        if not preds.empty:
            pred_exp_cols = [c for c in ["pred_task_code","pred_task_name","rel_type","lag_days"]
                             if c in preds.columns]
            export_sheets["Predecessors"] = preds[pred_exp_cols].rename(columns={
                "pred_task_code": "Activity ID", "pred_task_name": "Activity Name",
                "rel_type": "Link", "lag_days": "Lag (days)"
            })
        if not succs.empty:
            succ_exp_cols = [c for c in ["succ_task_code","succ_task_name","rel_type","lag_days"]
                             if c in succs.columns]
            export_sheets["Successors"] = succs[succ_exp_cols].rename(columns={
                "succ_task_code": "Activity ID", "succ_task_name": "Activity Name",
                "rel_type": "Link", "lag_days": "Lag (days)"
            })

    xls_bytes = export_df_to_excel(export_sheets)
    st.download_button(
        label="📥  Export Activity Detail to Excel",
        data=xls_bytes,
        file_name=f"activity_{_col(row, 'task_code', 'detail')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# -----------------------------------------------------------------------------
# PAGE: LOGIC TRACE
# -----------------------------------------------------------------------------

# Relationship type codes -> readable labels
REL_TYPE_LABELS = {
    "FS": "Finish-to-Start (FS)",
    "FF": "Finish-to-Finish (FF)",
    "SS": "Start-to-Start (SS)",
    "SF": "Start-to-Finish (SF)",
    "PR_FS": "Finish-to-Start (FS)",
    "PR_FF": "Finish-to-Finish (FF)",
    "PR_SS": "Start-to-Start (SS)",
    "PR_SF": "Start-to-Finish (SF)",
}


def _rel_label(code: str) -> str:
    return REL_TYPE_LABELS.get(str(code).strip(), str(code).strip() or "-")


def _crit_flag(tf) -> str:
    """Return a text critical flag suitable for display and export."""
    if tf is None:
        return "-"
    try:
        f = float(tf)
    except (TypeError, ValueError):
        return "-"
    if f < 0:
        return "Negative Float"
    if f == 0:
        return "Critical"
    if f <= 10:
        return "Near-Critical"
    return "Float"


def _build_full_trace_df(
    G: nx.DiGraph,
    rels_df: pd.DataFrame,
    task_lookup: dict,
    selected_id: str,
    trace_list: list,     # [(task_id, depth), ...]
    direction: str,       # "pred" | "succ" | "both"
) -> pd.DataFrame:
    """
    Build the trace results DataFrame.

    For each activity in trace_list, look up the relationship that
    connects it to the selected activity (or to the activity one step
    closer in the chain) so we can display the correct link type and lag.

    direction:
        "pred"  - activities are predecessors; depth counts backwards  (1 = direct pred)
        "succ"  - activities are successors;   depth counts forwards   (1 = direct succ)
        "both"  - mixed: negative depth = pred side, positive = succ side
    """
    rows = []

    for tid, depth in trace_list:
        t = task_lookup.get(tid, {})
        tf = t.get("total_float_days")

        # ---- find the relationship between this activity and the one at
        #      depth-1 in the chain (i.e. the step that led here) ----------
        rel_type_str = "-"
        lag_val = 0

        if not rels_df.empty:
            if direction in ("pred", "both") and depth > 0:
                # this activity is a predecessor of something; find its
                # outgoing relationship toward the selected direction
                mask = (
                    (rels_df.get("pred_task_id", pd.Series(dtype=str)) == tid) |
                    (rels_df.get("succ_task_id", pd.Series(dtype=str)) == tid)
                )
                rel_candidates = rels_df[mask]
            elif direction == "succ":
                mask = (
                    (rels_df.get("pred_task_id", pd.Series(dtype=str)) == tid) |
                    (rels_df.get("succ_task_id", pd.Series(dtype=str)) == tid)
                )
                rel_candidates = rels_df[mask]
            else:
                rel_candidates = pd.DataFrame()

            if not rel_candidates.empty and "rel_type" in rel_candidates.columns:
                rel_type_str = _rel_label(rel_candidates["rel_type"].iloc[0])
                lag_val = rel_candidates["lag_days"].iloc[0] if "lag_days" in rel_candidates.columns else 0
                try:
                    lag_val = float(lag_val) if lag_val is not None else 0
                except (TypeError, ValueError):
                    lag_val = 0

        rows.append({
            "Depth": depth,
            "Direction": "Predecessor" if depth < 0 or direction == "pred"
                         else ("Successor" if direction == "succ" else "Both"),
            "Activity ID":    t.get("task_code", tid),
            "Activity Name":  t.get("task_name", ""),
            "Link Type":      rel_type_str,
            "Lag (days)":     lag_val,
            "Start":          format_date(t.get("eff_start")),
            "Finish":         format_date(t.get("eff_finish")),
            "Total Float (d)": tf if tf is not None else "-",
            "Status":         _status_label(str(t.get("status", ""))),
            "Critical Flag":  _crit_flag(tf),
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("Depth").reset_index(drop=True)
    return df


def _summary_bar(label: str, value: int, colour: str) -> str:
    return (
        f'<div style="display:inline-block;background:{colour};color:white;'
        f'border-radius:8px;padding:8px 18px;margin:4px 6px 4px 0;font-size:13px;">'
        f'<strong>{value}</strong> {label}</div>'
    )


def page_logic_trace(data: dict, near_crit_days: float):
    """
    Logic Trace page.
    Search and select any activity, then trace its predecessor and successor
    chains through the schedule network. Results show depth, link type, lag,
    dates, float, WBS and critical status with filters and Excel export.
    """
    st.title("🔗 Logic Trace")
    st.caption(
        "Select an activity then trace its logic through the schedule network. "
        "Predecessors flow INTO the activity. Successors flow OUT of it."
    )

    tasks = data["tasks_df"]
    rels  = data["relationships_df"]

    # -------------------------------------------------------------------------
    # GUARD RAILS
    # -------------------------------------------------------------------------
    if tasks.empty:
        st.warning("No activities found in this programme.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    if rels.empty:
        st.warning(
            "**No relationship data found in this XER file.** "
            "Logic tracing requires both activities and relationships. "
            "Check the XER was exported with relationships included."
        )
        st.dataframe(
            tasks[["task_code", "task_name", "total_float_days", "status"]].head(50),
            use_container_width=True, hide_index=True,
        )
        return

    # -------------------------------------------------------------------------
    # BUILD NETWORK GRAPH
    # -------------------------------------------------------------------------
    G           = build_graph(tasks, rels)
    task_lookup = tasks.set_index("task_id").to_dict("index")

    # -------------------------------------------------------------------------
    # ACTIVITY SEARCH & SELECTOR
    # -------------------------------------------------------------------------
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">'
        'Find Activity</div>',
        unsafe_allow_html=True,
    )

    search_col, select_col = st.columns([1, 2])

    with search_col:
        search_text = st.text_input(
            "Search by ID or name",
            placeholder="e.g. A1000 or Excavation",
            key="lt_search",
            label_visibility="collapsed",
        )

    # Filter task list by search text
    if search_text.strip():
        mask = (
            tasks["task_code"].astype(str).str.contains(search_text.strip(), case=False, na=False) |
            tasks["task_name"].astype(str).str.contains(search_text.strip(), case=False, na=False)
        )
        filtered_tasks = tasks[mask]
        if filtered_tasks.empty:
            st.warning(f"No activities match '{search_text}'. Showing all activities.")
            filtered_tasks = tasks
    else:
        filtered_tasks = tasks

    with select_col:
        def _act_label(r):
            tf   = r.get("total_float_days")
            flag = " [CRITICAL]" if (tf is not None and safe_float(tf, 1) <= 0) else ""
            return f"{r.get('task_code','?')}  --  {r.get('task_name','?')}{flag}"

        act_labels = filtered_tasks.apply(_act_label, axis=1).tolist()

        if not act_labels:
            st.warning("No activities to select.")
            return

        selected_label = st.selectbox(
            "Select activity",
            options=act_labels,
            key="logic_trace_selector",
            label_visibility="collapsed",
        )

    sel_idx      = act_labels.index(selected_label)
    selected_row = filtered_tasks.iloc[sel_idx]
    selected_id  = selected_row["task_id"]
    sel_code     = str(selected_row.get("task_code", "-"))
    sel_name     = str(selected_row.get("task_name", "-"))

    # Clear results when activity changes
    if st.session_state.get("_trace_last_id") != selected_id:
        for k in ("trace_df", "trace_label", "trace_direction"):
            st.session_state.pop(k, None)
        st.session_state["_trace_last_id"] = selected_id

    # -------------------------------------------------------------------------
    # SELECTED ACTIVITY SUMMARY CARD
    # -------------------------------------------------------------------------
    sel_tf   = safe_float(selected_row.get("total_float_days"), None) if "total_float_days" in selected_row.index else None
    sel_ff   = safe_float(selected_row.get("free_float_days"),  None) if "free_float_days"  in selected_row.index else None
    sel_fcol = _float_color(sel_tf)
    sel_crit = bool(selected_row.get("is_critical", False)) if "is_critical" in selected_row.index else False
    sel_stat = _status_label(str(selected_row.get("status", "")))
    sel_scol = _status_colour(str(selected_row.get("status", "")))
    sel_wbs  = str(selected_row.get("wbs_path", "-")) if "wbs_path" in selected_row.index else "-"
    sel_start  = format_date(selected_row.get("eff_start")  if "eff_start"  in selected_row.index else None)
    sel_finish = format_date(selected_row.get("eff_finish") if "eff_finish" in selected_row.index else None)

    crit_pill = (
        '<span style="background:#dc2626;color:white;padding:2px 10px;'
        'border-radius:12px;font-size:11px;font-weight:700;margin-left:8px;">CRITICAL</span>'
        if sel_crit else ""
    )

    # Card layout: left = identity, right = schedule metrics
    card_left = f"""
        <div style="flex:1;min-width:0;">
            <div style="font-size:11px;color:#93c5fd;font-weight:600;
                        letter-spacing:1px;text-transform:uppercase;
                        margin-bottom:6px;">Selected Activity</div>
            <div style="font-size:20px;font-weight:800;white-space:nowrap;
                        overflow:hidden;text-overflow:ellipsis;">
                {sel_code}{crit_pill}
            </div>
            <div style="font-size:14px;color:#bfdbfe;margin-top:4px;
                        white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
                 title="{sel_name}">{sel_name}</div>
            <div style="font-size:12px;color:#64748B;margin-top:4px;
                        white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
                 title="{sel_wbs}">{sel_wbs}</div>
            <div style="margin-top:10px;display:flex;gap:6px;flex-wrap:wrap;">
                <span style="background:{sel_scol};color:white;padding:3px 10px;
                             border-radius:12px;font-size:11px;">{sel_stat}</span>
                <span style="background:{sel_fcol};color:white;padding:3px 10px;
                             border-radius:12px;font-size:11px;">
                    Float: {sel_tf if sel_tf is not None else "-"} days
                </span>
            </div>
        </div>"""

    card_right = f"""
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;
                    min-width:200px;align-self:center;">
            <div style="background:#0d2640;border-radius:6px;padding:8px 12px;">
                <div style="font-size:9px;color:#64748B;text-transform:uppercase;
                            letter-spacing:0.8px;">Start</div>
                <div style="font-size:13px;font-weight:600;color:#e2e8f0;
                            margin-top:2px;">{sel_start}</div>
            </div>
            <div style="background:#0d2640;border-radius:6px;padding:8px 12px;">
                <div style="font-size:9px;color:#64748B;text-transform:uppercase;
                            letter-spacing:0.8px;">Finish</div>
                <div style="font-size:13px;font-weight:600;color:#e2e8f0;
                            margin-top:2px;">{sel_finish}</div>
            </div>
            <div style="background:#0d2640;border-radius:6px;padding:8px 12px;">
                <div style="font-size:9px;color:#64748B;text-transform:uppercase;
                            letter-spacing:0.8px;">Total Float</div>
                <div style="font-size:15px;font-weight:700;color:#F5A623;
                            margin-top:2px;">{(str(sel_tf) + "d") if sel_tf is not None else "-"}</div>
            </div>
            <div style="background:#0d2640;border-radius:6px;padding:8px 12px;">
                <div style="font-size:9px;color:#64748B;text-transform:uppercase;
                            letter-spacing:0.8px;">Free Float</div>
                <div style="font-size:13px;font-weight:600;color:#e2e8f0;
                            margin-top:2px;">{(str(sel_ff) + "d") if sel_ff is not None else "-"}</div>
            </div>
        </div>"""

    st.markdown(
        f"""
        <div style="background:#1e3a5f;color:white;border-radius:12px;
                    padding:20px 24px;margin:10px 0 20px 0;
                    display:flex;gap:24px;align-items:flex-start;flex-wrap:wrap;">
            {card_left}
            {card_right}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # OPEN-END WARNINGS
    # -------------------------------------------------------------------------
    direct_preds = list(G.predecessors(selected_id))
    direct_succs = list(G.successors(selected_id))
    has_preds    = len(direct_preds) > 0
    has_succs    = len(direct_succs) > 0

    if not has_preds and not has_succs:
        st.error(
            f"**{sel_code} has no logic connections.** "
            "This activity is completely isolated in the programme network - "
            "it has neither predecessors nor successors."
        )
        return

    if not has_preds:
        st.warning(
            f"**Open Start:** {sel_code} has no predecessors. "
            "It is not driven by any logic and may have artificially high float."
        )
    if not has_succs:
        st.warning(
            f"**Open Finish:** {sel_code} has no successors. "
            "Nothing in the programme is dependent on it completing."
        )

    # -------------------------------------------------------------------------
    # QUICK STATS ROW
    # -------------------------------------------------------------------------
    all_pred_ids = [tid for tid, _ in trace_predecessors(G, selected_id)]
    all_succ_ids = [tid for tid, _ in trace_successors(G,  selected_id)]

    def count_crit_in(id_list):
        return sum(
            1 for tid in id_list
            if safe_float(task_lookup.get(tid, {}).get("total_float_days"), 1) <= 0
        )

    st.markdown(
        _summary_bar(f"direct predecessors",   len(direct_preds), "#1d4ed8") +
        _summary_bar(f"direct successors",     len(direct_succs), "#1d4ed8") +
        _summary_bar(f"total predecessors",    len(all_pred_ids), "#4338ca") +
        _summary_bar(f"total successors",      len(all_succ_ids), "#4338ca") +
        (_summary_bar("critical in pred network", count_crit_in(all_pred_ids), "#dc2626") if all_pred_ids else "") +
        (_summary_bar("critical in succ network", count_crit_in(all_succ_ids), "#dc2626") if all_succ_ids else ""),
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # TRACE BUTTONS
    # -------------------------------------------------------------------------
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">'
        'Choose what to trace</div>',
        unsafe_allow_html=True,
    )

    b1, b2, b3, b4 = st.columns(4)
    btn_dir_pred = b1.button(
        "◀  Direct Predecessors",
        key="btn_dp", use_container_width=True, disabled=not has_preds,
        help="Show only the activities that directly link into this activity (depth 1).",
    )
    btn_dir_succ = b2.button(
        "▶  Direct Successors",
        key="btn_ds", use_container_width=True, disabled=not has_succs,
        help="Show only the activities this activity directly links into (depth 1).",
    )
    btn_all_pred = b3.button(
        "◀◀  All Predecessors",
        key="btn_ap", use_container_width=True, disabled=not has_preds,
        help="Trace the full predecessor network back through all levels.",
    )
    btn_all_succ = b4.button(
        "▶▶  All Successors",
        key="btn_as", use_container_width=True, disabled=not has_succs,
        help="Trace the full successor network forward through all levels.",
    )

    # -------------------------------------------------------------------------
    # HANDLE BUTTONS
    # -------------------------------------------------------------------------
    new_trace = None
    new_label = None
    new_dir   = None

    if btn_dir_pred:
        new_trace = [(p, 1) for p in direct_preds]
        new_label = f"Direct Predecessors of {sel_code}"
        new_dir   = "pred"
    elif btn_dir_succ:
        new_trace = [(s, 1) for s in direct_succs]
        new_label = f"Direct Successors of {sel_code}"
        new_dir   = "succ"
    elif btn_all_pred:
        new_trace = trace_predecessors(G, selected_id)
        new_label = f"All Predecessors of {sel_code}"
        new_dir   = "pred"
    elif btn_all_succ:
        new_trace = trace_successors(G, selected_id)
        new_label = f"All Successors of {sel_code}"
        new_dir   = "succ"

    if new_trace is not None:
        raw_df = _build_full_trace_df(G, rels, task_lookup, selected_id, new_trace, new_dir)

        # Enrich with WBS column
        if not raw_df.empty and "Activity ID" in raw_df.columns:
            code_to_wbs = tasks.set_index("task_code")["wbs_path"].to_dict() if "wbs_path" in tasks.columns else {}
            raw_df.insert(
                raw_df.columns.get_loc("Activity Name") + 1,
                "WBS",
                raw_df["Activity ID"].map(code_to_wbs).fillna("-"),
            )

        st.session_state["trace_df"]        = raw_df
        st.session_state["trace_label"]     = new_label
        st.session_state["trace_direction"] = new_dir

    # -------------------------------------------------------------------------
    # RESULTS
    # -------------------------------------------------------------------------
    if "trace_df" not in st.session_state or st.session_state["trace_df"].empty:
        st.markdown(
            '<div style="background:#f0f9ff;border:2px dashed #93c5fd;border-radius:10px;'
            'padding:32px;text-align:center;color:#1e40af;margin-top:20px;">'
            '<div style="font-size:28px;margin-bottom:10px;">🔗</div>'
            '<strong style="font-size:15px;">Press a button above to trace the logic</strong><br>'
            '<span style="font-size:13px;color:#64748B;margin-top:6px;display:block;">'
            'Results will appear here.</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    trace_df  = st.session_state["trace_df"]
    trace_lbl = st.session_state.get("trace_label", "Trace Results")
    trace_dir = st.session_state.get("trace_direction", "pred")

    st.divider()

    # Counts
    n_res  = len(trace_df)
    n_crit = int((trace_df["Critical Flag"] == "Critical").sum())
    n_neg  = int((trace_df["Critical Flag"] == "Negative Float").sum())
    n_near = int((trace_df["Critical Flag"] == "Near-Critical").sum())

    st.markdown(
        f"<h4 style='color:#0B1F33;margin-bottom:6px;'>{trace_lbl}</h4>",
        unsafe_allow_html=True,
    )
    st.markdown(
        _summary_bar(f"activities",      n_res,  "#374151") +
        (_summary_bar("critical",        n_crit, "#dc2626") if n_crit else "") +
        (_summary_bar("negative float",  n_neg,  "#7f1d1d") if n_neg  else "") +
        (_summary_bar("near-critical",   n_near, "#d97706") if n_near else ""),
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    if trace_dir == "pred":
        st.caption("Depth = steps back from the selected activity. Depth 1 = directly connected.")
    elif trace_dir == "succ":
        st.caption("Depth = steps forward from the selected activity. Depth 1 = directly connected.")

    # -------------------------------------------------------------------------
    # RESULT FILTERS
    # -------------------------------------------------------------------------
    with st.expander("Filter results", expanded=False):
        fc1, fc2, fc3 = st.columns(3)

        filt_flag = fc1.selectbox(
            "Float status",
            ["All", "Critical & Negative Float", "Near-Critical", "Has Float"],
            key="lt_filt_flag",
        )

        filt_wbs = fc2.text_input(
            "WBS contains",
            placeholder="e.g. Civil",
            key="lt_filt_wbs",
        )

        filt_float_max = fc3.number_input(
            "Max float (days)",
            value=9999,
            step=1,
            min_value=-9999,
            key="lt_filt_float",
            help="Show only activities with total float at or below this value.",
        )

        # Date range filter
        fd1, fd2 = st.columns(2)
        filt_finish_from = fd1.text_input(
            "Finish on or after (dd/mm/yyyy)",
            placeholder="01/01/2024",
            key="lt_filt_df",
        )
        filt_finish_to = fd2.text_input(
            "Finish on or before (dd/mm/yyyy)",
            placeholder="31/12/2025",
            key="lt_filt_dt",
        )

    # Apply filters to a copy of trace_df
    display_df = trace_df.copy()

    if filt_flag == "Critical & Negative Float":
        display_df = display_df[display_df["Critical Flag"].isin(["Critical", "Negative Float"])]
    elif filt_flag == "Near-Critical":
        display_df = display_df[display_df["Critical Flag"] == "Near-Critical"]
    elif filt_flag == "Has Float":
        display_df = display_df[display_df["Critical Flag"] == "Float"]

    if "WBS" in display_df.columns and filt_wbs.strip():
        display_df = display_df[
            display_df["WBS"].astype(str).str.contains(filt_wbs.strip(), case=False, na=False)
        ]

    if filt_float_max < 9999 and "Total Float (d)" in display_df.columns:
        display_df = display_df[
            display_df["Total Float (d)"].apply(
                lambda v: safe_float(v, 9999) <= filt_float_max
            )
        ]

    def _parse_date_filter(s):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(s.strip(), fmt)
            except Exception:
                pass
        return None

    if filt_finish_from.strip() and "Finish" in display_df.columns:
        d_from = _parse_date_filter(filt_finish_from)
        if d_from:
            # Finish column is a formatted string - compare via task lookup
            finish_dates = {}
            for _, row in tasks.iterrows():
                finish_dates[str(row.get("task_code",""))] = row.get("eff_finish")
            display_df = display_df[
                display_df["Activity ID"].apply(
                    lambda code: (lambda fd: fd is not None and fd >= d_from)(finish_dates.get(code))
                )
            ]

    if filt_finish_to.strip() and "Finish" in display_df.columns:
        d_to = _parse_date_filter(filt_finish_to)
        if d_to:
            finish_dates = {str(r.get("task_code","")): r.get("eff_finish") for _, r in tasks.iterrows()}
            display_df = display_df[
                display_df["Activity ID"].apply(
                    lambda code: (lambda fd: fd is not None and fd <= d_to)(finish_dates.get(code))
                )
            ]

    n_display = len(display_df)
    if n_display < n_res:
        st.caption(f"Showing {n_display} of {n_res} activities after filters.")

    if display_df.empty:
        st.info("No activities match the current filters.")
    else:
        # Styled table
        def _colour_flag(val):
            return {
                "Critical":       "background-color:#fee2e2;color:#991b1b;font-weight:600;",
                "Negative Float": "background-color:#fecaca;color:#7f1d1d;font-weight:700;",
                "Near-Critical":  "background-color:#fef3c7;color:#92400e;font-weight:600;",
                "Float":          "background-color:#dcfce7;color:#166534;",
            }.get(val, "")

        styled = display_df.style.applymap(_colour_flag, subset=["Critical Flag"])
        st.dataframe(styled, use_container_width=True, hide_index=True,
                     height=min(600, 45 + n_display * 35))

    # -------------------------------------------------------------------------
    # TABS: FILTERED VIEWS
    # -------------------------------------------------------------------------
    if n_res > 5:
        st.markdown("<br>", unsafe_allow_html=True)
        tab_all, tab_crit, tab_near, tab_open = st.tabs(
            ["All Results", "Critical & Neg Float", "Near-Critical", "Open Ends in Chain"]
        )

        with tab_all:
            st.dataframe(trace_df, use_container_width=True, hide_index=True)

        with tab_crit:
            crit_df = trace_df[trace_df["Critical Flag"].isin(["Critical", "Negative Float"])]
            if crit_df.empty:
                st.success("No critical or negative float activities in this trace.")
            else:
                st.dataframe(crit_df, use_container_width=True, hide_index=True)

        with tab_near:
            near_df = trace_df[trace_df["Critical Flag"] == "Near-Critical"]
            if near_df.empty:
                st.success("No near-critical activities in this trace.")
            else:
                st.dataframe(near_df, use_container_width=True, hide_index=True)

        with tab_open:
            open_rows = []
            for act_code in trace_df["Activity ID"].unique():
                match = tasks[tasks["task_code"] == act_code]
                if match.empty:
                    continue
                mid  = match.iloc[0]["task_id"]
                no_p = len(list(G.predecessors(mid))) == 0
                no_s = len(list(G.successors(mid))) == 0
                if no_p or no_s:
                    t = task_lookup.get(mid, {})
                    open_rows.append({
                        "Activity ID":   t.get("task_code", mid),
                        "Activity Name": t.get("task_name", ""),
                        "Issue": ("No predecessors" if no_p else "") +
                                 (" | No successors" if no_s else ""),
                        "Float (d)":     t.get("total_float_days", "-"),
                    })
            if open_rows:
                st.dataframe(pd.DataFrame(open_rows), use_container_width=True, hide_index=True)
            else:
                st.success("No open-ended activities found in this trace.")

    # -------------------------------------------------------------------------
    # GANTT CHART
    # -------------------------------------------------------------------------
    st.divider()
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">'
        'Trace Timeline</div>',
        unsafe_allow_html=True,
    )

    gantt_data = trace_df.copy().merge(
        tasks[["task_code","eff_start","eff_finish"]].rename(columns={"task_code":"Activity ID"}),
        on="Activity ID", how="left",
    ).dropna(subset=["eff_start","eff_finish"])

    if not gantt_data.empty:
        gantt_data["Label"] = (
            gantt_data["Activity ID"].astype(str) + "  " +
            gantt_data["Activity Name"].astype(str).str[:38]
        )
        gantt_data["Bar Colour"] = gantt_data["Critical Flag"].map({
            "Critical":       "Critical",
            "Negative Float": "Critical",
            "Near-Critical":  "Near-Critical",
            "Float":          "Has Float",
        }).fillna("Has Float")

        fig = px.timeline(
            gantt_data.head(80),
            x_start="eff_start", x_end="eff_finish", y="Label",
            color="Bar Colour",
            color_discrete_map={
                "Critical":     "#dc2626",
                "Near-Critical":"#d97706",
                "Has Float":    "#2563eb",
            },
            title=f"Timeline: {trace_lbl}",
            labels={"Label": ""},
        )
        fig.update_yaxes(autorange="reversed")
        fig.add_vline(
            x=datetime.now(), line_dash="dot", line_color="#94A3B8",
            annotation_text="Today", annotation_position="top left",
        )
        fig.update_layout(
            legend_title_text="Float Status",
            height=max(300, min(700, 60 + len(gantt_data.head(80)) * 28)),
            margin=dict(l=10, r=10, t=40, b=10),
            plot_bgcolor="#ffffff",
            paper_bgcolor="#ffffff",
        )
        st.plotly_chart(fig, use_container_width=True)
        if len(gantt_data) > 80:
            st.caption(f"Showing first 80 of {len(gantt_data)} activities on the timeline.")
    else:
        st.info("No date data available to build the timeline.")

    # -------------------------------------------------------------------------
    # EXCEL EXPORT
    # -------------------------------------------------------------------------
    st.divider()

    summary_data = {
        "Item": [
            "Selected Activity ID", "Activity Name", "WBS",
            "Trace Type", "Total in chain",
            "Critical in chain", "Near-Critical in chain", "Negative Float in chain",
        ],
        "Value": [
            sel_code, sel_name, sel_wbs,
            trace_lbl, n_res, n_crit, n_near, n_neg,
        ],
    }

    export_sheets = {"Summary": pd.DataFrame(summary_data), "Full Trace": trace_df}

    crit_exp = trace_df[trace_df["Critical Flag"].isin(["Critical","Negative Float"])]
    near_exp = trace_df[trace_df["Critical Flag"] == "Near-Critical"]
    if not crit_exp.empty:
        export_sheets["Critical Activities"] = crit_exp
    if not near_exp.empty:
        export_sheets["Near-Critical"]       = near_exp
    if n_display < n_res and not display_df.empty:
        export_sheets["Filtered View"] = display_df

    xls_bytes = export_df_to_excel(export_sheets)

    st.download_button(
        label="📥  Export Logic Trace to Excel",
        data=xls_bytes,
        file_name=f"logic_trace_{sel_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Exports Summary, Full Trace, Critical Activities, Near-Critical and any filtered view.",
    )


# -----------------------------------------------------------------------------
# PAGE: CRITICAL PATH ANALYSIS
# -----------------------------------------------------------------------------

def page_critical_path(data: dict, near_crit_days: float):
    st.title("🚨 Critical Path Analysis")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities loaded.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Critical Activities", "Near-Critical", "Negative Float", "By WBS / Package"]
    )

    with tab1:
        critical = tasks[tasks["is_critical"]].sort_values("total_float_days")
        st.metric("Critical Activities", len(critical))
        disp = ["task_code", "task_name", "wbs_path", "eff_start", "eff_finish",
                "total_float_days", "status"]
        avail = [c for c in disp if c in critical.columns]
        st.dataframe(critical[avail], use_container_width=True)

        if not critical.empty and "eff_start" in critical.columns:
            st.subheader("Critical Path Gantt")
            gantt_df = critical.dropna(subset=["eff_start", "eff_finish"]).copy()
            gantt_df["Start"] = gantt_df["eff_start"]
            gantt_df["Finish"] = gantt_df["eff_finish"]
            gantt_df["Task"] = gantt_df["task_code"] + " - " + gantt_df["task_name"]
            if len(gantt_df) > 0:
                fig = px.timeline(
                    gantt_df.head(50),
                    x_start="Start", x_end="Finish", y="Task",
                    title="Critical Path Activities (top 50)",
                    color_discrete_sequence=["#dc2626"],
                )
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(fig, use_container_width=True)

        xls = export_df_to_excel({"Critical Path": critical[avail]})
        st.download_button("📥 Export Critical Path", xls,
                           "critical_path.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        near_crit = tasks[tasks["is_near_critical"]].sort_values("total_float_days")
        st.metric(f"Near-Critical (0 < float <= {near_crit_days}d)", len(near_crit))
        avail = [c for c in ["task_code","task_name","wbs_path","eff_start","eff_finish",
                              "total_float_days","status"] if c in near_crit.columns]
        st.dataframe(near_crit[avail], use_container_width=True)

    with tab3:
        neg = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f < 0)].sort_values("total_float_days")
        st.metric("Negative Float Activities", len(neg))
        if not neg.empty:
            st.warning("⚠️ Activities with negative float indicate the schedule cannot be met -- investigate immediately.")
            avail = [c for c in ["task_code","task_name","total_float_days","eff_start",
                                  "eff_finish","status"] if c in neg.columns]
            st.dataframe(neg[avail], use_container_width=True)

    with tab4:
        if "wbs_path" not in tasks.columns:
            st.info("WBS data not available.")
            return
        tasks["wbs_top"] = tasks["wbs_path"].apply(
            lambda x: str(x).split(" > ")[0] if pd.notna(x) else "Unknown"
        )
        wbs_crit = tasks.groupby("wbs_top").agg(
            total=("task_id", "count"),
            critical=("is_critical", "sum"),
            near_critical=("is_near_critical", "sum"),
        ).reset_index()
        wbs_crit["crit_%"] = (wbs_crit["critical"] / wbs_crit["total"] * 100).round(1)
        fig = px.bar(
            wbs_crit, x="wbs_top", y=["critical", "near_critical"],
            title="Critical & Near-Critical by WBS",
            labels={"value": "Activities", "wbs_top": "WBS"},
            color_discrete_map={"critical": "#dc2626", "near_critical": "#f59e0b"},
            barmode="group",
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(wbs_crit, use_container_width=True)


# -----------------------------------------------------------------------------
# PAGE: CRITICAL PATH TO SELECTED ACTIVITY
# -----------------------------------------------------------------------------

def _network_diagram_html(
    path_ids: list,
    all_pred_ids: list,
    task_lookup: dict,
    rels_df: pd.DataFrame,
) -> str:
    """
    Build a lightweight SVG-based network diagram showing the driving path
    as a horizontal chain, with branch predecessors shown above/below.
    Returns an HTML string ready for st.components.v1.html().
    """
    if not path_ids:
        return ""

    # Colour helpers
    def node_colour(tid):
        t  = task_lookup.get(tid, {})
        tf = safe_float(t.get("total_float_days"), 9999)
        if tf < 0:  return "#7f1d1d", "#fecaca"   # bg, border
        if tf == 0: return "#dc2626", "#fee2e2"
        if tf <= 10:return "#d97706", "#fef3c7"
        return       "#1d4ed8", "#dbeafe"

    BOX_W, BOX_H = 160, 54
    H_GAP, V_GAP = 40, 70
    n = len(path_ids)
    canvas_w = n * (BOX_W + H_GAP) + H_GAP
    canvas_h = 200

    svg_parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{canvas_w}" height="{canvas_h}" '
        f'style="font-family:Arial,sans-serif;font-size:11px;">'
    ]

    # Draw boxes for driving path (middle row y=70)
    mid_y = 80
    box_centres = {}
    for i, tid in enumerate(path_ids):
        t    = task_lookup.get(tid, {})
        x    = H_GAP + i * (BOX_W + H_GAP)
        y    = mid_y
        cx   = x + BOX_W // 2
        cy   = y + BOX_H // 2
        box_centres[tid] = (cx, cy)
        fc, bc = node_colour(tid)
        code = str(t.get("task_code", tid))[:14]
        name = str(t.get("task_name", ""))[:20]
        tf   = t.get("total_float_days")
        tf_s = f"Float: {tf}d" if tf is not None else ""
        svg_parts.append(
            f'<rect x="{x}" y="{y}" width="{BOX_W}" height="{BOX_H}" '
            f'rx="6" fill="{bc}" stroke="{fc}" stroke-width="2"/>'
            f'<text x="{cx}" y="{y+16}" text-anchor="middle" '
            f'font-weight="bold" fill="{fc}">{code}</text>'
            f'<text x="{cx}" y="{y+30}" text-anchor="middle" fill="#374151">{name}</text>'
            f'<text x="{cx}" y="{y+44}" text-anchor="middle" fill="{fc}" '
            f'font-size="10">{tf_s}</text>'
        )

        # Arrow from previous box
        if i > 0:
            prev_tid = path_ids[i - 1]
            px2, py2 = box_centres[prev_tid]
            # Get rel type
            rel_label = "FS"
            lag_label = ""
            if not rels_df.empty:
                rel = rels_df[
                    (rels_df.get("pred_task_id", pd.Series(dtype=str)) == prev_tid) &
                    (rels_df.get("succ_task_id", pd.Series(dtype=str)) == tid)
                ]
                if not rel.empty:
                    rel_label = str(rel["rel_type"].iloc[0])[-2:] if "rel_type" in rel.columns else "FS"
                    lag = safe_float(rel["lag_days"].iloc[0] if "lag_days" in rel.columns else 0, 0)
                    lag_label = f" +{int(lag)}d" if lag > 0 else (f" {int(lag)}d" if lag < 0 else "")
            ax1 = px2 + BOX_W // 2
            ax2 = x
            ay  = cy
            svg_parts.append(
                f'<line x1="{ax1}" y1="{ay}" x2="{ax2}" y2="{ay}" '
                f'stroke="#6b7280" stroke-width="2" marker-end="url(#arr)"/>'
                f'<text x="{(ax1+ax2)//2}" y="{ay-5}" text-anchor="middle" '
                f'fill="#6b7280" font-size="10">{rel_label}{lag_label}</text>'
            )

    # Arrow marker def
    svg_parts.insert(1,
        '<defs><marker id="arr" markerWidth="8" markerHeight="8" '
        'refX="6" refY="3" orient="auto">'
        '<path d="M0,0 L0,6 L8,3 z" fill="#6b7280"/>'
        '</marker></defs>'
    )

    svg_parts.append('</svg>')
    html = (
        '<div style="overflow-x:auto;background:#f8fafc;border:1px solid #e2e8f0;'
        'border-radius:8px;padding:12px;">'
        + "".join(svg_parts) +
        '</div>'
    )
    return html


def page_critical_path_to_activity(data: dict, near_crit_days: float):
    """
    Critical Path to Selected Activity page.

    Traces backwards through predecessor logic to identify the most likely
    driving chain into any selected activity or milestone.
    Uses float, finish dates, relationship type and lag to determine the driver.
    """
    st.title("🎯 Critical Path to Selected Activity")

    # -------------------------------------------------------------------------
    # EXPLANATION BANNER
    # -------------------------------------------------------------------------
    st.markdown(
        '<div style="background:#eff6ff;border-left:4px solid #3b82f6;'
        'border-radius:6px;padding:14px 18px;margin-bottom:18px;">'
        '<strong>How this works</strong><br>'
        'This shows the likely chain of activities driving the selected activity. '
        'It traces backwards through predecessor logic and identifies the most critical '
        'path based on total float, latest finish dates, relationship types and lag. '
        '<br><br>'
        '<em>This is based on available XER logic and float values and should be '
        'reviewed with the planner before being used for decision-making.</em>'
        '</div>',
        unsafe_allow_html=True,
    )

    tasks = data["tasks_df"]
    rels  = data["relationships_df"]

    # Guard rails
    if tasks.empty:
        st.warning("No activities found. Please upload an XER file first.")
        return
    if rels.empty:
        st.warning(
            "No relationship data found in this XER file. "
            "This page requires both activities and relationships."
        )
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    G           = build_graph(tasks, rels)
    task_lookup = tasks.set_index("task_id").to_dict("index")

    # -------------------------------------------------------------------------
    # ACTIVITY SELECTOR
    # -------------------------------------------------------------------------
    def _label(r):
        code = str(r.get("task_code", "?"))
        name = str(r.get("task_name", "?"))
        tf   = r.get("total_float_days")
        flag = " [CRITICAL]" if (tf is not None and safe_float(tf, 1) <= 0) else ""
        return f"{code}  --  {name}{flag}"

    act_labels = tasks.apply(_label, axis=1).tolist()
    selected_label = st.selectbox(
        "Select target activity or milestone",
        options=act_labels,
        key="cpta_selector",
        help="Choose the activity you want to understand the driving path for.",
    )
    sel_idx    = act_labels.index(selected_label)
    target_row = tasks.iloc[sel_idx]
    target_id  = target_row["task_id"]
    tgt_code   = str(target_row.get("task_code", "-"))
    tgt_name   = str(target_row.get("task_name", "-"))
    tgt_tf     = safe_float(target_row.get("total_float_days"), None) if "total_float_days" in target_row.index else None
    tgt_fcol   = _float_color(tgt_tf)
    tgt_crit   = bool(target_row.get("is_critical", False)) if "is_critical" in target_row.index else False
    tgt_stat   = _status_label(str(target_row.get("status", "")))
    tgt_scol   = _status_colour(str(target_row.get("status", "")))

    # Clear cached results when target changes
    if st.session_state.get("_cpta_last_id") != target_id:
        for k in ("cpta_path", "cpta_all_preds"):
            st.session_state.pop(k, None)
        st.session_state["_cpta_last_id"] = target_id

    # -------------------------------------------------------------------------
    # TARGET ACTIVITY BANNER
    # -------------------------------------------------------------------------
    crit_pill = (
        '<span style="background:#dc2626;color:white;padding:2px 10px;'
        'border-radius:12px;font-size:11px;font-weight:700;margin-left:8px;">CRITICAL</span>'
        if tgt_crit else ""
    )
    st.markdown(
        f"""
        <div style="background:#1e3a5f;color:white;border-radius:10px;
                    padding:16px 22px;margin:8px 0 18px 0;">
            <div style="font-size:12px;color:#93c5fd;font-weight:600;
                        letter-spacing:1px;text-transform:uppercase;">Target Activity</div>
            <div style="font-size:20px;font-weight:700;margin-top:4px;">
                {tgt_code}{crit_pill}
            </div>
            <div style="font-size:14px;color:#bfdbfe;margin-top:2px;">{tgt_name}</div>
            <div style="margin-top:10px;">
                <span style="background:{tgt_scol};color:white;padding:3px 10px;
                             border-radius:12px;font-size:12px;">{tgt_stat}</span>
                <span style="background:{tgt_fcol};color:white;padding:3px 10px;
                             border-radius:12px;font-size:12px;margin-left:6px;">
                    Float: {tgt_tf if tgt_tf is not None else "-"} days
                </span>
                <span style="color:#93c5fd;font-size:12px;margin-left:12px;">
                    Finish: {format_date(target_row.get("eff_finish") if "eff_finish" in target_row.index else None)}
                </span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Check if target has any predecessors at all
    direct_preds = list(G.predecessors(target_id))
    if not direct_preds:
        st.warning(
            f"**{tgt_code} has no predecessors.** This activity has an open start "
            "and is not driven by any logic in the programme. "
            "Nothing can be identified as the driving path."
        )
        return

    # -------------------------------------------------------------------------
    # RUN BUTTON
    # -------------------------------------------------------------------------
    run_col, _ = st.columns([1, 3])
    run_btn = run_col.button(
        "🔍  Find Driving Path",
        key="cpta_run",
        use_container_width=True,
        type="primary",
    )

    if run_btn:
        with st.spinner("Tracing predecessor network..."):
            driving_path   = driving_path_to_activity(G, tasks, rels, target_id)
            all_pred_pairs = trace_predecessors(G, target_id)
            all_pred_ids   = [p for p, _ in all_pred_pairs]
        st.session_state["cpta_path"]      = driving_path
        st.session_state["cpta_all_preds"] = all_pred_ids

    # -------------------------------------------------------------------------
    # RESULTS
    # -------------------------------------------------------------------------
    if "cpta_path" not in st.session_state:
        st.markdown(
            '<div style="background:#f0f9ff;border:1px dashed #93c5fd;border-radius:8px;'
            'padding:24px;text-align:center;color:#1e40af;margin-top:16px;">'
            '<strong>Press "Find Driving Path" above to run the analysis.</strong>'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    driving_path = st.session_state["cpta_path"]
    all_pred_ids = st.session_state["cpta_all_preds"]

    # ---- KEY METRICS --------------------------------------------------------
    chain_tasks = tasks[tasks["task_id"].isin(driving_path)]
    n_chain     = len(driving_path)
    min_float   = chain_tasks["total_float_days"].min() if "total_float_days" in chain_tasks.columns else None
    n_crit_chain = int((chain_tasks["total_float_days"].apply(
        lambda f: safe_float(f, 1) <= 0
    )).sum()) if "total_float_days" in chain_tasks.columns else 0
    n_neg_chain  = int((chain_tasks["total_float_days"].apply(
        lambda f: safe_float(f, 0) < 0
    )).sum()) if "total_float_days" in chain_tasks.columns else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Activities in Driving Chain",  n_chain)
    m2.metric("Lowest Float in Chain",         f"{min_float:.1f} days" if min_float is not None else "-")
    m3.metric("Critical in Chain",             n_crit_chain)
    m4.metric("Total Predecessor Network",     len(all_pred_ids))

    if n_neg_chain > 0:
        st.error(
            f"⚠️ **{n_neg_chain} activit{'y' if n_neg_chain == 1 else 'ies'} with negative float** "
            "in the driving chain. The current schedule cannot meet its target dates for this path."
        )

    st.divider()

    # ---- TABS ---------------------------------------------------------------
    tab_path, tab_network, tab_all_preds, tab_constraints = st.tabs([
        "Driving Path", "Network Diagram", "All Predecessors", "Constraints & Issues"
    ])

    # =========================================================================
    # TAB 1: DRIVING PATH TABLE
    # =========================================================================
    with tab_path:
        st.markdown(
            "The table below shows the most likely chain of activities driving "
            f"**{tgt_code}**, ordered from the earliest activity to the target. "
            "Activities are selected based on lowest float, latest finish date "
            "and relationship constraints."
        )

        path_rows = []
        for i, tid in enumerate(driving_path):
            t         = task_lookup.get(tid, {})
            tf        = t.get("total_float_days")
            is_target = (tid == target_id)

            # Relationship to next activity in chain
            rel_label = "-"
            lag_val   = 0
            if i < len(driving_path) - 1:
                next_tid = driving_path[i + 1]
                if not rels.empty:
                    rel = rels[
                        (rels.get("pred_task_id", pd.Series(dtype=str)) == tid) &
                        (rels.get("succ_task_id", pd.Series(dtype=str)) == next_tid)
                    ]
                    if not rel.empty:
                        rel_label = _rel_label(rel["rel_type"].iloc[0] if "rel_type" in rel.columns else "FS")
                        lag_val   = safe_float(rel["lag_days"].iloc[0] if "lag_days" in rel.columns else 0, 0)

            cstr = str(t.get("cstr_type", "")) if "cstr_type" in t else ""
            has_cstr = cstr.strip() not in ("", "None", "nan")

            path_rows.append({
                "Step":            i + 1,
                "Activity ID":     t.get("task_code", tid),
                "Activity Name":   t.get("task_name", ""),
                "Start":           format_date(t.get("eff_start")),
                "Finish":          format_date(t.get("eff_finish")),
                "Orig Dur (d)":    t.get("orig_dur_days", "-"),
                "Total Float (d)": tf if tf is not None else "-",
                "Link to Next":    rel_label if not is_target else "-",
                "Lag (d)":         lag_val if not is_target else "-",
                "Critical Flag":   _crit_flag(tf),
                "Constraint":      cstr if has_cstr else "",
                "Status":          _status_label(str(t.get("status", ""))),
                "Target":          "TARGET" if is_target else "",
            })

        path_df = pd.DataFrame(path_rows)

        # Colour code
        def _style_path_row(row):
            flag = row.get("Critical Flag", "")
            is_tgt = row.get("Target", "") == "TARGET"
            if is_tgt:
                return ["background-color:#1e3a5f;color:white;font-weight:700;"] * len(row)
            colour_map = {
                "Negative Float": "background-color:#fecaca;",
                "Critical":       "background-color:#fee2e2;",
                "Near-Critical":  "background-color:#fef3c7;",
            }
            style = colour_map.get(flag, "")
            return [style] * len(row)

        styled_path = path_df.style.apply(_style_path_row, axis=1)
        st.dataframe(styled_path, use_container_width=True, hide_index=True)

        # Gantt for driving path
        st.markdown("**Driving Path Timeline**")
        gantt_src = chain_tasks.dropna(subset=["eff_start","eff_finish"]).copy() if "eff_start" in chain_tasks.columns else pd.DataFrame()
        if not gantt_src.empty:
            gantt_src = gantt_src.merge(
                tasks[["task_id","task_code","task_name"]],
                on="task_id", how="left", suffixes=("","_t")
            )
            gantt_src["Label"]   = gantt_src["task_code"].astype(str) + "  " + gantt_src["task_name"].astype(str).str[:35]
            gantt_src["Colour"]  = gantt_src["task_id"].apply(
                lambda t: "Target" if t == target_id else (
                    "Critical" if safe_float(task_lookup.get(t,{}).get("total_float_days"), 1) <= 0
                    else "Near-Critical" if safe_float(task_lookup.get(t,{}).get("total_float_days"), 11) <= 10
                    else "Has Float"
                )
            )
            fig = px.timeline(
                gantt_src,
                x_start="eff_start", x_end="eff_finish", y="Label",
                color="Colour",
                color_discrete_map={
                    "Target":       "#1e3a5f",
                    "Critical":     "#dc2626",
                    "Near-Critical":"#d97706",
                    "Has Float":    "#2563eb",
                },
                title=f"Driving Path to {tgt_code}",
            )
            fig.update_yaxes(autorange="reversed")
            fig.add_vline(
                x=datetime.now(), line_dash="dot", line_color="#6b7280",
                annotation_text="Today", annotation_position="top left",
            )
            fig.update_layout(
                height=max(280, 50 + len(gantt_src) * 30),
                margin=dict(l=10, r=10, t=40, b=10),
                legend_title_text="Float Status",
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No date data available for the Gantt chart.")

    # =========================================================================
    # TAB 2: NETWORK DIAGRAM
    # =========================================================================
    with tab_network:
        st.markdown(
            "A simple left-to-right network diagram of the driving path. "
            "Each box shows the activity ID, name and float. "
            "Colours: **red** = critical/negative float, **amber** = near-critical, **blue** = has float, "
            "**navy** = target activity."
        )

        if len(driving_path) > 0:
            diagram_html = _network_diagram_html(
                driving_path, all_pred_ids, task_lookup, rels
            )
            if diagram_html:
                import streamlit.components.v1 as components
                n_boxes = len(driving_path)
                diagram_w = n_boxes * 200 + 80
                components.html(diagram_html, height=220, scrolling=True)
            else:
                st.info("Could not generate network diagram.")

            st.caption(
                "Note: The diagram shows the identified driving path only. "
                "Use the All Predecessors tab to see the full predecessor network."
            )
        else:
            st.info("No path data to display.")

    # =========================================================================
    # TAB 3: ALL PREDECESSORS
    # =========================================================================
    with tab_all_preds:
        all_pred_tasks = tasks[tasks["task_id"].isin(all_pred_ids)].copy()

        if all_pred_tasks.empty:
            st.info("No predecessor activities found.")
        else:
            n_ap     = len(all_pred_tasks)
            n_ap_crit = int((all_pred_tasks["total_float_days"].apply(
                lambda f: safe_float(f, 1) <= 0
            )).sum()) if "total_float_days" in all_pred_tasks.columns else 0

            st.markdown(
                _summary_bar(f"total predecessors", n_ap, "#374151") +
                (_summary_bar("critical", n_ap_crit, "#dc2626") if n_ap_crit else ""),
                unsafe_allow_html=True,
            )
            st.markdown("<br>", unsafe_allow_html=True)
            st.caption(
                "All activities in the predecessor network of the target activity, "
                "sorted by float (most critical first)."
            )

            all_pred_tasks = all_pred_tasks.sort_values("total_float_days")

            AP_COLS = {
                "task_code":        "Activity ID",
                "task_name":        "Activity Name",
                "wbs_path":         "WBS",
                "eff_start":        "Start",
                "eff_finish":       "Finish",
                "total_float_days": "Float (d)",
                "status":           "Status",
                "is_critical":      "Critical",
            }
            ap_show = {k: v for k, v in AP_COLS.items() if k in all_pred_tasks.columns}
            ap_df   = all_pred_tasks[list(ap_show.keys())].rename(columns=ap_show).copy()
            for col in ["Start","Finish"]:
                if col in ap_df.columns:
                    ap_df[col] = ap_df[col].apply(format_date)
            if "Critical" in ap_df.columns:
                ap_df["Critical"] = ap_df["Critical"].apply(lambda x: "Yes" if x else "")
            if "Status" in ap_df.columns:
                ap_df["Status"] = ap_df["Status"].apply(_status_label)

            st.dataframe(ap_df, use_container_width=True, hide_index=True, height=400)

    # =========================================================================
    # TAB 4: CONSTRAINTS & ISSUES
    # =========================================================================
    with tab_constraints:
        st.markdown(
            "Activities in the driving path or predecessor network that have "
            "constraints, negative float, or other schedule quality issues."
        )

        issues_found = False

        # --- Negative float in driving path ---
        neg_in_path = chain_tasks[
            chain_tasks["total_float_days"].apply(lambda f: safe_float(f, 0) < 0)
        ] if "total_float_days" in chain_tasks.columns else pd.DataFrame()

        if not neg_in_path.empty:
            issues_found = True
            st.markdown(
                '<div style="background:#fef2f2;border-left:4px solid #dc2626;'
                'border-radius:6px;padding:10px 14px;margin-bottom:12px;">'
                f'<strong>⚠️ Negative Float in Driving Chain ({len(neg_in_path)} activities)</strong><br>'
                'These activities are beyond their target dates. '
                'The driving chain cannot currently meet its schedule.'
                '</div>',
                unsafe_allow_html=True,
            )
            neg_cols = {k: v for k, v in {
                "task_code":"Activity ID","task_name":"Activity Name",
                "total_float_days":"Float (d)","eff_finish":"Finish","status":"Status"
            }.items() if k in neg_in_path.columns}
            neg_disp = neg_in_path[list(neg_cols.keys())].rename(columns=neg_cols).copy()
            if "Finish" in neg_disp.columns:
                neg_disp["Finish"] = neg_disp["Finish"].apply(format_date)
            st.dataframe(neg_disp, use_container_width=True, hide_index=True)

        # --- Constraints in driving path ---
        if "cstr_type" in chain_tasks.columns:
            constrained = chain_tasks[
                chain_tasks["cstr_type"].apply(
                    lambda x: bool(x) and str(x).strip() not in ("","None","nan")
                )
            ]
            if not constrained.empty:
                issues_found = True
                st.markdown(
                    '<div style="background:#fffbeb;border-left:4px solid #f59e0b;'
                    'border-radius:6px;padding:10px 14px;margin-bottom:12px;">'
                    f'<strong>Constraints in Driving Chain ({len(constrained)} activities)</strong><br>'
                    'Constraints override schedule logic and can cause artificial float or '
                    'negative float. Each one should be reviewed with the planner.'
                    '</div>',
                    unsafe_allow_html=True,
                )
                cstr_cols = {k: v for k, v in {
                    "task_code":"Activity ID","task_name":"Activity Name",
                    "cstr_type":"Constraint Type","cstr_date":"Constraint Date",
                    "total_float_days":"Float (d)"
                }.items() if k in constrained.columns}
                cstr_disp = constrained[list(cstr_cols.keys())].rename(columns=cstr_cols).copy()
                if "Constraint Date" in cstr_disp.columns:
                    cstr_disp["Constraint Date"] = cstr_disp["Constraint Date"].apply(format_date)
                st.dataframe(cstr_disp, use_container_width=True, hide_index=True)

        # --- Constraints in full predecessor network ---
        all_pred_tasks_full = tasks[tasks["task_id"].isin(all_pred_ids)].copy()
        if "cstr_type" in all_pred_tasks_full.columns:
            all_constrained = all_pred_tasks_full[
                all_pred_tasks_full["cstr_type"].apply(
                    lambda x: bool(x) and str(x).strip() not in ("","None","nan")
                )
            ]
            if not all_constrained.empty:
                issues_found = True
                with st.expander(f"Constraints in Full Predecessor Network ({len(all_constrained)})"):
                    cstr_cols2 = {k: v for k, v in {
                        "task_code":"Activity ID","task_name":"Activity Name",
                        "cstr_type":"Constraint Type","cstr_date":"Constraint Date",
                        "total_float_days":"Float (d)"
                    }.items() if k in all_constrained.columns}
                    cstr_disp2 = all_constrained[list(cstr_cols2.keys())].rename(columns=cstr_cols2).copy()
                    if "Constraint Date" in cstr_disp2.columns:
                        cstr_disp2["Constraint Date"] = cstr_disp2["Constraint Date"].apply(format_date)
                    st.dataframe(cstr_disp2, use_container_width=True, hide_index=True)

        # --- High lag in driving path relationships ---
        if not rels.empty and "lag_days" in rels.columns:
            path_set  = set(driving_path)
            path_rels = rels[
                rels.get("pred_task_id", pd.Series(dtype=str)).isin(path_set) &
                rels.get("succ_task_id", pd.Series(dtype=str)).isin(path_set)
            ]
            high_lag = path_rels[
                path_rels["lag_days"].apply(lambda l: abs(safe_float(l, 0)) > 5)
            ] if not path_rels.empty else pd.DataFrame()

            if not high_lag.empty:
                issues_found = True
                st.markdown(
                    '<div style="background:#eff6ff;border-left:4px solid #3b82f6;'
                    'border-radius:6px;padding:10px 14px;margin-bottom:12px;">'
                    f'<strong>Significant Lag in Driving Path ({len(high_lag)} relationships)</strong><br>'
                    'Lag of more than 5 days can hide logic issues and affect float calculations.'
                    '</div>',
                    unsafe_allow_html=True,
                )
                lag_cols = {k: v for k, v in {
                    "pred_task_code":"From","pred_task_name":"From Name",
                    "succ_task_code":"To","succ_task_name":"To Name",
                    "rel_type":"Link","lag_days":"Lag (d)"
                }.items() if k in high_lag.columns}
                if lag_cols:
                    st.dataframe(
                        high_lag[list(lag_cols.keys())].rename(columns=lag_cols),
                        use_container_width=True, hide_index=True,
                    )

        if not issues_found:
            st.success(
                "No constraints, negative float or significant lag found in the driving path. "
                "The chain appears logically sound."
            )

    # =========================================================================
    # EXCEL EXPORT
    # =========================================================================
    st.divider()

    # Rebuild path_df for export (already built above in tab_path)
    export_path_df = pd.DataFrame(path_rows) if path_rows else pd.DataFrame()

    summary_rows = {
        "Item":  [
            "Target Activity ID", "Target Activity Name", "Target Finish",
            "Target Float (days)", "Activities in Driving Chain",
            "Lowest Float in Chain", "Critical in Chain", "Negative Float in Chain",
            "Total Predecessor Network",
        ],
        "Value": [
            tgt_code, tgt_name,
            format_date(target_row.get("eff_finish") if "eff_finish" in target_row.index else None),
            tgt_tf,
            n_chain, min_float, n_crit_chain, n_neg_chain, len(all_pred_ids),
        ],
    }

    export_sheets = {
        "Summary":        pd.DataFrame(summary_rows),
        "Driving Path":   export_path_df,
        "All Predecessors": ap_df if not all_pred_tasks.empty else pd.DataFrame(columns=["No data"]),
    }

    if not neg_in_path.empty:
        export_sheets["Negative Float"] = neg_disp
    if "cstr_type" in chain_tasks.columns and not constrained.empty:
        export_sheets["Constraints"] = cstr_disp

    xls_bytes = export_df_to_excel(export_sheets)

    dl_col, _ = st.columns([1, 3])
    dl_col.download_button(
        label="📥  Export Driving Path Report to Excel",
        data=xls_bytes,
        file_name=f"driving_path_{tgt_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Exports Summary, Driving Path, All Predecessors, Negative Float and Constraints sheets.",
        use_container_width=True,
    )

# -----------------------------------------------------------------------------
# PAGE: LABOUR HISTOGRAM
# -----------------------------------------------------------------------------

def page_labour_histogram(data: dict):
    st.title("👷 Labour Histogram")

    task_res = data["task_resources_df"]
    tasks = data["tasks_df"]
    resources = data["resources_df"]

    if task_res.empty:
        st.markdown("""
        <div class="warn-box">
        ⚠️ <strong>No resource loading found in this XER file.</strong><br>
        This usually means the programme was not resourced in P6, or resource data was not exported.
        <br><br>
        You can upload a separate resource CSV or Excel file below.
        </div>
        """, unsafe_allow_html=True)

        st.subheader("Upload Resource Loading File")
        res_file = st.file_uploader("Upload CSV or Excel (columns: task_code, rsrc_name, target_qty, target_start, target_finish)", type=["csv","xlsx"])
        if res_file:
            try:
                if res_file.name.endswith(".csv"):
                    task_res = pd.read_csv(res_file)
                else:
                    task_res = pd.read_excel(res_file)
                for col in ["target_start","target_finish"]:
                    if col in task_res.columns:
                        task_res[col] = pd.to_datetime(task_res[col], errors="coerce")
                st.success(f"Loaded {len(task_res)} resource rows.")
            except Exception as e:
                st.error(f"Could not read resource file: {e}")
                return
        else:
            return

    # Merge with task info and resource names
    if not tasks.empty and "task_id" in task_res.columns:
        task_res = task_res.merge(
            tasks[["task_id","task_code","task_name","wbs_path","is_critical" if "is_critical" in tasks.columns else "task_id"]].drop_duplicates(),
            on="task_id", how="left", suffixes=("","_task")
        )
    if not resources.empty and "rsrc_id" in task_res.columns:
        task_res = task_res.merge(resources[["rsrc_id","rsrc_name"]], on="rsrc_id", how="left", suffixes=("","_res"))
        if "rsrc_name_res" in task_res.columns:
            task_res["rsrc_name"] = task_res["rsrc_name_res"].fillna(task_res.get("rsrc_name",""))

    # Expand resource loading to weekly intervals
    def expand_to_weeks(df):
        rows = []
        for _, r in df.iterrows():
            s = pd.to_datetime(r.get("target_start") or r.get("target_start_date"))
            e = pd.to_datetime(r.get("target_finish") or r.get("target_end_date"))
            if pd.isna(s) or pd.isna(e) or s > e:
                continue
            qty = safe_float(r.get("target_qty", 0), 0)
            if qty == 0:
                continue
            weeks = max(1, math.ceil((e - s).days / 7))
            qty_per_week = qty / weeks
            current = s
            for _ in range(weeks):
                rows.append({
                    "week": current.to_period("W").start_time,
                    "month": current.to_period("M").start_time,
                    "qty": qty_per_week,
                    "rsrc_name": r.get("rsrc_name","Unknown"),
                    "task_code": r.get("task_code",""),
                    "task_name": r.get("task_name",""),
                    "wbs_path": r.get("wbs_path",""),
                })
                current += timedelta(weeks=1)
        return pd.DataFrame(rows)

    weekly = expand_to_weeks(task_res)

    if weekly.empty:
        st.warning("Could not generate histogram -- resource dates or quantities may be missing.")
        return

    # Filters
    st.sidebar.divider()
    st.sidebar.subheader("Labour Filters")
    all_resources = sorted(weekly["rsrc_name"].unique().tolist())
    sel_res = st.sidebar.multiselect("Resource / Trade", all_resources, default=all_resources)
    if sel_res:
        weekly = weekly[weekly["rsrc_name"].isin(sel_res)]

    # Metrics
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Planned Hours", f"{weekly['qty'].sum():,.0f}")
    weekly_totals = weekly.groupby("week")["qty"].sum()
    c2.metric("Peak Week (hrs)", f"{weekly_totals.max():,.0f}" if not weekly_totals.empty else "-")
    c3.metric("Average Week (hrs)", f"{weekly_totals.mean():,.0f}" if not weekly_totals.empty else "-")

    tab1, tab2, tab3, tab4 = st.tabs(["By Week", "By Month", "By Resource", "By WBS"])

    with tab1:
        weekly_sum = weekly.groupby("week")["qty"].sum().reset_index()
        fig = px.bar(weekly_sum, x="week", y="qty",
                     title="Labour Loading by Week (Hours)",
                     labels={"week":"Week","qty":"Hours"},
                     color_discrete_sequence=["#2563eb"])
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        monthly_sum = weekly.groupby("month")["qty"].sum().reset_index()
        fig = px.bar(monthly_sum, x="month", y="qty",
                     title="Labour Loading by Month (Hours)",
                     labels={"month":"Month","qty":"Hours"},
                     color_discrete_sequence=["#1e3a5f"])
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        res_sum = weekly.groupby("rsrc_name")["qty"].sum().reset_index().sort_values("qty", ascending=False)
        fig = px.bar(res_sum, x="rsrc_name", y="qty",
                     title="Total Hours by Resource / Trade",
                     labels={"rsrc_name":"Resource","qty":"Hours"},
                     color_discrete_sequence=["#7c3aed"])
        st.plotly_chart(fig, use_container_width=True)

        # By week and resource stacked
        if len(sel_res) <= 10:
            by_res_week = weekly.groupby(["week","rsrc_name"])["qty"].sum().reset_index()
            fig2 = px.bar(by_res_week, x="week", y="qty", color="rsrc_name",
                          title="Weekly Labour by Resource",
                          labels={"week":"Week","qty":"Hours","rsrc_name":"Resource"})
            st.plotly_chart(fig2, use_container_width=True)

    with tab4:
        if "wbs_path" in weekly.columns:
            weekly["wbs_top"] = weekly["wbs_path"].apply(
                lambda x: str(x).split(" > ")[0] if pd.notna(x) and x else "Unknown"
            )
            wbs_sum = weekly.groupby("wbs_top")["qty"].sum().reset_index().sort_values("qty", ascending=False)
            fig = px.bar(wbs_sum, x="qty", y="wbs_top", orientation="h",
                         title="Total Hours by WBS",
                         color_discrete_sequence=["#059669"])
            st.plotly_chart(fig, use_container_width=True)

    # Export
    xls = export_df_to_excel({
        "Weekly Labour": weekly.groupby(["week","rsrc_name"])["qty"].sum().reset_index(),
        "Monthly Labour": weekly.groupby(["month","rsrc_name"])["qty"].sum().reset_index(),
        "By Resource": res_sum,
    })
    st.download_button("📥 Export Labour Data", xls, "labour_histogram.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# PAGE: SCHEDULE HEALTH CHECK
# -----------------------------------------------------------------------------

def page_health_check(data: dict, near_crit_days: float):
    st.title("🩺 Schedule Health Check")
    st.markdown("> Automated quality checks to identify common schedule issues.")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities loaded.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # Build predecessor/successor sets
    tasks_with_pred = set()
    tasks_with_succ = set()
    if not rels.empty:
        tasks_with_pred = set(rels["succ_task_id"].dropna()) if "succ_task_id" in rels.columns else set()
        tasks_with_succ = set(rels["pred_task_id"].dropna()) if "pred_task_id" in rels.columns else set()

    # Define checks
    checks = []

    # 1. No predecessors (excl. milestones at start)
    no_pred = tasks[~tasks["task_id"].isin(tasks_with_pred)]
    checks.append({
        "Check": "No Predecessors",
        "Count": len(no_pred),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Activities with no predecessors are open-ended. They cannot be driven by logic and may cause float calculation issues.",
        "df": no_pred,
    })

    # 2. No successors
    no_succ = tasks[~tasks["task_id"].isin(tasks_with_succ)]
    checks.append({
        "Check": "No Successors",
        "Count": len(no_succ),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Activities with no successors are open-ended and may have artificially high float.",
        "df": no_succ,
    })

    # 3. Negative float
    neg_float = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f < 0)]
    checks.append({
        "Check": "Negative Float",
        "Count": len(neg_float),
        "Severity": "🔴 Critical",
        "Why It Matters": "Negative float means the current schedule cannot meet its target dates. Immediate attention required.",
        "df": neg_float,
    })

    # 4. High float (> 60 days)
    high_float = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f > 60)]
    checks.append({
        "Check": "Very High Float (>60 days)",
        "Count": len(high_float),
        "Severity": "ℹ️ Info",
        "Why It Matters": "Activities with very high float may have missing logic or may not be properly constrained.",
        "df": high_float,
    })

    # 5. Excessive duration (> 60 working days)
    excess_dur = tasks[tasks["orig_dur_days"].apply(lambda d: d is not None and d > 60)]
    checks.append({
        "Check": "Excessive Duration (>60 days)",
        "Count": len(excess_dur),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Very long activities are difficult to control and should usually be broken down into smaller work packages.",
        "df": excess_dur,
    })

    # 6. Constraints
    constrained = tasks[tasks["cstr_type"].apply(
        lambda x: bool(x) and str(x).strip() not in ("", "None")
    )] if "cstr_type" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Constrained Activities",
        "Count": len(constrained),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Constraints override schedule logic and can create artificial float or negative float. Each constraint should be justified.",
        "df": constrained,
    })

    # 7. Excessive lag (> 10 days)
    if not rels.empty and "lag_days" in rels.columns:
        high_lag = rels[rels["lag_days"].apply(lambda l: l is not None and abs(safe_float(l,0)) > 10)]
        checks.append({
            "Check": "Excessive Lag (|lag| > 10 days)",
            "Count": len(high_lag),
            "Severity": "⚠️ Warning",
            "Why It Matters": "Excessive lag can hide critical path issues. Lag should be replaced with properly sequenced activities.",
            "df": high_lag,
        })

    # 8. Missing dates
    missing_dates = tasks[tasks["eff_start"].isna() | tasks["eff_finish"].isna()]
    checks.append({
        "Check": "Missing Start or Finish Dates",
        "Count": len(missing_dates),
        "Severity": "🔴 Critical",
        "Why It Matters": "Activities with no dates cannot be scheduled or reported on.",
        "df": missing_dates,
    })

    # 9. Actual dates in future
    now = datetime.now()
    future_actuals = tasks[
        tasks["act_start"].apply(lambda d: d is not None and d > now) |
        tasks["act_finish"].apply(lambda d: d is not None and d > now)
    ] if "act_start" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Future Actual Dates",
        "Count": len(future_actuals),
        "Severity": "🔴 Critical",
        "Why It Matters": "Actual start/finish dates should not be in the future. This indicates data entry errors.",
        "df": future_actuals,
    })

    # 10. Critical not started
    crit_not_started = tasks[
        tasks["is_critical"] &
        tasks["status"].apply(lambda s: str(s) in ("TK_NotStart", "Not Started") if pd.notna(s) else False)
    ] if "status" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Critical Activities Not Started",
        "Count": len(crit_not_started),
        "Severity": "🔴 Critical",
        "Why It Matters": "Critical activities that haven't started need immediate attention to avoid slippage.",
        "df": crit_not_started,
    })

    # 11. Near-critical due in 8 weeks
    eight_weeks = now + timedelta(weeks=8)
    near_due = tasks[
        tasks["is_near_critical"] &
        tasks["eff_finish"].apply(lambda d: d is not None and d <= eight_weeks)
    ] if "eff_finish" in tasks.columns else pd.DataFrame()
    checks.append({
        "Check": "Near-Critical Due in 8 Weeks",
        "Count": len(near_due),
        "Severity": "⚠️ Warning",
        "Why It Matters": "Near-critical activities finishing soon may become critical if not progressed.",
        "df": near_due,
    })

    # Scorecard
    st.subheader("Health Check Scorecard")
    score_data = [
        {"Check": c["Check"], "Count": c["Count"], "Severity": c["Severity"]}
        for c in checks
    ]
    score_df = pd.DataFrame(score_data)
    st.dataframe(score_df, use_container_width=True)

    # Detail per check
    st.divider()
    for chk in checks:
        with st.expander(f"{chk['Severity']} -- {chk['Check']} ({chk['Count']})"):
            st.markdown(f"**Why it matters:** {chk['Why It Matters']}")
            df = chk["df"]
            if not df.empty:
                disp = [c for c in ["task_code","task_name","wbs_path","eff_start",
                                     "eff_finish","total_float_days","status",
                                     "cstr_type","lag_days"] if c in df.columns]
                st.dataframe(df[disp].head(100), use_container_width=True)
                # Export individual check
                xls = export_df_to_excel({chk["Check"][:31]: df[disp]})
                st.download_button(
                    f"📥 Export: {chk['Check']}", xls,
                    f"health_{chk['Check'][:20].replace(' ','_')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.success("✅ No issues found for this check.")

    # Full export
    all_export = {chk["Check"][:31]: chk["df"][[c for c in ["task_code","task_name","total_float_days","status"] if c in chk["df"].columns]] if not chk["df"].empty else pd.DataFrame(columns=["No issues"]) for chk in checks}
    xls_all = export_df_to_excel(all_export)
    st.download_button("📥 Export Full Health Check Report", xls_all, "schedule_health_check.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# PAGE: PLANNING NOTES
# -----------------------------------------------------------------------------

HIGHLIGHT_WORDS = [
    "risk", "delay", "delayed", "blocked", "constraint", "access",
    "design", "procurement", "client", "instruction", "CE", "EWN",
    "change", "issue", "hold", "pending", "late", "overrun",
]

def highlight_text(text: str) -> str:
    """Wrap highlight words in HTML span."""
    for word in HIGHLIGHT_WORDS:
        pattern = re.compile(r"\b(" + re.escape(word) + r")\b", re.IGNORECASE)
        text = pattern.sub(r'<span style="background:#fef08a;font-weight:bold;">\1</span>', text)
    return text


def page_planning_notes(data: dict):
    st.title("📝 Planning Notes")
    st.markdown("> Upload planning notes and link them to activities in the programme.")

    tasks = data["tasks_df"]
    notes_file = st.file_uploader("Upload Planning Notes (CSV, Excel, TXT, or DOCX)",
                                   type=["csv","xlsx","txt","docx"])

    if notes_file is None:
        st.info("Upload a notes file to get started. The file should contain free-text notes referencing activity IDs.")
        return

    # Read notes
    notes_text = ""
    notes_rows = []

    try:
        if notes_file.name.endswith(".csv"):
            df = pd.read_csv(notes_file)
            notes_text = " ".join(df.astype(str).values.flatten())
            notes_rows = df.to_dict("records")
        elif notes_file.name.endswith(".xlsx"):
            df = pd.read_excel(notes_file)
            notes_text = " ".join(df.astype(str).values.flatten())
            notes_rows = df.to_dict("records")
        elif notes_file.name.endswith(".txt"):
            notes_text = notes_file.read().decode("utf-8", errors="replace")
            notes_rows = [{"line": i+1, "text": line} for i, line in enumerate(notes_text.splitlines()) if line.strip()]
        elif notes_file.name.endswith(".docx"):
            from docx import Document
            doc = Document(io.BytesIO(notes_file.read()))
            lines = [p.text for p in doc.paragraphs if p.text.strip()]
            notes_text = "\n".join(lines)
            notes_rows = [{"paragraph": i+1, "text": line} for i, line in enumerate(lines)]
        else:
            st.error("Unsupported file format.")
            return
        st.success(f"Loaded notes file: {notes_file.name}")
    except Exception as e:
        st.error(f"Could not read notes file: {e}")
        return

    # Find activity IDs mentioned in notes
    if not tasks.empty and "task_code" in tasks.columns:
        task_codes = tasks["task_code"].dropna().tolist()
        found_codes = [code for code in task_codes if code in notes_text]

        st.subheader(f"Activity IDs Found in Notes: {len(found_codes)}")
        if found_codes:
            matched_tasks = tasks[tasks["task_code"].isin(found_codes)][
                ["task_code","task_name","eff_start","eff_finish","total_float_days","status"]
            ]
            st.dataframe(matched_tasks, use_container_width=True)
        else:
            st.info("No activity IDs from the programme were found in the notes.")

        # Not found
        not_found = [code for code in task_codes if code not in notes_text]
        st.caption(f"{len(not_found)} activities not mentioned in notes.")

    # Keyword search
    st.divider()
    st.subheader("Keyword Search")
    keyword = st.text_input("Search notes for keyword")

    display_rows = notes_rows
    if keyword:
        display_rows = [r for r in notes_rows if keyword.lower() in str(r).lower()]
        st.caption(f"{len(display_rows)} matching entries")

    # Display with highlights
    for row in display_rows[:100]:
        text = str(row.get("text","") or list(row.values())[-1])
        highlighted = highlight_text(text)
        st.markdown(f"<div style='background:#f8fafc;border-left:3px solid #2563eb;padding:8px;margin:4px 0;font-size:13px;'>{highlighted}</div>", unsafe_allow_html=True)

    # Full highlighted dump
    st.divider()
    st.subheader("Full Notes (with keyword highlighting)")
    highlighted_full = highlight_text(notes_text.replace("\n","<br>"))
    st.markdown(f"<div style='background:white;border:1px solid #e2e8f0;padding:16px;border-radius:8px;max-height:400px;overflow-y:auto;font-size:12px;'>{highlighted_full}</div>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# PAGE: PROGRAMME COMPARISON / MOVEMENT INTELLIGENCE
# -----------------------------------------------------------------------------

def _mi_card(title: str, value, subtitle: str = "", colour: str = "#0B1F33",
             bg: str = "#ffffff", border: str = "#E2E8F0") -> str:
    """Render a compact metric card as HTML."""
    return (
        f'<div style="background:{bg};border:1px solid {border};border-radius:10px;'
        f'padding:16px 18px;box-shadow:0 1px 4px rgba(11,31,51,0.07);">'
        f'<div style="font-size:10px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
        f'text-transform:uppercase;margin-bottom:6px;">{title}</div>'
        f'<div style="font-size:26px;font-weight:800;color:{colour};line-height:1;">{value}</div>'
        f'{"" if not subtitle else f"<div style=font-size:11px;color:#64748B;margin-top:4px;>{subtitle}</div>"}'
        f'</div>'
    )


def _commentary_box(heading: str, body: str, colour: str = "#0B1F33",
                    bg: str = "#eff6ff", border: str = "#3b82f6") -> str:
    return (
        f'<div style="background:{bg};border-left:4px solid {border};border-radius:6px;'
        f'padding:14px 18px;margin-bottom:10px;">'
        f'<div style="font-weight:700;color:{colour};margin-bottom:4px;">{heading}</div>'
        f'<div style="font-size:14px;color:#334155;line-height:1.6;">{body}</div>'
        f'</div>'
    )


def _compute_movement(prev_tasks: pd.DataFrame, curr_tasks: pd.DataFrame,
                      near_crit: float = 10.0) -> dict:
    """
    Merge previous and current task sets and compute all movement metrics.
    Returns a dict of DataFrames and scalar stats.
    """
    prev = get_critical_threshold(prev_tasks.copy(), near_crit)
    curr = get_critical_threshold(curr_tasks.copy(), near_crit)

    # Common activities (matched on task_code)
    merged = prev.merge(curr, on="task_code", how="outer", suffixes=("_p", "_c"))

    added   = curr[~curr["task_code"].isin(prev["task_code"])].copy()
    removed = prev[~prev["task_code"].isin(curr["task_code"])].copy()
    common  = merged.dropna(subset=["task_code"]).copy()

    # ---- Date movement -------------------------------------------------------
    def _days_diff(a, b):
        """b minus a in calendar days. Positive = slipped."""
        try:
            if pd.isna(a) or pd.isna(b):
                return None
            return int((pd.Timestamp(b) - pd.Timestamp(a)).days)
        except Exception:
            return None

    common["finish_move"] = common.apply(
        lambda r: _days_diff(r.get("eff_finish_p"), r.get("eff_finish_c")), axis=1
    )
    common["start_move"] = common.apply(
        lambda r: _days_diff(r.get("eff_start_p"), r.get("eff_start_c")), axis=1
    )

    # ---- Float movement ------------------------------------------------------
    common["float_move"] = common.apply(
        lambda r: (
            safe_float(r.get("total_float_days_c"), None) is not None and
            safe_float(r.get("total_float_days_p"), None) is not None
        ) and safe_float(r.get("total_float_days_c"), 0) - safe_float(r.get("total_float_days_p"), 0)
        if (
            safe_float(r.get("total_float_days_c"), None) is not None and
            safe_float(r.get("total_float_days_p"), None) is not None
        ) else None,
        axis=1,
    )

    # ---- Critical transitions ------------------------------------------------
    def _is_crit(row, suffix):
        tf = safe_float(row.get(f"total_float_days_{suffix}"), 9999)
        return tf <= 0

    common["was_crit"]  = common.apply(lambda r: _is_crit(r, "p"), axis=1)
    common["now_crit"]  = common.apply(lambda r: _is_crit(r, "c"), axis=1)
    became_crit  = common[~common["was_crit"] & common["now_crit"]]
    no_longer_crit = common[common["was_crit"] & ~common["now_crit"]]

    # ---- Slipped / improved --------------------------------------------------
    delayed  = common[common["finish_move"].apply(lambda x: x is not None and x > 0)]
    improved = common[common["finish_move"].apply(lambda x: x is not None and x < 0)]
    float_lost   = common[common["float_move"].apply(lambda x: x is not None and x < 0)]
    float_gained = common[common["float_move"].apply(lambda x: x is not None and x > 0)]

    # ---- Project finish movement --------------------------------------------
    prev_end = prev["eff_finish"].dropna().max() if "eff_finish" in prev.columns else None
    curr_end = curr["eff_finish"].dropna().max() if "eff_finish" in curr.columns else None
    proj_finish_move = _days_diff(prev_end, curr_end)

    # ---- WBS-level summary --------------------------------------------------
    wbs_col_p = "wbs_path_p" if "wbs_path_p" in common.columns else None
    wbs_col_c = "wbs_path_c" if "wbs_path_c" in common.columns else None
    wbs_col   = wbs_col_c or wbs_col_p

    if wbs_col and not common.empty:
        common["wbs_top"] = common[wbs_col].apply(
            lambda x: str(x).split(" > ")[0] if pd.notna(x) and str(x).strip() not in ("","nan") else "Unknown"
        )
        wbs_summary = common.groupby("wbs_top").agg(
            total_activities=("task_code", "count"),
            delayed_count=("finish_move", lambda x: (x > 0).sum()),
            improved_count=("finish_move", lambda x: (x < 0).sum()),
            avg_finish_move=("finish_move", lambda x: round(x.dropna().mean(), 1) if x.dropna().any() else 0),
            max_finish_slip=("finish_move", lambda x: x.dropna().max() if not x.dropna().empty else 0),
            crit_gained=("now_crit", lambda x: (x & ~common.loc[x.index, "was_crit"]).sum()),
            crit_lost=("now_crit", lambda x: (~x & common.loc[x.index, "was_crit"]).sum()),
        ).reset_index()
    else:
        wbs_summary = pd.DataFrame()

    return {
        "merged":        common,
        "added":         added,
        "removed":       removed,
        "delayed":       delayed.sort_values("finish_move", ascending=False),
        "improved":      improved.sort_values("finish_move"),
        "float_lost":    float_lost.sort_values("float_move"),
        "float_gained":  float_gained.sort_values("float_move", ascending=False),
        "became_crit":   became_crit,
        "no_longer_crit":no_longer_crit,
        "wbs_summary":   wbs_summary,
        "proj_finish_move": proj_finish_move,
        "prev_end":      prev_end,
        "curr_end":      curr_end,
        "n_added":       len(added),
        "n_removed":     len(removed),
        "n_delayed":     len(delayed),
        "n_improved":    len(improved),
        "n_became_crit": len(became_crit),
        "n_no_longer_crit": len(no_longer_crit),
        "n_float_lost":  len(float_lost),
        "n_float_gained":len(float_gained),
        "n_common":      len(common),
    }


def _build_commentary(m: dict, prev_name: str, curr_name: str) -> list:
    """
    Generate plain-English commentary bullets from movement data.
    Returns list of (heading, body, colour, bg, border) tuples.
    """
    items = []
    pfm = m["proj_finish_move"]

    # -- What changed? ---------------------------------------------------------
    changes = []
    if pfm is not None and pfm > 0:
        changes.append(f"The project finish date has slipped by <strong>{pfm} days</strong>.")
    elif pfm is not None and pfm < 0:
        changes.append(f"The project finish date has improved by <strong>{abs(pfm)} days</strong>.")
    elif pfm == 0:
        changes.append("The project finish date is unchanged between revisions.")

    if m["n_delayed"] > 0:
        changes.append(
            f"<strong>{m['n_delayed']}</strong> activities have a later finish date in the current programme."
        )
    if m["n_improved"] > 0:
        changes.append(f"<strong>{m['n_improved']}</strong> activities have an earlier finish date.")
    if m["n_added"] > 0:
        changes.append(f"<strong>{m['n_added']}</strong> new activities have been added.")
    if m["n_removed"] > 0:
        changes.append(f"<strong>{m['n_removed']}</strong> activities have been removed.")

    if changes:
        items.append((
            "What changed?",
            " ".join(changes),
            "#1e3a5f", "#eff6ff", "#3b82f6"
        ))

    # -- Where is the risk? ----------------------------------------------------
    risks = []
    if m["n_became_crit"] > 0:
        risks.append(
            f"<strong>{m['n_became_crit']} activities became critical</strong> in the current revision. "
            "These activities now have zero or negative float and must be monitored closely."
        )
    neg_float_curr = m["merged"]["total_float_days_c"].apply(
        lambda f: safe_float(f, 0) < 0
    ).sum() if "total_float_days_c" in m["merged"].columns else 0
    if neg_float_curr > 0:
        risks.append(
            f"<strong>{int(neg_float_curr)} activities have negative float</strong> in the current programme. "
            "This means the schedule cannot currently achieve its target dates on these paths."
        )
    if m["n_float_lost"] > 0:
        worst_loss = m["float_lost"]["float_move"].min()
        risks.append(
            f"<strong>{m['n_float_lost']} activities lost float</strong>. "
            f"The worst case is a loss of {abs(worst_loss):.1f} days."
        )
    if not m["delayed"].empty:
        worst_slip = m["delayed"]["finish_move"].max()
        worst_act  = m["delayed"].iloc[0].get("task_name_c", m["delayed"].iloc[0].get("task_code",""))
        risks.append(
            f"The biggest finish date slip is <strong>{int(worst_slip)} days</strong> "
            f"({str(worst_act)[:60]})."
        )

    if risks:
        items.append((
            "Where is the risk?",
            " ".join(risks),
            "#7f1d1d", "#fef2f2", "#dc2626"
        ))
    else:
        items.append((
            "Where is the risk?",
            "No significant new risks identified between these two revisions.",
            "#166534", "#f0fdf4", "#16a34a"
        ))

    # -- What needs PM attention? ----------------------------------------------
    attention = []
    if m["n_became_crit"] > 0:
        codes = m["became_crit"]["task_code"].head(3).tolist()
        attention.append(
            f"Review the {m['n_became_crit']} newly critical activities, "
            f"including: {', '.join(str(c) for c in codes)}."
        )
    if neg_float_curr > 0:
        attention.append(
            "Investigate all activities with negative float. "
            "These require a recovery plan or a revised target date."
        )
    if not m["delayed"].empty:
        attention.append(
            "Review the biggest finish date slips in the Biggest Slips section below "
            "and confirm whether recovery actions are in place."
        )
    if m["n_removed"] > 0:
        attention.append(
            f"{m['n_removed']} activities were removed. "
            "Confirm these were intentional scope reductions and not accidental deletions."
        )

    if attention:
        items.append((
            "What needs PM attention?",
            " ".join(attention),
            "#92400e", "#fffbeb", "#F5A623"
        ))

    # -- What improved? --------------------------------------------------------
    good = []
    if m["n_no_longer_crit"] > 0:
        good.append(
            f"<strong>{m['n_no_longer_crit']} activities are no longer critical</strong>, "
            "indicating improved schedule recovery on those paths."
        )
    if m["n_improved"] > 0:
        best = abs(m["improved"]["finish_move"].min())
        good.append(
            f"<strong>{m['n_improved']} activities finished earlier</strong> than the previous revision. "
            f"The best improvement is {int(best)} days."
        )
    if m["n_float_gained"] > 0:
        good.append(f"<strong>{m['n_float_gained']} activities gained float</strong>.")
    if pfm is not None and pfm < 0:
        good.append(
            f"The overall project finish has <strong>improved by {abs(pfm)} days</strong>."
        )

    if good:
        items.append((
            "What improved?",
            " ".join(good),
            "#166534", "#f0fdf4", "#16a34a"
        ))

    return items


def _display_cols(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """Return a display-ready copy with renamed columns and formatted dates."""
    avail = {k: v for k, v in col_map.items() if k in df.columns}
    out = df[list(avail.keys())].copy().rename(columns=avail)
    for col in out.columns:
        if any(kw in col.lower() for kw in ("start","finish","date")):
            try:
                out[col] = out[col].apply(format_date)
            except Exception:
                pass
    return out


def page_programme_comparison():
    """
    Programme Movement Intelligence page.
    Upload two XER files (previous and current) and get a plain-English
    analysis of what changed, where the risk is, and what improved.
    """
    st.title("📅 Movement Intelligence")
    st.caption(
        "Upload the previous and current version of your XER programme to see "
        "a plain-English analysis of what has changed, where the risk is, "
        "and what needs your attention."
    )

    # ---- File uploads --------------------------------------------------------
    up_col1, up_col2 = st.columns(2)
    with up_col1:
        st.markdown(
            '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
            'letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">'
            'Previous Programme</div>',
            unsafe_allow_html=True,
        )
        prev_file = st.file_uploader("Previous XER", type=["xer"], key="mi_prev_xer",
                                     label_visibility="collapsed")
    with up_col2:
        st.markdown(
            '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
            'letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">'
            'Current Programme</div>',
            unsafe_allow_html=True,
        )
        curr_file = st.file_uploader("Current XER", type=["xer"], key="mi_curr_xer",
                                     label_visibility="collapsed")

    if not prev_file or not curr_file:
        st.markdown(
            '<div style="background:#f8fafc;border:2px dashed #CBD5E1;border-radius:10px;'
            'padding:32px;text-align:center;color:#64748B;margin-top:20px;">'
            '<div style="font-size:28px;margin-bottom:10px;">📅</div>'
            '<strong style="font-size:15px;">Upload both programmes above to begin</strong><br>'
            '<span style="font-size:13px;margin-top:6px;display:block;">'
            'Upload the previous revision on the left and the current revision on the right.</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    # ---- Parse & cache -------------------------------------------------------
    cache_key = f"mi_{prev_file.name}_{prev_file.size}_{curr_file.name}_{curr_file.size}"
    if st.session_state.get("_mi_cache_key") != cache_key:
        with st.spinner("Parsing both programmes..."):
            try:
                prev_data = parse_xer(prev_file.read())
                curr_data = parse_xer(curr_file.read())
            except Exception as e:
                st.error(f"Could not parse XER files: {e}")
                return
        st.session_state["_mi_prev"] = prev_data
        st.session_state["_mi_curr"] = curr_data
        st.session_state["_mi_cache_key"] = cache_key

    prev_data = st.session_state["_mi_prev"]
    curr_data = st.session_state["_mi_curr"]

    prev_tasks = prev_data["tasks_df"]
    curr_tasks = curr_data["tasks_df"]

    if prev_tasks.empty or curr_tasks.empty:
        st.error("Could not extract activities from one or both files. Check the XER exports.")
        return

    near_crit_days = 10.0

    # ---- Run movement analysis -----------------------------------------------
    m = _compute_movement(prev_tasks, curr_tasks, near_crit_days)

    prev_name = prev_data.get("project_info", {}).get("name", prev_file.name)
    curr_name = curr_data.get("project_info", {}).get("name", curr_file.name)

    # ---- File header strip ---------------------------------------------------
    pfm = m["proj_finish_move"]
    pfm_str = (
        f"+{pfm}d (slipped)" if pfm and pfm > 0 else
        f"{pfm}d (improved)" if pfm and pfm < 0 else
        "No change" if pfm == 0 else "N/A"
    )
    pfm_colour = "#dc2626" if pfm and pfm > 0 else "#16a34a" if pfm and pfm < 0 else "#6b7280"

    st.markdown(
        f"""
        <div style="background:#0B1F33;border-radius:12px;padding:20px 24px;
                    margin-bottom:24px;display:flex;gap:32px;flex-wrap:wrap;
                    align-items:center;">
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Previous</div>
                <div style="font-size:14px;font-weight:600;color:#CBD5E1;margin-top:2px;">
                    {prev_name}</div>
                <div style="font-size:11px;color:#475569;">
                    {format_date(m["prev_end"])}</div>
            </div>
            <div style="font-size:24px;color:#F5A623;">&#8594;</div>
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Current</div>
                <div style="font-size:14px;font-weight:600;color:#CBD5E1;margin-top:2px;">
                    {curr_name}</div>
                <div style="font-size:11px;color:#475569;">
                    {format_date(m["curr_end"])}</div>
            </div>
            <div style="margin-left:auto;text-align:right;">
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Net Project Finish Movement</div>
                <div style="font-size:28px;font-weight:800;color:{pfm_colour};margin-top:2px;">
                    {pfm_str}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Summary metric cards ------------------------------------------------
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;">'
        'Summary</div>',
        unsafe_allow_html=True,
    )

    r1 = st.columns(5)
    r1[0].markdown(_mi_card("Activities Added",    m["n_added"],    colour="#16a34a"), unsafe_allow_html=True)
    r1[1].markdown(_mi_card("Activities Removed",  m["n_removed"],  colour="#6b7280"), unsafe_allow_html=True)
    r1[2].markdown(_mi_card("Finish Date Moved",   m["n_delayed"] + m["n_improved"], colour="#0B1F33"), unsafe_allow_html=True)
    r1[3].markdown(_mi_card("Delayed",             m["n_delayed"],  colour="#dc2626"), unsafe_allow_html=True)
    r1[4].markdown(_mi_card("Improved",            m["n_improved"], colour="#16a34a"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    r2 = st.columns(5)
    r2[0].markdown(_mi_card("Became Critical",     m["n_became_crit"],    colour="#dc2626", bg="#fef2f2", border="#fca5a5"), unsafe_allow_html=True)
    r2[1].markdown(_mi_card("No Longer Critical",  m["n_no_longer_crit"], colour="#16a34a", bg="#f0fdf4", border="#86efac"), unsafe_allow_html=True)
    r2[2].markdown(_mi_card("Lost Float",          m["n_float_lost"],     colour="#d97706", bg="#fffbeb", border="#fcd34d"), unsafe_allow_html=True)
    r2[3].markdown(_mi_card("Gained Float",        m["n_float_gained"],   colour="#16a34a", bg="#f0fdf4", border="#86efac"), unsafe_allow_html=True)
    r2[4].markdown(_mi_card("Net Project Movement", pfm_str,              colour=pfm_colour), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- Written commentary --------------------------------------------------
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;">'
        'Intelligence Summary</div>',
        unsafe_allow_html=True,
    )

    commentary = _build_commentary(m, prev_name, curr_name)
    for heading, body, col, bg, border in commentary:
        st.markdown(_commentary_box(heading, body, col, bg, border), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- Detailed sections (tabs) --------------------------------------------
    (tab_slips, tab_gains, tab_float_loss, tab_float_gain,
     tab_new_crit, tab_rem_crit, tab_added, tab_removed,
     tab_wbs, tab_export) = st.tabs([
        "Biggest Slips",
        "Biggest Gains",
        "Float Losses",
        "Float Gains",
        "New Critical",
        "No Longer Critical",
        "Added",
        "Removed",
        "WBS Movement",
        "Export",
    ])

    # Helper: common columns for merged common activities
    MOVE_COLS = {
        "task_code":           "Activity ID",
        "task_name_c":         "Activity Name",
        "wbs_path_c":          "WBS",
        "eff_finish_p":        "Previous Finish",
        "eff_finish_c":        "Current Finish",
        "finish_move":         "Finish Movement (d)",
        "total_float_days_p":  "Prev Float (d)",
        "total_float_days_c":  "Curr Float (d)",
        "float_move":          "Float Movement (d)",
    }
    ALT_NAME = {k.replace("_c", "_p"): v for k, v in MOVE_COLS.items()}

    def _safe_display(df, col_map=None):
        if df is None or df.empty:
            return pd.DataFrame(columns=["No data"])
        cols = col_map or MOVE_COLS
        avail = {k: v for k, v in cols.items() if k in df.columns}
        # Fallbacks for name column
        if "task_name_c" not in df.columns and "task_name_p" in df.columns:
            avail["task_name_p"] = "Activity Name"
        out = df[list(avail.keys())].copy().rename(columns=avail)
        for col in ["Previous Finish","Current Finish","Previous Start","Current Start"]:
            if col in out.columns:
                out[col] = out[col].apply(format_date)
        for col in ["Prev Float (d)","Curr Float (d)","Float Movement (d)","Finish Movement (d)"]:
            if col in out.columns:
                out[col] = out[col].apply(lambda x: round(float(x),1) if x is not None and str(x) not in ("","nan") else "-")
        return out

    # -- Biggest slips ---------------------------------------------------------
    with tab_slips:
        st.markdown("**Activities with the largest finish date slips** (most slipped first).")
        df_slips = _safe_display(m["delayed"].head(50))
        if df_slips.empty or "No data" in df_slips.columns:
            st.success("No activities slipped between these two revisions.")
        else:
            # Colour the movement column
            def _red_pos(val):
                try:
                    return "color:#dc2626;font-weight:600;" if float(val) > 0 else ""
                except Exception:
                    return ""
            if "Finish Movement (d)" in df_slips.columns:
                st.dataframe(
                    df_slips.style.applymap(_red_pos, subset=["Finish Movement (d)"]),
                    use_container_width=True, hide_index=True
                )
            else:
                st.dataframe(df_slips, use_container_width=True, hide_index=True)

            if not m["delayed"].empty:
                top20 = m["delayed"].head(20).copy()
                top20 = top20.dropna(subset=["finish_move"])
                label_col = "task_name_c" if "task_name_c" in top20.columns else "task_code"
                top20["Label"] = top20["task_code"].astype(str) + "  " + top20[label_col].astype(str).str[:35]
                fig = px.bar(
                    top20, x="finish_move", y="Label", orientation="h",
                    title="Top 20 Biggest Finish Date Slips (days)",
                    labels={"finish_move":"Slip (days)","Label":""},
                    color_discrete_sequence=["#dc2626"],
                )
                fig.update_yaxes(autorange="reversed")
                fig.update_layout(height=500, margin=dict(l=10,r=10,t=40,b=10),
                                  plot_bgcolor="#ffffff", paper_bgcolor="#ffffff")
                st.plotly_chart(fig, use_container_width=True)

    # -- Biggest gains ---------------------------------------------------------
    with tab_gains:
        st.markdown("**Activities with the largest finish date improvements** (most improved first).")
        df_gains = _safe_display(m["improved"].head(50))
        if df_gains.empty or "No data" in df_gains.columns:
            st.success("No activities improved their finish date between these two revisions.")
        else:
            st.dataframe(df_gains, use_container_width=True, hide_index=True)
            if not m["improved"].empty:
                top20g = m["improved"].head(20).copy().dropna(subset=["finish_move"])
                label_col = "task_name_c" if "task_name_c" in top20g.columns else "task_code"
                top20g["Label"] = top20g["task_code"].astype(str) + "  " + top20g[label_col].astype(str).str[:35]
                top20g["improvement"] = top20g["finish_move"].abs()
                fig2 = px.bar(
                    top20g, x="improvement", y="Label", orientation="h",
                    title="Top 20 Biggest Finish Date Improvements (days)",
                    labels={"improvement":"Improvement (days)","Label":""},
                    color_discrete_sequence=["#16a34a"],
                )
                fig2.update_yaxes(autorange="reversed")
                fig2.update_layout(height=500, margin=dict(l=10,r=10,t=40,b=10),
                                   plot_bgcolor="#ffffff", paper_bgcolor="#ffffff")
                st.plotly_chart(fig2, use_container_width=True)

    # -- Float losses ----------------------------------------------------------
    with tab_float_loss:
        st.markdown("**Activities that lost the most float** -- these are moving towards the critical path.")
        df_fl = _safe_display(m["float_lost"].head(50))
        if df_fl.empty or "No data" in df_fl.columns:
            st.success("No activities lost float between these two revisions.")
        else:
            st.dataframe(df_fl, use_container_width=True, hide_index=True)

    # -- Float gains -----------------------------------------------------------
    with tab_float_gain:
        st.markdown("**Activities that gained the most float** -- these have moved away from the critical path.")
        df_fg = _safe_display(m["float_gained"].head(50))
        if df_fg.empty or "No data" in df_fg.columns:
            st.success("No activities gained float between these two revisions.")
        else:
            st.dataframe(df_fg, use_container_width=True, hide_index=True)

    # -- New critical activities ------------------------------------------------
    with tab_new_crit:
        if m["became_crit"].empty:
            st.success("No activities became critical in the current revision. Good news.")
        else:
            st.warning(
                f"**{len(m['became_crit'])} activities became critical** in the current revision. "
                "These activities now have zero or negative float and require immediate attention."
            )
            df_nc = _safe_display(m["became_crit"])
            st.dataframe(df_nc, use_container_width=True, hide_index=True)

    # -- No longer critical ----------------------------------------------------
    with tab_rem_crit:
        if m["no_longer_crit"].empty:
            st.info("No activities moved off the critical path in the current revision.")
        else:
            st.success(
                f"**{len(m['no_longer_crit'])} activities are no longer critical.** "
                "This indicates recovery on those paths."
            )
            df_rc = _safe_display(m["no_longer_crit"])
            st.dataframe(df_rc, use_container_width=True, hide_index=True)

    # -- Added activities ------------------------------------------------------
    with tab_added:
        if m["added"].empty:
            st.info("No new activities were added in the current revision.")
        else:
            st.info(f"**{len(m['added'])} activities added** to the current programme.")
            added_cols = {
                "task_code":"Activity ID","task_name":"Activity Name",
                "wbs_path":"WBS","eff_start":"Start","eff_finish":"Finish",
                "total_float_days":"Float (d)","status":"Status",
            }
            st.dataframe(_safe_display(m["added"], added_cols), use_container_width=True, hide_index=True)

    # -- Removed activities ----------------------------------------------------
    with tab_removed:
        if m["removed"].empty:
            st.info("No activities were removed in the current revision.")
        else:
            st.warning(
                f"**{len(m['removed'])} activities removed.** "
                "Confirm these are intentional and not accidental deletions."
            )
            rem_cols = {
                "task_code":"Activity ID","task_name":"Activity Name",
                "wbs_path":"WBS","eff_start":"Start","eff_finish":"Finish",
                "total_float_days":"Float (d)","status":"Status",
            }
            st.dataframe(_safe_display(m["removed"], rem_cols), use_container_width=True, hide_index=True)

    # -- WBS-level movement ----------------------------------------------------
    with tab_wbs:
        st.markdown("**Finish date and critical path movement broken down by WBS area.**")
        if m["wbs_summary"].empty:
            st.info("WBS data not available for movement analysis.")
        else:
            wbs_disp = m["wbs_summary"].copy()
            wbs_disp = wbs_disp.rename(columns={
                "wbs_top":           "WBS Area",
                "total_activities":  "Total Activities",
                "delayed_count":     "Delayed",
                "improved_count":    "Improved",
                "avg_finish_move":   "Avg Finish Move (d)",
                "max_finish_slip":   "Max Slip (d)",
                "crit_gained":       "Became Critical",
                "crit_lost":         "No Longer Critical",
            })
            wbs_disp = wbs_disp.sort_values("Max Slip (d)", ascending=False)
            st.dataframe(wbs_disp, use_container_width=True, hide_index=True)

            # Chart: delayed count by WBS
            if not wbs_disp.empty and "Delayed" in wbs_disp.columns:
                fig_wbs = px.bar(
                    wbs_disp[wbs_disp["Delayed"] > 0].head(20),
                    x="Delayed", y="WBS Area", orientation="h",
                    title="Delayed Activities by WBS Area",
                    labels={"Delayed":"Activities Delayed","WBS Area":""},
                    color_discrete_sequence=["#dc2626"],
                )
                fig_wbs.update_yaxes(autorange="reversed")
                fig_wbs.update_layout(height=400, margin=dict(l=10,r=10,t=40,b=10),
                                      plot_bgcolor="#ffffff", paper_bgcolor="#ffffff")
                st.plotly_chart(fig_wbs, use_container_width=True)

            # Chart: average movement by WBS
            wbs_nonzero = wbs_disp[wbs_disp["Avg Finish Move (d)"] != 0].head(20)
            if not wbs_nonzero.empty:
                fig_avg = px.bar(
                    wbs_nonzero,
                    x="Avg Finish Move (d)", y="WBS Area", orientation="h",
                    title="Average Finish Movement by WBS (positive = slipped)",
                    labels={"Avg Finish Move (d)":"Average (days)","WBS Area":""},
                    color="Avg Finish Move (d)",
                    color_continuous_scale=["#16a34a","#f5f5f5","#dc2626"],
                    color_continuous_midpoint=0,
                )
                fig_avg.update_yaxes(autorange="reversed")
                fig_avg.update_layout(height=400, margin=dict(l=10,r=10,t=40,b=10),
                                      plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
                                      coloraxis_showscale=False)
                st.plotly_chart(fig_avg, use_container_width=True)

    # -- Export ----------------------------------------------------------------
    with tab_export:
        st.markdown("**Download the full movement analysis as a formatted Excel workbook.**")

        summary_sheet = pd.DataFrame({
            "Metric": [
                "Previous Programme", "Current Programme",
                "Previous Finish", "Current Finish",
                "Net Project Finish Movement",
                "Activities Added", "Activities Removed",
                "Activities Delayed", "Activities Improved",
                "Became Critical", "No Longer Critical",
                "Lost Float", "Gained Float",
            ],
            "Value": [
                prev_name, curr_name,
                format_date(m["prev_end"]), format_date(m["curr_end"]),
                pfm_str,
                m["n_added"], m["n_removed"],
                m["n_delayed"], m["n_improved"],
                m["n_became_crit"], m["n_no_longer_crit"],
                m["n_float_lost"], m["n_float_gained"],
            ],
        })

        commentary_sheet = pd.DataFrame({
            "Section": [h for h, *_ in commentary],
            "Commentary": [b.replace("<strong>","").replace("</strong>","") for _,b,*_ in commentary],
        })

        export_sheets = {
            "Summary":          summary_sheet,
            "Commentary":       commentary_sheet,
            "Biggest Slips":    _safe_display(m["delayed"].head(100)),
            "Biggest Gains":    _safe_display(m["improved"].head(100)),
            "Float Losses":     _safe_display(m["float_lost"].head(100)),
            "Float Gains":      _safe_display(m["float_gained"].head(100)),
            "Became Critical":  _safe_display(m["became_crit"]),
            "No Longer Critical": _safe_display(m["no_longer_crit"]),
            "Added Activities": _safe_display(
                m["added"],
                {"task_code":"Activity ID","task_name":"Activity Name",
                 "wbs_path":"WBS","eff_start":"Start","eff_finish":"Finish",
                 "total_float_days":"Float (d)","status":"Status"},
            ),
            "Removed Activities": _safe_display(
                m["removed"],
                {"task_code":"Activity ID","task_name":"Activity Name",
                 "wbs_path":"WBS","eff_start":"Start","eff_finish":"Finish",
                 "total_float_days":"Float (d)","status":"Status"},
            ),
        }
        if not m["wbs_summary"].empty:
            export_sheets["WBS Movement"] = m["wbs_summary"].rename(columns={
                "wbs_top":"WBS Area","total_activities":"Total Activities",
                "delayed_count":"Delayed","improved_count":"Improved",
                "avg_finish_move":"Avg Finish Move (d)","max_finish_slip":"Max Slip (d)",
                "crit_gained":"Became Critical","crit_lost":"No Longer Critical",
            })

        xls_bytes = export_df_to_excel(export_sheets)

        st.download_button(
            label="📥  Download Movement Intelligence Report",
            data=xls_bytes,
            file_name=f"movement_intelligence_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Exports all movement data across multiple sheets including commentary.",
        )

        st.markdown(
            f"""
            <div style="background:#f8fafc;border:1px solid #E2E8F0;border-radius:8px;
                        padding:14px 18px;margin-top:12px;">
                <div style="font-size:12px;font-weight:700;color:#0B1F33;
                            margin-bottom:8px;">Workbook contents</div>
                <div style="font-size:12px;color:#64748B;line-height:2;">
                    {"".join(f"<strong>{k}</strong> &nbsp;|&nbsp; " for k in export_sheets.keys())}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )




# -----------------------------------------------------------------------------
# PAGE: PM ACTIONS
# -----------------------------------------------------------------------------

# Risk keywords for notes scanning
_RISK_WORDS = [
    "risk","delay","delayed","delays","blocked","block","constraint","access",
    "design","procurement","client","instruction","CE","EWN","change",
    "hold","pending","late","overrun","issue","dispute","claim",
]

def _generate_actions(
    tasks: pd.DataFrame,
    rels: pd.DataFrame,
    near_crit_days: float,
    notes_text: str = "",
) -> pd.DataFrame:
    """
    Analyse the programme and generate a prioritised PM action list.
    Returns a DataFrame with one row per action.
    """
    rows = []
    now  = datetime.now()
    eight_weeks = now + timedelta(weeks=8)
    four_weeks  = now + timedelta(weeks=4)

    tasks = get_critical_threshold(tasks, near_crit_days)

    # Build predecessor/successor sets from rels
    tasks_with_pred = set()
    tasks_with_succ = set()
    if not rels.empty:
        if "succ_task_id" in rels.columns:
            tasks_with_pred = set(rels["succ_task_id"].dropna())
        if "pred_task_id" in rels.columns:
            tasks_with_succ = set(rels["pred_task_id"].dropna())

    def _wbs(row):
        wbs = row.get("wbs_path", "") if "wbs_path" in row.index else ""
        return str(wbs).split(" > ")[0] if wbs and str(wbs).strip() not in ("","nan") else "-"

    def _row(priority, category, task_code, task_name, wbs, issue, why, action):
        return {
            "Priority":         priority,
            "Category":         category,
            "Activity ID":      str(task_code),
            "Activity Name":    str(task_name),
            "WBS":              str(wbs),
            "Issue":            str(issue),
            "Why It Matters":   str(why),
            "Suggested Action": str(action),
            "Owner":            "",
            "Due Date":         "",
            "Status":           "Open",
        }

    # -- 1. Negative float -----------------------------------------------------
    neg = tasks[tasks["total_float_days"].apply(lambda f: safe_float(f,0) < 0)] \
        if "total_float_days" in tasks.columns else pd.DataFrame()
    for _, t in neg.iterrows():
        tf = round(safe_float(t.get("total_float_days"), 0), 1)
        rows.append(_row(
            "High", "Negative Float",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            f"Activity has {tf} days negative float.",
            "Negative float means the current schedule cannot meet its target date on this path. "
            "Every day of inaction increases the delay.",
            f"Investigate the cause of the {abs(tf)}-day overrun. "
            "Agree a recovery plan with the planner and update the programme.",
        ))

    # -- 2. Critical activities not started ------------------------------------
    crit_ns = tasks[
        tasks["is_critical"] &
        tasks["status"].apply(lambda s: str(s) in ("TK_NotStart","Not Started"))
    ] if "status" in tasks.columns and "is_critical" in tasks.columns else pd.DataFrame()
    for _, t in crit_ns.iterrows():
        finish = t.get("eff_finish")
        finish_str = format_date(finish) if finish else "unknown"
        rows.append(_row(
            "High", "Critical Not Started",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            f"Critical activity not yet started. Target finish: {finish_str}.",
            "Any delay to a critical activity directly delays the project finish date. "
            "There is no float to absorb slippage.",
            "Confirm start date with the responsible party. "
            "If start is at risk, escalate immediately and review recovery options.",
        ))

    # -- 3. Near-critical due within 4 weeks -----------------------------------
    nc_due = tasks[
        tasks["is_near_critical"] &
        tasks["eff_finish"].apply(
            lambda d: d is not None and hasattr(d,"date") and d <= four_weeks
        )
    ] if "eff_finish" in tasks.columns else pd.DataFrame()
    for _, t in nc_due.iterrows():
        tf = round(safe_float(t.get("total_float_days"),0), 1)
        rows.append(_row(
            "High", "Near-Critical Due Soon",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            f"Near-critical activity due within 4 weeks. Float: {tf} days.",
            "With only a small float buffer and a near-term finish, any disruption "
            "will push this activity onto the critical path.",
            "Confirm the activity is progressing on schedule. "
            "If at risk, treat as critical and raise with the team now.",
        ))

    # -- 4. Near-critical due within 4-8 weeks --------------------------------
    nc_due_med = tasks[
        tasks["is_near_critical"] &
        tasks["eff_finish"].apply(
            lambda d: d is not None and hasattr(d,"date") and four_weeks < d <= eight_weeks
        )
    ] if "eff_finish" in tasks.columns else pd.DataFrame()
    for _, t in nc_due_med.iterrows():
        tf = round(safe_float(t.get("total_float_days"),0), 1)
        rows.append(_row(
            "Medium", "Near-Critical Due Soon",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            f"Near-critical activity due in 4-8 weeks. Float: {tf} days.",
            "Limited float with an upcoming finish date. Monitor closely.",
            "Review progress and confirm no blockers. Add to weekly look-ahead.",
        ))

    # -- 5. No predecessor (open start) ----------------------------------------
    no_pred = tasks[~tasks["task_id"].isin(tasks_with_pred)] \
        if "task_id" in tasks.columns else pd.DataFrame()
    # Exclude milestones with no predecessors intentionally (start milestones)
    no_pred = no_pred[no_pred.get("task_type","").apply(
        lambda t: "Milestone" not in str(t) and "LOE" not in str(t)
    )] if "task_type" in no_pred.columns else no_pred
    for _, t in no_pred.head(20).iterrows():
        rows.append(_row(
            "Medium", "Open Logic - No Predecessor",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            "Activity has no predecessor. Logic is open at the start.",
            "Open-start activities are not driven by schedule logic. "
            "Float calculations may be unreliable for this activity.",
            "Review with the planner. Add a predecessor or confirm the open start is intentional. "
            "If intentional, add a Start On or After constraint.",
        ))

    # -- 6. No successor (open finish) -----------------------------------------
    no_succ = tasks[~tasks["task_id"].isin(tasks_with_succ)] \
        if "task_id" in tasks.columns else pd.DataFrame()
    no_succ = no_succ[no_succ.get("task_type","").apply(
        lambda t: "Finish Milestone" not in str(t) and "LOE" not in str(t)
    )] if "task_type" in no_succ.columns else no_succ
    for _, t in no_succ.head(20).iterrows():
        rows.append(_row(
            "Medium", "Open Logic - No Successor",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            "Activity has no successor. Logic is open at the finish.",
            "Open-finish activities may have artificially high float and can mask "
            "schedule risk downstream.",
            "Review with the planner. Add a successor or confirm the open finish "
            "is a deliberate programme end point.",
        ))

    # -- 7. Excessive lag (> 10 days) ------------------------------------------
    if not rels.empty and "lag_days" in rels.columns:
        big_lag = rels[rels["lag_days"].apply(lambda l: safe_float(l,0) > 10)]
        for _, r in big_lag.head(15).iterrows():
            code = r.get("succ_task_code", r.get("succ_task_id",""))
            name = r.get("succ_task_name","")
            lag  = safe_float(r.get("lag_days",0),0)
            pred = r.get("pred_task_code", r.get("pred_task_id",""))
            # Get WBS from tasks
            match = tasks[tasks["task_code"] == code]
            wbs   = _wbs(match.iloc[0]) if not match.empty else "-"
            rows.append(_row(
                "Medium", "Excessive Lag",
                code, name, wbs,
                f"Relationship from {pred} has {int(lag)} days lag.",
                "Excessive lag hides schedule risk. It should be replaced with properly "
                "sequenced activities so the plan reflects real work.",
                f"Challenge the {int(lag)}-day lag with the planner. "
                "Replace with an intermediate activity or reduce the lag to the minimum justified.",
            ))

    # -- 8. Long duration activities (> 60 working days) -----------------------
    long_dur = tasks[tasks["orig_dur_days"].apply(lambda d: safe_float(d,0) > 60)] \
        if "orig_dur_days" in tasks.columns else pd.DataFrame()
    for _, t in long_dur.head(15).iterrows():
        dur = int(safe_float(t.get("orig_dur_days",0),0))
        rows.append(_row(
            "Low", "Long Duration",
            t.get("task_code",""), t.get("task_name",""), _wbs(t),
            f"Activity duration is {dur} working days.",
            "Activities longer than 60 days are difficult to manage and monitor. "
            "Progress is hard to measure and problems can go undetected for weeks.",
            f"Review whether the {dur}-day activity can be broken into smaller "
            "work packages of 20-30 days. Discuss with the planner.",
        ))

    # -- 9. Activities flagged in planning notes --------------------------------
    if notes_text and "task_code" in tasks.columns:
        for word in _RISK_WORDS:
            pattern = re.compile(r"\b" + re.escape(word) + r"\b", re.IGNORECASE)
            if not pattern.search(notes_text):
                continue
            # Find activity IDs mentioned near this risk word in the notes
            for _, t in tasks.iterrows():
                code = str(t.get("task_code",""))
                if code and code in notes_text:
                    # Check the risk word appears within 200 chars of the activity ID
                    idx = notes_text.find(code)
                    snippet = notes_text[max(0,idx-200):idx+200]
                    if pattern.search(snippet):
                        rows.append(_row(
                            "High", "Planning Notes Risk",
                            code, t.get("task_name",""), _wbs(t),
                            f"Activity mentioned in planning notes alongside the word '{word}'.",
                            "Planning notes flag a potential issue against this activity that "
                            "may not be visible in the programme alone.",
                            f"Review the planning notes for {code} and confirm the '{word}' "
                            "item has been actioned or is being tracked.",
                        ))
                        break   # one action per risk word per activity

    # -- 10. Activities with constraints ---------------------------------------
    if "cstr_type" in tasks.columns:
        constrained = tasks[tasks["cstr_type"].apply(
            lambda x: bool(x) and str(x).strip() not in ("","None","nan")
        )]
        for _, t in constrained.head(20).iterrows():
            cstr = str(t.get("cstr_type",""))
            rows.append(_row(
                "Low", "Constraint",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                f"Activity has a constraint: {cstr}.",
                "Constraints override schedule logic and can cause artificial float "
                "or negative float. Each constraint should be justified.",
                f"Confirm the {cstr} constraint is still valid. "
                "If the constraint is no longer required, ask the planner to remove it.",
            ))

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    # Deduplicate: same Activity ID + Category
    df = df.drop_duplicates(subset=["Activity ID","Category"]).reset_index(drop=True)

    # Sort: High first, then Medium, then Low; within each by Category
    priority_order = {"High": 0, "Medium": 1, "Low": 2}
    df["_sort"] = df["Priority"].map(priority_order)
    df = df.sort_values(["_sort","Category","Activity ID"]).drop(columns=["_sort"]).reset_index(drop=True)

    return df


def page_pm_actions(data: dict, near_crit_days: float):
    """
    PM Action Dashboard.
    Automatically generates a prioritised action list from the uploaded programme.
    Each action has priority, category, issue, why it matters, suggested action,
    and editable owner / due date / status fields.
    """
    st.title("📋 PM Actions")
    st.caption(
        "Automatically generated from your programme. "
        "Each action is prioritised and includes a suggested next step. "
        "Update the Owner, Due Date and Status columns to track progress."
    )

    tasks = data["tasks_df"]
    rels  = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities found. Please upload a programme first.")
        return

    # Retrieve any planning notes text from session state (set by Planning Notes page)
    notes_text = st.session_state.get("_notes_text", "")

    # ---- Generate / cache actions -------------------------------------------
    prog_key = st.session_state.get("_xer_cache_key", "")
    cache_key = f"_pm_actions_{prog_key}_{near_crit_days}"

    if st.session_state.get("_pm_actions_key") != cache_key:
        with st.spinner("Analysing programme..."):
            actions_df = _generate_actions(tasks, rels, near_crit_days, notes_text)
        st.session_state["_pm_actions_df"]  = actions_df
        st.session_state["_pm_actions_key"] = cache_key
    else:
        actions_df = st.session_state["_pm_actions_df"]

    if actions_df.empty:
        st.success(
            "No actions generated. The programme has no obvious issues detected by "
            "the automated checks. Review the Health Check page for more detail."
        )
        return

    # ---- Summary banner ------------------------------------------------------
    n_high   = int((actions_df["Priority"] == "High").sum())
    n_med    = int((actions_df["Priority"] == "Medium").sum())
    n_low    = int((actions_df["Priority"] == "Low").sum())
    n_total  = len(actions_df)

    st.markdown(
        f"""
        <div style="background:#0B1F33;border-radius:12px;padding:20px 24px;
                    margin-bottom:20px;display:flex;gap:24px;flex-wrap:wrap;
                    align-items:center;">
            <div>
                <div style="font-size:11px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Total Actions</div>
                <div style="font-size:32px;font-weight:800;color:#F5A623;
                            line-height:1;margin-top:4px;">{n_total}</div>
            </div>
            <div style="width:1px;background:#1e3a5f;align-self:stretch;"></div>
            <div>
                <div style="font-size:11px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">High Priority</div>
                <div style="font-size:28px;font-weight:800;color:#dc2626;
                            line-height:1;margin-top:4px;">{n_high}</div>
            </div>
            <div>
                <div style="font-size:11px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Medium</div>
                <div style="font-size:28px;font-weight:800;color:#d97706;
                            line-height:1;margin-top:4px;">{n_med}</div>
            </div>
            <div>
                <div style="font-size:11px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Low</div>
                <div style="font-size:28px;font-weight:800;color:#64748B;
                            line-height:1;margin-top:4px;">{n_low}</div>
            </div>
            <div style="margin-left:auto;">
                <div style="font-size:11px;color:#475569;line-height:1.6;">
                    Generated from negative float, critical path,<br>
                    open logic, lag, duration and planning notes.
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Category breakdown cards -------------------------------------------
    cats = actions_df.groupby("Category").size().reset_index(name="n")
    cols = st.columns(min(len(cats), 5))
    cat_colours = {
        "Negative Float":           "#dc2626",
        "Critical Not Started":     "#b91c1c",
        "Near-Critical Due Soon":   "#d97706",
        "Open Logic - No Predecessor": "#7c3aed",
        "Open Logic - No Successor":   "#6d28d9",
        "Excessive Lag":            "#0369a1",
        "Long Duration":            "#0891b2",
        "Planning Notes Risk":      "#dc2626",
        "Constraint":               "#475569",
    }
    for i, (_, cat_row) in enumerate(cats.iterrows()):
        col = cols[i % len(cols)]
        colour = cat_colours.get(cat_row["Category"], "#374151")
        col.markdown(
            f'<div style="background:#ffffff;border:1px solid #E2E8F0;border-radius:8px;'
            f'padding:12px 14px;border-top:3px solid {colour};margin-bottom:6px;">'
            f'<div style="font-size:11px;color:#94A3B8;text-transform:uppercase;'
            f'letter-spacing:0.8px;margin-bottom:4px;">{cat_row["Category"]}</div>'
            f'<div style="font-size:22px;font-weight:800;color:{colour};">{cat_row["n"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- Filters ------------------------------------------------------------
    with st.expander("Filter actions", expanded=False):
        fc1, fc2, fc3, fc4, fc5 = st.columns(5)

        all_priorities = ["All"] + sorted(actions_df["Priority"].unique().tolist(),
                          key=lambda p: {"High":0,"Medium":1,"Low":2}.get(p,9))
        f_priority = fc1.selectbox("Priority",  all_priorities, key="pma_f_pri")

        all_cats = ["All"] + sorted(actions_df["Category"].unique().tolist())
        f_cat    = fc2.selectbox("Category",  all_cats, key="pma_f_cat")

        all_wbs = ["All"] + sorted(actions_df["WBS"].unique().tolist())
        f_wbs   = fc3.selectbox("WBS",        all_wbs, key="pma_f_wbs")

        all_owners = ["All"] + sorted([o for o in actions_df["Owner"].unique() if o])
        f_owner  = fc4.selectbox("Owner",      all_owners, key="pma_f_own")

        all_status = ["All", "Open", "In Progress", "Closed"]
        f_status = fc5.selectbox("Status",     all_status, key="pma_f_sta")

    # Apply filters
    display_df = actions_df.copy()
    if f_priority != "All":
        display_df = display_df[display_df["Priority"] == f_priority]
    if f_cat != "All":
        display_df = display_df[display_df["Category"] == f_cat]
    if f_wbs != "All":
        display_df = display_df[display_df["WBS"] == f_wbs]
    if f_owner != "All":
        display_df = display_df[display_df["Owner"] == f_owner]
    if f_status != "All":
        display_df = display_df[display_df["Status"] == f_status]

    n_shown = len(display_df)
    st.caption(f"Showing {n_shown} of {n_total} actions.")

    if display_df.empty:
        st.info("No actions match the current filters.")
        return

    # ---- Editable action table ----------------------------------------------
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
        'text-transform:uppercase;margin-bottom:8px;">Action List</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Edit the Owner, Due Date and Status columns directly in the table below. "
        "Changes are saved while you are on this page."
    )

    # Colour-code Priority column
    def _style_priority(row):
        p = row.get("Priority","")
        base = [""] * len(row)
        idx  = list(row.index)
        try:
            pi = idx.index("Priority")
        except ValueError:
            return base
        colours = {"High":"background-color:#fef2f2;color:#991b1b;font-weight:700;",
                   "Medium":"background-color:#fffbeb;color:#92400e;font-weight:600;",
                   "Low":"background-color:#f8fafc;color:#475569;"}
        base[pi] = colours.get(p,"")
        return base

    # Use st.data_editor for the editable fields
    edit_cols = [
        "Priority","Category","Activity ID","Activity Name","WBS",
        "Issue","Why It Matters","Suggested Action",
        "Owner","Due Date","Status",
    ]
    avail_edit = [c for c in edit_cols if c in display_df.columns]

    edited_df = st.data_editor(
        display_df[avail_edit],
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Priority": st.column_config.SelectboxColumn(
                "Priority",
                options=["High","Medium","Low"],
                width="small",
            ),
            "Status": st.column_config.SelectboxColumn(
                "Status",
                options=["Open","In Progress","Closed"],
                width="small",
            ),
            "Owner": st.column_config.TextColumn("Owner", width="small"),
            "Due Date": st.column_config.TextColumn("Due Date", width="small",
                                                     help="e.g. 01/06/2025"),
            "Issue":            st.column_config.TextColumn("Issue",            width="medium"),
            "Why It Matters":   st.column_config.TextColumn("Why It Matters",   width="large"),
            "Suggested Action": st.column_config.TextColumn("Suggested Action", width="large"),
            "Activity ID":      st.column_config.TextColumn("Activity ID",      width="small"),
            "Activity Name":    st.column_config.TextColumn("Activity Name",    width="medium"),
            "WBS":              st.column_config.TextColumn("WBS",              width="medium"),
            "Category":         st.column_config.TextColumn("Category",         width="medium"),
        },
        key="pma_editor",
    )

    # Persist edits back to session state
    if edited_df is not None and not edited_df.empty:
        for col in ["Owner","Due Date","Status"]:
            if col in edited_df.columns:
                actions_df.loc[display_df.index, col] = edited_df[col].values
        st.session_state["_pm_actions_df"] = actions_df

    # ---- Highlighted high-priority section ----------------------------------
    high_df = display_df[display_df["Priority"] == "High"]
    if not high_df.empty:
        st.divider()
        st.markdown(
            f'<div style="font-size:12px;font-weight:700;color:#dc2626;letter-spacing:1px;'
            f'text-transform:uppercase;margin-bottom:12px;">'
            f'High Priority Actions ({len(high_df)})</div>',
            unsafe_allow_html=True,
        )
        for _, act in high_df.iterrows():
            st.markdown(
                f"""
                <div style="background:#ffffff;border:1px solid #fca5a5;border-left:4px solid #dc2626;
                            border-radius:8px;padding:14px 18px;margin-bottom:8px;">
                    <div style="display:flex;justify-content:space-between;align-items:flex-start;
                                flex-wrap:wrap;gap:8px;">
                        <div style="flex:1;min-width:200px;">
                            <span style="background:#dc2626;color:white;padding:2px 8px;
                                         border-radius:4px;font-size:10px;font-weight:700;
                                         letter-spacing:0.5px;text-transform:uppercase;">
                                {act.get("Category","")}
                            </span>
                            <div style="font-weight:700;color:#0B1F33;font-size:14px;
                                        margin-top:6px;">
                                {act.get("Activity ID","")} - {act.get("Activity Name","")}
                            </div>
                            <div style="font-size:12px;color:#64748B;margin-top:2px;">
                                {act.get("WBS","")}
                            </div>
                        </div>
                    </div>
                    <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;
                                gap:10px;">
                        <div>
                            <div style="font-size:10px;font-weight:700;color:#94A3B8;
                                        text-transform:uppercase;letter-spacing:0.8px;">Issue</div>
                            <div style="font-size:13px;color:#334155;margin-top:3px;">
                                {act.get("Issue","")}</div>
                        </div>
                        <div>
                            <div style="font-size:10px;font-weight:700;color:#94A3B8;
                                        text-transform:uppercase;letter-spacing:0.8px;">
                                Suggested Action</div>
                            <div style="font-size:13px;color:#334155;margin-top:3px;">
                                {act.get("Suggested Action","")}</div>
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ---- Export -------------------------------------------------------------
    st.divider()

    # Merge any edits back for export
    export_df = st.session_state.get("_pm_actions_df", actions_df).copy()

    # Add summary sheet
    summary_rows = {
        "Metric": ["Total Actions","High Priority","Medium Priority","Low Priority",
                   "Open","In Progress","Closed"],
        "Count":  [
            len(export_df),
            int((export_df["Priority"]=="High").sum()),
            int((export_df["Priority"]=="Medium").sum()),
            int((export_df["Priority"]=="Low").sum()),
            int((export_df["Status"]=="Open").sum()),
            int((export_df["Status"]=="In Progress").sum()),
            int((export_df["Status"]=="Closed").sum()),
        ],
    }

    export_sheets = {
        "Summary":           pd.DataFrame(summary_rows),
        "All Actions":       export_df,
        "High Priority":     export_df[export_df["Priority"]=="High"],
        "Medium Priority":   export_df[export_df["Priority"]=="Medium"],
        "Low Priority":      export_df[export_df["Priority"]=="Low"],
    }
    # Category sheets
    for cat in export_df["Category"].unique():
        safe_name = cat[:31]
        export_sheets[safe_name] = export_df[export_df["Category"]==cat]

    xls_bytes = export_df_to_excel(export_sheets)

    dl_col, _ = st.columns([1,3])
    dl_col.download_button(
        label="📥  Export Action List to Excel",
        data=xls_bytes,
        file_name=f"pm_actions_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Exports all actions with Summary, High/Medium/Low priority sheets and one sheet per category.",
    )

    st.caption(
        "Tip: Export the action list and share it with your delivery team. "
        "Update Owner and Status as actions are completed."
    )



# -----------------------------------------------------------------------------
# PAGE: LOOKAHEAD PLANNER
# -----------------------------------------------------------------------------

def _lookahead_card(label: str, value, sublabel: str = "",
                    colour: str = "#0B1F33", bg: str = "#ffffff",
                    border_top: str = "#0B1F33") -> str:
    """Compact metric card matching PlanTrace brand."""
    return (
        f'<div style="background:{bg};border:1px solid #E2E8F0;border-radius:10px;'
        f'padding:16px 18px;border-top:3px solid {border_top};'
        f'box-shadow:0 1px 4px rgba(11,31,51,0.06);">'
        f'<div style="font-size:10px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
        f'text-transform:uppercase;margin-bottom:6px;">{label}</div>'
        f'<div style="font-size:26px;font-weight:800;color:{colour};line-height:1;">{value}</div>'
        f'{"" if not sublabel else f"<div style=font-size:11px;color:#64748B;margin-top:4px;>{sublabel}</div>"}'
        f'</div>'
    )


def _lookahead_section_header(title: str, count: int, colour: str = "#0B1F33") -> str:
    return (
        f'<div style="display:flex;align-items:center;gap:10px;margin:16px 0 8px 0;">'
        f'<div style="font-size:14px;font-weight:700;color:{colour};">{title}</div>'
        f'<div style="background:{colour};color:white;border-radius:12px;'
        f'padding:1px 9px;font-size:11px;font-weight:700;">{count}</div>'
        f'</div>'
    )


def _week_labour(task_res_df: pd.DataFrame, tasks_df: pd.DataFrame,
                 window_start: datetime, window_end: datetime) -> pd.DataFrame:
    """
    Expand resource loading into a weekly series within the lookahead window.
    Returns a DataFrame with columns: week, rsrc_name, qty.
    """
    if task_res_df.empty:
        return pd.DataFrame()

    rows = []
    for _, r in task_res_df.iterrows():
        s = pd.to_datetime(r.get("target_start") or r.get("target_start_date"), errors="coerce")
        e = pd.to_datetime(r.get("target_finish") or r.get("target_end_date"), errors="coerce")
        if pd.isna(s) or pd.isna(e) or s > e:
            continue

        # Clip to window
        s_clip = max(s, pd.Timestamp(window_start))
        e_clip = min(e, pd.Timestamp(window_end))
        if s_clip > e_clip:
            continue

        qty = safe_float(r.get("target_qty", 0), 0)
        if qty == 0:
            continue

        total_days = max(1, (e - s).days)
        window_days = max(1, (e_clip - s_clip).days)
        qty_in_window = qty * (window_days / total_days)

        weeks = max(1, math.ceil(window_days / 7))
        qty_per_week = qty_in_window / weeks

        current = s_clip
        for _ in range(weeks):
            rows.append({
                "week":      current.to_period("W").start_time,
                "rsrc_name": str(r.get("rsrc_name", r.get("rsrc_id", "Unknown"))),
                "qty":       qty_per_week,
            })
            current += timedelta(weeks=1)

    return pd.DataFrame(rows) if rows else pd.DataFrame()


def page_lookahead(data: dict, near_crit_days: float):
    """
    Lookahead Planner page.
    Shows all activities starting, finishing, or at risk within a user-defined
    lookahead window, with labour demand and export.
    """
    st.title("📅 Lookahead Planner")
    st.caption(
        "A short-term delivery view from your programme. "
        "Select a lookahead window to see what is starting, finishing and at risk."
    )

    tasks    = data["tasks_df"]
    rels     = data["relationships_df"]
    task_res = data.get("task_resources_df", pd.DataFrame())

    if tasks.empty:
        st.warning("No activities found. Please upload a programme first.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)

    # ---- Lookahead window selector ------------------------------------------
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
        'text-transform:uppercase;margin-bottom:8px;">Lookahead Window</div>',
        unsafe_allow_html=True,
    )

    wcol1, wcol2 = st.columns([2, 3])

    with wcol1:
        window_opt = st.selectbox(
            "Window",
            ["2 Weeks", "4 Weeks", "6 Weeks", "12 Weeks", "Custom Date Range"],
            index=1,
            label_visibility="collapsed",
            key="la_window_opt",
        )

    now = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    window_map = {
        "2 Weeks":  14,
        "4 Weeks":  28,
        "6 Weeks":  42,
        "12 Weeks": 84,
    }

    with wcol2:
        if window_opt == "Custom Date Range":
            dc1, dc2 = st.columns(2)
            window_start = dc1.date_input(
                "From", value=now.date(), key="la_custom_start"
            )
            window_end = dc2.date_input(
                "To", value=(now + timedelta(weeks=4)).date(), key="la_custom_end"
            )
            window_start = datetime.combine(window_start, datetime.min.time())
            window_end   = datetime.combine(window_end,   datetime.max.time().replace(microsecond=0))
        else:
            days = window_map[window_opt]
            window_start = now
            window_end   = now + timedelta(days=days)
            st.markdown(
                f'<div style="padding:8px 14px;background:#0B1F33;border-radius:8px;'
                f'font-size:13px;color:#CBD5E1;display:inline-block;">'
                f'{window_start.strftime("%d %b %Y")} &rarr; '
                f'{window_end.strftime("%d %b %Y")}'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- Filters ------------------------------------------------------------
    with st.expander("Filters", expanded=False):
        fl1, fl2, fl3, fl4 = st.columns(4)

        wbs_opts = ["All"]
        if "wbs_path" in tasks.columns:
            tops = tasks["wbs_path"].dropna().apply(
                lambda x: str(x).split(" > ")[0] if x and str(x).strip() not in ("","nan") else "Unknown"
            ).unique().tolist()
            wbs_opts += sorted(tops)
        f_wbs = fl1.selectbox("WBS / Area", wbs_opts, key="la_f_wbs")

        status_opts = ["All"] + sorted([
            s for s in tasks["status"].dropna().unique() if str(s).strip()
        ]) if "status" in tasks.columns else ["All"]
        f_status = fl2.selectbox("Activity Status", status_opts, key="la_f_stat")

        f_crit = fl3.selectbox(
            "Float Filter",
            ["All", "Critical only", "Near-critical only", "Negative float only"],
            key="la_f_crit",
        )

        f_type = fl4.selectbox(
            "Activity Type",
            ["All"] + (sorted([t for t in tasks["task_type"].dropna().unique() if str(t).strip()])
                       if "task_type" in tasks.columns else []),
            key="la_f_type",
        )

    # ---- Build lookahead activity sets --------------------------------------
    def _in_window(d, strict_start=False):
        """True if datetime d falls within the lookahead window."""
        if d is None or not hasattr(d, "date"):
            return False
        if strict_start:
            return window_start <= d <= window_end
        return d <= window_end

    # Activities starting in window
    starting = tasks[
        tasks["eff_start"].apply(lambda d: _in_window(d, strict_start=True))
    ].copy() if "eff_start" in tasks.columns else pd.DataFrame()

    # Activities finishing in window
    finishing = tasks[
        tasks["eff_finish"].apply(lambda d: _in_window(d, strict_start=True))
    ].copy() if "eff_finish" in tasks.columns else pd.DataFrame()

    # Activities spanning the window (started before, finish after window start)
    spanning = tasks[
        tasks["eff_start"].apply(lambda d: d is not None and hasattr(d,"date") and d < window_start) &
        tasks["eff_finish"].apply(lambda d: d is not None and hasattr(d,"date") and d >= window_start)
    ].copy() if ("eff_start" in tasks.columns and "eff_finish" in tasks.columns) else pd.DataFrame()

    # All activities in window = union
    in_window_ids = (
        set(starting["task_id"].tolist() if not starting.empty else []) |
        set(finishing["task_id"].tolist() if not finishing.empty else []) |
        set(spanning["task_id"].tolist() if not spanning.empty else [])
    )
    all_window = tasks[tasks["task_id"].isin(in_window_ids)].copy()

    # Apply filters to all_window
    def _apply_filters(df):
        if df.empty:
            return df
        if f_wbs != "All" and "wbs_path" in df.columns:
            df = df[df["wbs_path"].astype(str).str.startswith(f_wbs)]
        if f_status != "All" and "status" in df.columns:
            df = df[df["status"] == f_status]
        if f_crit == "Critical only" and "is_critical" in df.columns:
            df = df[df["is_critical"] == True]
        elif f_crit == "Near-critical only" and "is_near_critical" in df.columns:
            df = df[df["is_near_critical"] == True]
        elif f_crit == "Negative float only" and "total_float_days" in df.columns:
            df = df[df["total_float_days"].apply(lambda f: safe_float(f, 0) < 0)]
        if f_type != "All" and "task_type" in df.columns:
            df = df[df["task_type"] == f_type]
        return df

    starting  = _apply_filters(starting)
    finishing = _apply_filters(finishing)
    spanning  = _apply_filters(spanning)
    all_window = _apply_filters(all_window)

    # Open logic in window
    tasks_with_pred = set(rels["succ_task_id"].dropna()) if not rels.empty and "succ_task_id" in rels.columns else set()
    tasks_with_succ = set(rels["pred_task_id"].dropna()) if not rels.empty and "pred_task_id" in rels.columns else set()

    open_start_win  = all_window[~all_window["task_id"].isin(tasks_with_pred)] if "task_id" in all_window.columns else pd.DataFrame()
    open_finish_win = all_window[~all_window["task_id"].isin(tasks_with_succ)] if "task_id" in all_window.columns else pd.DataFrame()

    # Critical & near-critical in window
    crit_win  = all_window[all_window["is_critical"] == True] if "is_critical" in all_window.columns else pd.DataFrame()
    nc_win    = all_window[all_window["is_near_critical"] == True] if "is_near_critical" in all_window.columns else pd.DataFrame()
    neg_win   = all_window[all_window["total_float_days"].apply(lambda f: safe_float(f,0) < 0)] if "total_float_days" in all_window.columns else pd.DataFrame()

    # Notes matching
    notes_text = st.session_state.get("_notes_text", "")
    notes_ids  = set()
    if notes_text and "task_code" in all_window.columns:
        for code in all_window["task_code"].dropna():
            if str(code) in notes_text:
                notes_ids.add(str(code))
    notes_win = all_window[all_window["task_code"].astype(str).isin(notes_ids)] if notes_ids else pd.DataFrame()

    risk_ids = set()
    if notes_text and "task_code" in all_window.columns:
        for _, t in all_window.iterrows():
            code = str(t.get("task_code",""))
            if not code:
                continue
            idx = notes_text.find(code)
            if idx < 0:
                continue
            snippet = notes_text[max(0,idx-300):idx+300]
            for word in _RISK_WORDS:
                if re.search(r'\b' + re.escape(word) + r'\b', snippet, re.IGNORECASE):
                    risk_ids.add(code)
                    break
    risk_win = all_window[all_window["task_code"].astype(str).isin(risk_ids)] if risk_ids else pd.DataFrame()

    # Labour in window
    labour_weekly = _week_labour(task_res, tasks, window_start, window_end) if not task_res.empty else pd.DataFrame()
    peak_labour   = int(labour_weekly.groupby("week")["qty"].sum().max()) if not labour_weekly.empty else 0

    # ---- Summary metric cards -----------------------------------------------
    open_logic_count = len(open_start_win) + len(open_finish_win)

    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
        'text-transform:uppercase;margin-bottom:10px;">Window Summary</div>',
        unsafe_allow_html=True,
    )

    r1 = st.columns(6)
    r1[0].markdown(_lookahead_card(
        "Starting",       len(starting),
        f"in {window_opt.lower()}",
        "#2563eb", "#eff6ff", "#2563eb"
    ), unsafe_allow_html=True)
    r1[1].markdown(_lookahead_card(
        "Finishing",      len(finishing),
        f"in {window_opt.lower()}",
        "#0B1F33", "#f0f9ff", "#0B1F33"
    ), unsafe_allow_html=True)
    r1[2].markdown(_lookahead_card(
        "Critical",       len(crit_win),
        "in window",
        "#dc2626", "#fef2f2" if crit_win.empty else "#fef2f2", "#dc2626"
    ), unsafe_allow_html=True)
    r1[3].markdown(_lookahead_card(
        "Negative Float", len(neg_win),
        "in window",
        "#dc2626" if not neg_win.empty else "#64748B",
        "#fef2f2" if not neg_win.empty else "#f8fafc",
        "#dc2626" if not neg_win.empty else "#E2E8F0"
    ), unsafe_allow_html=True)
    r1[4].markdown(_lookahead_card(
        "Labour Peak",
        f"{peak_labour:,}" if peak_labour else "N/A",
        "hrs/week in window",
        "#7c3aed", "#faf5ff", "#7c3aed"
    ), unsafe_allow_html=True)
    r1[5].markdown(_lookahead_card(
        "Open Logic",     open_logic_count,
        "in window",
        "#d97706" if open_logic_count else "#64748B",
        "#fffbeb" if open_logic_count else "#f8fafc",
        "#d97706" if open_logic_count else "#E2E8F0"
    ), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if all_window.empty:
        st.info(
            f"No activities found in the selected window "
            f"({window_start.strftime('%d %b %Y')} to {window_end.strftime('%d %b %Y')}). "
            "Try extending the lookahead window or removing filters."
        )
        return

    # ---- Display columns helper ---------------------------------------------
    BASE_COLS = {
        "task_code":        "Activity ID",
        "task_name":        "Activity Name",
        "wbs_path":         "WBS",
        "eff_start":        "Start",
        "eff_finish":       "Finish",
        "orig_dur_days":    "Orig Dur (d)",
        "rem_dur_days":     "Rem Dur (d)",
        "total_float_days": "Float (d)",
        "status":           "Status",
        "is_critical":      "Critical",
    }

    def _fmt_table(df, extra_cols=None):
        cols = dict(BASE_COLS)
        if extra_cols:
            cols.update(extra_cols)
        avail = {k: v for k, v in cols.items() if k in df.columns}
        out = df[list(avail.keys())].copy().rename(columns=avail)
        for col in ["Start","Finish"]:
            if col in out.columns:
                out[col] = out[col].apply(format_date)
        if "Status" in out.columns:
            out["Status"] = out["Status"].apply(_status_label)
        if "Critical" in out.columns:
            out["Critical"] = out["Critical"].apply(lambda x: "Yes" if x else "")
        if "Float (d)" in out.columns:
            out["Float (d)"] = out["Float (d)"].apply(
                lambda x: round(float(x),1) if x is not None and str(x) not in ("","nan") else "-"
            )
        return out

    def _crit_style(df):
        """Apply row colour based on float."""
        def _row_style(row):
            flag = _crit_flag(row.get("Float (d)", None))
            colour_map = {
                "Negative Float": "background-color:#fecaca;",
                "Critical":       "background-color:#fee2e2;",
                "Near-Critical":  "background-color:#fef3c7;",
            }
            style = colour_map.get(flag, "")
            return [style] * len(row)
        try:
            return df.style.apply(_row_style, axis=1)
        except Exception:
            return df

    # ---- Main tabs ----------------------------------------------------------
    (tab_start, tab_finish, tab_crit, tab_nc, tab_neg,
     tab_open, tab_notes, tab_labour, tab_gantt, tab_export) = st.tabs([
        f"Starting ({len(starting)})",
        f"Finishing ({len(finishing)})",
        f"Critical ({len(crit_win)})",
        f"Near-Critical ({len(nc_win)})",
        f"Negative Float ({len(neg_win)})",
        f"Open Logic ({open_logic_count})",
        f"Planning Notes ({len(notes_win) + len(risk_win)})",
        "Labour Demand",
        "Gantt",
        "Export",
    ])

    # -- Starting -------------------------------------------------------------
    with tab_start:
        st.markdown(
            f"Activities with a scheduled start date between "
            f"**{window_start.strftime('%d %b %Y')}** and "
            f"**{window_end.strftime('%d %b %Y')}**."
        )
        if starting.empty:
            st.info("No activities are scheduled to start in this window.")
        else:
            st.dataframe(_crit_style(_fmt_table(starting.sort_values("eff_start"))),
                         use_container_width=True, hide_index=True)

    # -- Finishing -------------------------------------------------------------
    with tab_finish:
        st.markdown(
            f"Activities with a scheduled finish date between "
            f"**{window_start.strftime('%d %b %Y')}** and "
            f"**{window_end.strftime('%d %b %Y')}**."
        )
        if finishing.empty:
            st.info("No activities are scheduled to finish in this window.")
        else:
            st.dataframe(_crit_style(_fmt_table(finishing.sort_values("eff_finish"))),
                         use_container_width=True, hide_index=True)

    # -- Critical --------------------------------------------------------------
    with tab_crit:
        if crit_win.empty:
            st.success("No critical activities fall within this window.")
        else:
            st.warning(
                f"**{len(crit_win)} critical activities** are active or due in this window. "
                "Any delay to these directly impacts the project finish date."
            )
            st.dataframe(_fmt_table(crit_win.sort_values("eff_finish")),
                         use_container_width=True, hide_index=True)

    # -- Near-Critical ---------------------------------------------------------
    with tab_nc:
        if nc_win.empty:
            st.success("No near-critical activities fall within this window.")
        else:
            st.info(
                f"**{len(nc_win)} near-critical activities** are active or due in this window. "
                f"Each has between 0 and {near_crit_days} days float."
            )
            st.dataframe(_fmt_table(nc_win.sort_values("total_float_days")),
                         use_container_width=True, hide_index=True)

    # -- Negative Float --------------------------------------------------------
    with tab_neg:
        if neg_win.empty:
            st.success("No activities with negative float are active in this window.")
        else:
            st.error(
                f"**{len(neg_win)} activities have negative float** in this window. "
                "These activities are already beyond their target dates and need immediate attention."
            )
            st.dataframe(_fmt_table(neg_win.sort_values("total_float_days")),
                         use_container_width=True, hide_index=True)

    # -- Open Logic ------------------------------------------------------------
    with tab_open:
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Open Start** -- no predecessors")
            if open_start_win.empty:
                st.success("No open-start activities in this window.")
            else:
                st.warning(f"{len(open_start_win)} activities have no predecessor in this window.")
                st.dataframe(_fmt_table(open_start_win), use_container_width=True, hide_index=True)
        with col_b:
            st.markdown("**Open Finish** -- no successors")
            if open_finish_win.empty:
                st.success("No open-finish activities in this window.")
            else:
                st.warning(f"{len(open_finish_win)} activities have no successor in this window.")
                st.dataframe(_fmt_table(open_finish_win), use_container_width=True, hide_index=True)

    # -- Planning Notes --------------------------------------------------------
    with tab_notes:
        if not notes_text:
            st.info(
                "No planning notes loaded. Upload notes on the Planning Notes page "
                "and return here to see which activities in this window are mentioned."
            )
        else:
            col_n1, col_n2 = st.columns(2)
            with col_n1:
                st.markdown("**Activities mentioned in planning notes**")
                if notes_win.empty:
                    st.info("No activities in this window are mentioned in the planning notes.")
                else:
                    st.dataframe(_fmt_table(notes_win), use_container_width=True, hide_index=True)
            with col_n2:
                st.markdown("**Activities with risk keywords in notes**")
                if risk_win.empty:
                    st.info("No risk keywords found against activities in this window.")
                else:
                    st.warning(
                        f"{len(risk_win)} activities in this window appear alongside "
                        "risk keywords (delay, blocked, CE, EWN, constraint etc.) in the planning notes."
                    )
                    st.dataframe(_fmt_table(risk_win), use_container_width=True, hide_index=True)

    # -- Labour Demand ---------------------------------------------------------
    with tab_labour:
        if labour_weekly.empty:
            st.info(
                "No resource loading data found in this XER. "
                "If the programme was resourced in P6, check the export included resources. "
                "You can also upload a resource CSV on the Labour Histogram page."
            )
        else:
            weekly_total = labour_weekly.groupby("week")["qty"].sum().reset_index()
            weekly_total["week_str"] = weekly_total["week"].dt.strftime("%d %b")

            peak_wk = weekly_total.loc[weekly_total["qty"].idxmax()]
            avg_hrs  = weekly_total["qty"].mean()
            total_hrs = weekly_total["qty"].sum()

            lm1, lm2, lm3 = st.columns(3)
            lm1.metric("Total Hours in Window", f"{total_hrs:,.0f}")
            lm2.metric("Peak Week",             f"{peak_wk['qty']:,.0f} hrs ({peak_wk['week_str']})")
            lm3.metric("Average per Week",      f"{avg_hrs:,.0f} hrs")

            fig_lab = px.bar(
                weekly_total, x="week_str", y="qty",
                title=f"Labour Demand by Week ({window_opt})",
                labels={"week_str": "Week", "qty": "Hours"},
                color_discrete_sequence=["#0B1F33"],
            )
            fig_lab.update_layout(
                height=320, margin=dict(l=10,r=10,t=40,b=10),
                plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
            )
            st.plotly_chart(fig_lab, use_container_width=True)

            # By resource
            if "rsrc_name" in labour_weekly.columns:
                by_res = labour_weekly.groupby("rsrc_name")["qty"].sum().reset_index() \
                             .sort_values("qty", ascending=False).head(15)
                fig_res = px.bar(
                    by_res, x="qty", y="rsrc_name", orientation="h",
                    title="Hours by Resource in Window",
                    labels={"qty":"Hours","rsrc_name":""},
                    color_discrete_sequence=["#7c3aed"],
                )
                fig_res.update_yaxes(autorange="reversed")
                fig_res.update_layout(
                    height=350, margin=dict(l=10,r=10,t=40,b=10),
                    plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
                )
                st.plotly_chart(fig_res, use_container_width=True)

    # -- Gantt -----------------------------------------------------------------
    with tab_gantt:
        st.markdown(
            f"All activities active in the **{window_opt.lower()}** window, "
            "colour-coded by float status."
        )
        gantt_src = all_window.dropna(subset=["eff_start","eff_finish"]).copy() \
            if "eff_start" in all_window.columns and "eff_finish" in all_window.columns \
            else pd.DataFrame()

        if gantt_src.empty:
            st.info("No date data available for the Gantt chart.")
        else:
            gantt_src["Label"] = (
                gantt_src["task_code"].astype(str) + "  " +
                gantt_src["task_name"].astype(str).str[:40]
            )
            gantt_src["Float Status"] = gantt_src["total_float_days"].apply(
                lambda f: (
                    "Negative Float" if safe_float(f,0) < 0 else
                    "Critical"       if safe_float(f,0) == 0 else
                    "Near-Critical"  if safe_float(f,0) <= near_crit_days else
                    "Has Float"
                )
            )
            gantt_src = gantt_src.sort_values("eff_start")
            max_gantt = 100

            fig_g = px.timeline(
                gantt_src.head(max_gantt),
                x_start="eff_start", x_end="eff_finish", y="Label",
                color="Float Status",
                color_discrete_map={
                    "Negative Float": "#7f1d1d",
                    "Critical":       "#dc2626",
                    "Near-Critical":  "#d97706",
                    "Has Float":      "#2563eb",
                },
                title=f"Lookahead Gantt: {window_opt} ({window_start.strftime('%d %b')} - {window_end.strftime('%d %b %Y')})",
                labels={"Label":""},
            )
            fig_g.update_yaxes(autorange="reversed")

            # Shade the window
            fig_g.add_vrect(
                x0=window_start, x1=window_end,
                fillcolor="#F5A623", opacity=0.06,
                line_width=0, annotation_text="Lookahead Window",
                annotation_position="top left",
                annotation=dict(font_color="#F5A623", font_size=11),
            )
            fig_g.add_vline(
                x=now, line_dash="dash", line_color="#0B1F33",
                annotation_text="Today", annotation_position="top right",
                annotation=dict(font_color="#0B1F33", font_size=11),
            )
            fig_g.update_layout(
                height=max(350, min(900, 55 + len(gantt_src.head(max_gantt)) * 26)),
                margin=dict(l=10,r=10,t=50,b=10),
                plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
                legend_title_text="Float Status",
            )
            st.plotly_chart(fig_g, use_container_width=True)
            if len(gantt_src) > max_gantt:
                st.caption(f"Showing first {max_gantt} of {len(gantt_src)} activities. Apply filters to narrow the view.")

    # -- Export ----------------------------------------------------------------
    with tab_export:
        st.markdown("**Download the full lookahead as a formatted Excel workbook.**")

        window_label = (
            f"{window_opt}: {window_start.strftime('%d %b %Y')} to {window_end.strftime('%d %b %Y')}"
        )

        summary_sheet = pd.DataFrame({
            "Metric": [
                "Lookahead Window", "From", "To",
                "Activities Starting", "Activities Finishing",
                "Activities in Window (total)",
                "Critical", "Near-Critical", "Negative Float",
                "Open Start", "Open Finish",
                "In Planning Notes", "Risk Keywords",
                "Labour Peak (hrs/week)",
            ],
            "Value": [
                window_label,
                window_start.strftime("%d %b %Y"),
                window_end.strftime("%d %b %Y"),
                len(starting), len(finishing), len(all_window),
                len(crit_win), len(nc_win), len(neg_win),
                len(open_start_win), len(open_finish_win),
                len(notes_win), len(risk_win),
                peak_labour if peak_labour else "N/A",
            ],
        })

        def _exp(df):
            if df is None or df.empty:
                return pd.DataFrame(columns=["No data"])
            return _fmt_table(df)

        export_sheets = {
            "Summary":          summary_sheet,
            "Starting":         _exp(starting.sort_values("eff_start")    if not starting.empty  else starting),
            "Finishing":        _exp(finishing.sort_values("eff_finish")   if not finishing.empty else finishing),
            "All in Window":    _exp(all_window.sort_values("eff_start")   if not all_window.empty else all_window),
            "Critical":         _exp(crit_win),
            "Near-Critical":    _exp(nc_win),
            "Negative Float":   _exp(neg_win),
            "Open Start":       _exp(open_start_win),
            "Open Finish":      _exp(open_finish_win),
        }
        if not notes_win.empty:
            export_sheets["In Planning Notes"] = _exp(notes_win)
        if not risk_win.empty:
            export_sheets["Risk Keywords"]     = _exp(risk_win)
        if not labour_weekly.empty:
            lw_exp = labour_weekly.copy()
            lw_exp["week"] = lw_exp["week"].dt.strftime("%d %b %Y")
            export_sheets["Labour by Week"] = lw_exp.groupby(["week","rsrc_name"])["qty"] \
                .sum().reset_index().rename(columns={"week":"Week","rsrc_name":"Resource","qty":"Hours"})

        xls_bytes = export_df_to_excel(export_sheets)

        st.download_button(
            label="📥  Download Lookahead Report",
            data=xls_bytes,
            file_name=f"lookahead_{window_opt.replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Exports summary, starting, finishing, critical, open logic and labour sheets.",
        )

        st.markdown(
            '<div style="background:#f8fafc;border:1px solid #E2E8F0;border-radius:8px;'
            'padding:14px 18px;margin-top:12px;">'
            '<div style="font-size:12px;font-weight:700;color:#0B1F33;margin-bottom:6px;">'
            'Workbook sheets</div>'
            '<div style="font-size:12px;color:#64748B;line-height:2;">'
            + "".join(f"<strong>{k}</strong> &nbsp;|&nbsp; " for k in export_sheets.keys()) +
            '</div></div>',
            unsafe_allow_html=True,
        )




# -----------------------------------------------------------------------------
# PAGE: MILESTONE TRACKER
# -----------------------------------------------------------------------------

# Keywords that indicate a milestone activity by name
_MILESTONE_KEYWORDS = [
    "milestone", "complete", "completion", "handover", "hand over",
    "energisation", "energization", "energise", "energize",
    "access", "installation start", "install start",
    "commissioning", "commission", "delivery", "deliver",
    "approval", "approve", "sign off", "sign-off", "signoff",
    "ready for", "available", "issued", "award", "award of contract",
    "start on site", "mobilisation", "mobilization", "end", "finish",
    "practical completion", "pc", "substantial completion",
    "first fix", "second fix", "inspection", "witness",
    "notice to proceed", "ntp", "key date",
]


def _detect_milestones(tasks: pd.DataFrame) -> pd.DataFrame:
    """
    Auto-detect milestone activities from the tasks DataFrame.
    Returns a copy with a boolean 'is_milestone_detected' column.
    """
    df = tasks.copy()

    flags = pd.Series(False, index=df.index)

    # 1. Activity type contains "milestone" (covers TT_Mile and TT_FinMile)
    if "task_type" in df.columns:
        flags |= df["task_type"].astype(str).str.lower().str.contains(
            "milestone", na=False
        )

    # 2. Zero original duration
    if "orig_dur_days" in df.columns:
        flags |= df["orig_dur_days"].apply(lambda d: safe_float(d, 1) == 0)

    # 3. Name contains a milestone keyword
    if "task_name" in df.columns:
        kw_pattern = "|".join(re.escape(k) for k in _MILESTONE_KEYWORDS)
        flags |= df["task_name"].astype(str).str.lower().str.contains(
            kw_pattern, na=False
        )

    df["is_milestone_detected"] = flags
    return df


def _risk_rating(tf, has_constraint: bool = False, in_notes_risk: bool = False) -> str:
    """Return a simple risk rating string for a milestone."""
    f = safe_float(tf, 9999)
    if f < 0 or in_notes_risk:
        return "High"
    if f == 0 or has_constraint:
        return "High"
    if f <= 10:
        return "Medium"
    return "Low"


def _risk_colour(rating: str) -> str:
    return {"High": "#dc2626", "Medium": "#d97706", "Low": "#16a34a"}.get(rating, "#6b7280")


def _risk_bg(rating: str) -> str:
    return {"High": "#fef2f2", "Medium": "#fffbeb", "Low": "#f0fdf4"}.get(rating, "#f8fafc")


def _milestone_header_card(row: pd.Series, movement_days=None) -> str:
    """Render a single milestone summary card as HTML."""
    tf        = safe_float(row.get("total_float_days"), None)
    f_col     = _float_color(tf)
    is_crit   = bool(row.get("is_critical", False))
    status    = _status_label(str(row.get("status", "")))
    s_col     = _status_colour(str(row.get("status", "")))
    code      = str(row.get("task_code", "-"))
    name      = str(row.get("task_name", "-"))
    wbs       = str(row.get("wbs_path", "-"))
    finish    = format_date(row.get("eff_finish"))
    start     = format_date(row.get("eff_start"))
    rating    = _risk_rating(
        tf,
        has_constraint=bool(row.get("cstr_type","")) and str(row.get("cstr_type","")).strip() not in ("","None","nan"),
    )
    r_col     = _risk_colour(rating)

    crit_pill = (
        '<span style="background:#dc2626;color:white;padding:1px 8px;'
        'border-radius:10px;font-size:10px;font-weight:700;margin-left:6px;">CRITICAL</span>'
        if is_crit else ""
    )

    move_html = ""
    if movement_days is not None:
        m_col  = "#dc2626" if movement_days > 0 else "#16a34a" if movement_days < 0 else "#6b7280"
        m_sign = "+" if movement_days > 0 else ""
        move_html = (
            f'<span style="background:{m_col};color:white;padding:2px 8px;'
            f'border-radius:10px;font-size:10px;font-weight:700;margin-left:6px;">'
            f'{m_sign}{movement_days}d movement</span>'
        )

    return f"""
    <div style="background:#ffffff;border:1px solid #E2E8F0;border-radius:12px;
                padding:18px 22px;margin-bottom:10px;
                border-left:5px solid {r_col};
                box-shadow:0 2px 6px rgba(11,31,51,0.07);">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px;">
            <div style="flex:1;min-width:0;">
                <div style="font-size:10px;font-weight:700;color:#94A3B8;letter-spacing:1px;
                            text-transform:uppercase;margin-bottom:4px;">Milestone</div>
                <div style="font-size:16px;font-weight:800;color:#0B1F33;">
                    {code}{crit_pill}{move_html}
                </div>
                <div style="font-size:14px;color:#334155;margin-top:3px;
                            white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"
                     title="{name}">{name}</div>
                <div style="font-size:11px;color:#94A3B8;margin-top:2px;">{wbs}</div>
            </div>
            <div style="display:grid;grid-template-columns:repeat(4,auto);gap:10px;align-items:start;">
                <div style="text-align:center;">
                    <div style="font-size:9px;color:#94A3B8;text-transform:uppercase;
                                letter-spacing:0.8px;">Start</div>
                    <div style="font-size:12px;font-weight:600;color:#0B1F33;margin-top:2px;">{start}</div>
                </div>
                <div style="text-align:center;">
                    <div style="font-size:9px;color:#94A3B8;text-transform:uppercase;
                                letter-spacing:0.8px;">Finish</div>
                    <div style="font-size:12px;font-weight:700;color:#0B1F33;margin-top:2px;">{finish}</div>
                </div>
                <div style="text-align:center;">
                    <div style="font-size:9px;color:#94A3B8;text-transform:uppercase;
                                letter-spacing:0.8px;">Float</div>
                    <div style="font-size:14px;font-weight:800;color:{f_col};margin-top:2px;">
                        {(str(tf) + "d") if tf is not None else "-"}</div>
                </div>
                <div style="text-align:center;">
                    <div style="font-size:9px;color:#94A3B8;text-transform:uppercase;
                                letter-spacing:0.8px;">Risk</div>
                    <div style="background:{r_col};color:white;border-radius:8px;
                                padding:2px 10px;font-size:11px;font-weight:700;margin-top:2px;">
                        {rating}</div>
                </div>
            </div>
        </div>
        <div style="margin-top:10px;">
            <span style="background:{s_col};color:white;padding:2px 9px;
                         border-radius:10px;font-size:10px;">{status}</span>
        </div>
    </div>"""


def page_milestone_tracker(data: dict, near_crit_days: float):
    """
    Milestone Tracker page.
    Auto-detects milestones from the XER, allows manual additions,
    and shows driving path, successors, movement, notes and risk per milestone.
    """
    st.title("🏁 Milestone Tracker")
    st.caption(
        "Key programme milestones auto-detected from your XER. "
        "Select any milestone to see what is driving it, what it drives, and its current risk."
    )

    tasks = data["tasks_df"]
    rels  = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities found. Please upload a programme first.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    tasks_with_milestones = _detect_milestones(tasks)

    # Build graph
    G           = build_graph(tasks, rels)
    task_lookup = tasks.set_index("task_id").to_dict("index")

    # ---- Comparison data (finish movement) ----------------------------------
    movement_map = {}
    if "_mi_prev" in st.session_state and "_mi_curr" in st.session_state:
        try:
            prev_t = st.session_state["_mi_prev"]["tasks_df"]
            curr_t = st.session_state["_mi_curr"]["tasks_df"]
            if not prev_t.empty and not curr_t.empty:
                merged = prev_t[["task_code","eff_finish"]].merge(
                    curr_t[["task_code","eff_finish"]], on="task_code",
                    suffixes=("_p","_c"), how="inner"
                )
                for _, r in merged.iterrows():
                    try:
                        p = pd.Timestamp(r["eff_finish_p"])
                        c = pd.Timestamp(r["eff_finish_c"])
                        movement_map[str(r["task_code"])] = int((c - p).days)
                    except Exception:
                        pass
        except Exception:
            pass

    # ---- Notes data ---------------------------------------------------------
    notes_text = st.session_state.get("_notes_text", "")
    notes_ids  = set()
    risk_ids   = set()
    if notes_text and "task_code" in tasks.columns:
        for _, t in tasks.iterrows():
            code = str(t.get("task_code",""))
            if not code or code not in notes_text:
                continue
            notes_ids.add(code)
            idx = notes_text.find(code)
            snippet = notes_text[max(0,idx-300):idx+300]
            for word in _RISK_WORDS:
                if re.search(r'\b' + re.escape(word) + r'\b', snippet, re.IGNORECASE):
                    risk_ids.add(code)
                    break

    # ---- Milestone selection -------------------------------------------------
    auto_milestones = tasks_with_milestones[
        tasks_with_milestones["is_milestone_detected"]
    ].copy()

    # Session state for manually added milestones
    if "ms_manual_ids" not in st.session_state:
        st.session_state["ms_manual_ids"] = set()

    # ---- Sidebar-style controls panel ---------------------------------------
    with st.expander("Configure Milestones", expanded=True):
        cfg1, cfg2 = st.columns([1, 2])

        with cfg1:
            st.markdown(
                '<div style="font-size:11px;font-weight:700;color:#94A3B8;'
                'letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">'
                'Auto-Detected</div>',
                unsafe_allow_html=True,
            )
            st.caption(
                f"{len(auto_milestones)} milestones detected from activity type, "
                "zero duration, or name keywords."
            )

            show_auto = st.checkbox("Include auto-detected milestones", value=True, key="ms_show_auto")

        with cfg2:
            st.markdown(
                '<div style="font-size:11px;font-weight:700;color:#94A3B8;'
                'letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;">'
                'Add Key Activities Manually</div>',
                unsafe_allow_html=True,
            )
            manual_search = st.text_input(
                "Search to add",
                placeholder="Type Activity ID or name",
                key="ms_manual_search",
                label_visibility="collapsed",
            )
            if manual_search.strip():
                mask = (
                    tasks["task_code"].astype(str).str.contains(manual_search.strip(), case=False, na=False) |
                    tasks["task_name"].astype(str).str.contains(manual_search.strip(), case=False, na=False)
                )
                search_results = tasks[mask]
                if not search_results.empty:
                    add_labels = search_results.apply(
                        lambda r: f"{r.get('task_code','')}  --  {r.get('task_name','')}", axis=1
                    ).tolist()
                    add_sel = st.selectbox("Select to add", add_labels, key="ms_add_sel", label_visibility="collapsed")
                    if st.button("Add as milestone", key="ms_add_btn"):
                        sel_code = add_sel.split("  --  ")[0].strip()
                        match = tasks[tasks["task_code"] == sel_code]
                        if not match.empty:
                            st.session_state["ms_manual_ids"].add(match.iloc[0]["task_id"])
                            st.success(f"Added {sel_code}")
                else:
                    st.caption("No activities match.")

            if st.session_state["ms_manual_ids"]:
                if st.button("Clear manual selections", key="ms_clear"):
                    st.session_state["ms_manual_ids"] = set()

        # Filters
        st.markdown("<hr style='border:none;border-top:1px solid #E2E8F0;margin:10px 0;'>", unsafe_allow_html=True)
        fa, fb, fc = st.columns(3)
        f_risk = fa.selectbox("Risk filter", ["All","High","Medium","Low"], key="ms_f_risk")
        f_crit = fb.selectbox("Float filter", ["All","Critical","Near-Critical","Negative Float"], key="ms_f_crit")
        f_wbs  = fc.text_input("WBS contains", placeholder="e.g. Civil", key="ms_f_wbs")

    # ---- Build final milestone list -----------------------------------------
    milestone_ids = set()
    if show_auto:
        milestone_ids |= set(auto_milestones["task_id"].tolist())
    milestone_ids |= st.session_state["ms_manual_ids"]

    if not milestone_ids:
        st.info(
            "No milestones detected or selected. "
            "Either enable auto-detection above or search for and add activities manually."
        )
        return

    milestones = tasks[tasks["task_id"].isin(milestone_ids)].copy()

    # Apply risk rating
    def _rate(row):
        tf   = safe_float(row.get("total_float_days"), 9999)
        cstr = str(row.get("cstr_type","")).strip() not in ("","None","nan")
        risk = str(row.get("task_code","")) in risk_ids
        return _risk_rating(tf, cstr, risk)

    milestones["risk_rating"] = milestones.apply(_rate, axis=1)

    # Apply filters
    if f_risk != "All":
        milestones = milestones[milestones["risk_rating"] == f_risk]
    if f_crit == "Critical":
        milestones = milestones[milestones["is_critical"] == True]
    elif f_crit == "Near-Critical":
        milestones = milestones[milestones["is_near_critical"] == True]
    elif f_crit == "Negative Float":
        milestones = milestones[milestones["total_float_days"].apply(lambda f: safe_float(f,0) < 0)]
    if f_wbs.strip() and "wbs_path" in milestones.columns:
        milestones = milestones[milestones["wbs_path"].astype(str).str.contains(f_wbs.strip(), case=False, na=False)]

    if milestones.empty:
        st.info("No milestones match the current filters.")
        return

    milestones = milestones.sort_values("eff_finish" if "eff_finish" in milestones.columns else "task_code")

    # ---- Summary metrics ----------------------------------------------------
    n_total  = len(milestones)
    n_high   = int((milestones["risk_rating"] == "High").sum())
    n_med    = int((milestones["risk_rating"] == "Medium").sum())
    n_low    = int((milestones["risk_rating"] == "Low").sum())
    n_crit   = int(milestones["is_critical"].sum()) if "is_critical" in milestones.columns else 0
    n_neg    = int(milestones["total_float_days"].apply(lambda f: safe_float(f,0) < 0).sum()) if "total_float_days" in milestones.columns else 0

    st.markdown(
        f"""
        <div style="background:#0B1F33;border-radius:12px;padding:16px 22px;
                    margin-bottom:18px;display:flex;gap:20px;flex-wrap:wrap;align-items:center;">
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;">Milestones</div>
                <div style="font-size:28px;font-weight:800;color:#F5A623;line-height:1;margin-top:3px;">{n_total}</div>
            </div>
            <div style="width:1px;background:#1e3a5f;align-self:stretch;"></div>
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:1px;">High Risk</div>
                <div style="font-size:24px;font-weight:800;color:#dc2626;line-height:1;margin-top:3px;">{n_high}</div>
            </div>
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:1px;">Medium</div>
                <div style="font-size:24px;font-weight:800;color:#d97706;line-height:1;margin-top:3px;">{n_med}</div>
            </div>
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:1px;">Low</div>
                <div style="font-size:24px;font-weight:800;color:#16a34a;line-height:1;margin-top:3px;">{n_low}</div>
            </div>
            <div style="width:1px;background:#1e3a5f;align-self:stretch;"></div>
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:1px;">Critical</div>
                <div style="font-size:24px;font-weight:800;color:#dc2626;line-height:1;margin-top:3px;">{n_crit}</div>
            </div>
            <div>
                <div style="font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:1px;">Neg Float</div>
                <div style="font-size:24px;font-weight:800;color:#7f1d1d;line-height:1;margin-top:3px;">{n_neg}</div>
            </div>
            {"" if not movement_map else '<div style="margin-left:auto;font-size:11px;color:#F5A623;">Comparison data loaded</div>'}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Summary table tab + individual milestone tabs ----------------------
    tab_names = ["Summary Table"] + [
        f"{row.get('task_code','?')}" for _, row in milestones.head(10).iterrows()
    ]
    if len(milestones) > 10:
        tab_names[-1] = f"{tab_names[-1]}..."
    tabs = st.tabs(tab_names)

    # =========================================================================
    # TAB 0: Summary Table
    # =========================================================================
    with tabs[0]:
        st.markdown(
            '<div style="font-size:12px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
            'text-transform:uppercase;margin-bottom:8px;">All Milestones</div>',
            unsafe_allow_html=True,
        )

        table_rows = []
        for _, ms in milestones.iterrows():
            code = str(ms.get("task_code",""))
            tf   = safe_float(ms.get("total_float_days"), None)
            move = movement_map.get(code)
            table_rows.append({
                "Activity ID":      code,
                "Activity Name":    str(ms.get("task_name","")),
                "WBS":              str(ms.get("wbs_path","")).split(" > ")[0],
                "Forecast Start":   format_date(ms.get("eff_start")),
                "Forecast Finish":  format_date(ms.get("eff_finish")),
                "Total Float (d)":  round(tf, 1) if tf is not None else "-",
                "Critical":         "Yes" if ms.get("is_critical") else "",
                "Movement (d)":     move if move is not None else "-",
                "Status":           _status_label(str(ms.get("status",""))),
                "Risk":             ms.get("risk_rating",""),
                "In Notes":         "Yes" if code in notes_ids else "",
                "Risk Keywords":    "Yes" if code in risk_ids else "",
            })

        summary_df = pd.DataFrame(table_rows)

        def _style_summary(row):
            risk = row.get("Risk","")
            colour_map = {
                "High":   "background-color:#fef2f2;",
                "Medium": "background-color:#fffbeb;",
                "Low":    "",
            }
            return [colour_map.get(risk,"")] * len(row)

        st.dataframe(
            summary_df.style.apply(_style_summary, axis=1),
            use_container_width=True, hide_index=True,
        )

        # Timeline chart of all milestones
        gantt_src = milestones.dropna(subset=["eff_finish"]).copy() \
            if "eff_finish" in milestones.columns else pd.DataFrame()
        if not gantt_src.empty:
            gantt_src["Label"] = gantt_src["task_code"].astype(str) + "  " + gantt_src["task_name"].astype(str).str[:40]
            gantt_src["Risk"]  = gantt_src["risk_rating"]
            start_col  = "eff_start"  if "eff_start"  in gantt_src.columns else "eff_finish"
            finish_col = "eff_finish"

            # For zero-duration milestones, give a 1-day bar for visibility
            gantt_src["_plot_start"]  = gantt_src[start_col]
            gantt_src["_plot_finish"] = gantt_src.apply(
                lambda r: r[finish_col] + timedelta(days=1)
                if r.get(start_col) == r.get(finish_col) or pd.isna(r.get(start_col))
                else r[finish_col],
                axis=1,
            )

            fig_ms = px.timeline(
                gantt_src,
                x_start="_plot_start", x_end="_plot_finish", y="Label",
                color="Risk",
                color_discrete_map={
                    "High":   "#dc2626",
                    "Medium": "#d97706",
                    "Low":    "#16a34a",
                },
                title="Milestone Timeline",
                labels={"Label":""},
            )
            fig_ms.update_yaxes(autorange="reversed")
            fig_ms.add_vline(
                x=datetime.now(), line_dash="dash", line_color="#0B1F33",
                annotation_text="Today", annotation_position="top right",
                annotation=dict(font_color="#0B1F33"),
            )
            fig_ms.update_layout(
                height=max(300, min(800, 50 + len(gantt_src) * 28)),
                margin=dict(l=10,r=10,t=50,b=10),
                plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
                legend_title_text="Risk",
            )
            st.plotly_chart(fig_ms, use_container_width=True)

    # =========================================================================
    # TABS 1-N: Individual milestone detail
    # =========================================================================
    for tab_idx, (_, ms) in enumerate(milestones.head(10).iterrows(), start=1):
        if tab_idx >= len(tabs):
            break
        with tabs[tab_idx]:
            ms_code = str(ms.get("task_code",""))
            ms_id   = ms["task_id"]
            move    = movement_map.get(ms_code)

            # Header card
            st.markdown(
                _milestone_header_card(ms, movement_days=move),
                unsafe_allow_html=True,
            )

            # Constraint warning
            if "cstr_type" in ms.index:
                cstr = str(ms.get("cstr_type","")).strip()
                if cstr not in ("","None","nan"):
                    cdate = format_date(ms.get("cstr_date") if "cstr_date" in ms.index else None)
                    st.markdown(
                        f'<div style="background:#fffbeb;border-left:4px solid #F5A623;'
                        f'border-radius:6px;padding:10px 16px;margin-bottom:10px;">'
                        f'<strong>Constraint:</strong> {cstr} &nbsp;|&nbsp; '
                        f'<strong>Date:</strong> {cdate}</div>',
                        unsafe_allow_html=True,
                    )

            # Notes against this milestone
            if ms_code in notes_ids:
                note_flag = "Risk keywords found in notes." if ms_code in risk_ids else "Mentioned in planning notes."
                note_col  = "#dc2626" if ms_code in risk_ids else "#2563eb"
                st.markdown(
                    f'<div style="background:#eff6ff;border-left:4px solid {note_col};'
                    f'border-radius:6px;padding:10px 16px;margin-bottom:10px;">'
                    f'<strong>Planning Notes:</strong> {note_flag}</div>',
                    unsafe_allow_html=True,
                )

            # ---- Action buttons ---------------------------------------------
            ba, bb = st.columns(2)
            btn_drive = ba.button("Find Driving Path",  key=f"ms_drive_{ms_id}", use_container_width=True)
            btn_succ  = bb.button("Show Successors",    key=f"ms_succ_{ms_id}",  use_container_width=True)

            # ---- Driving Path ------------------------------------------------
            if btn_drive:
                with st.spinner("Tracing driving path..."):
                    direct_preds = list(G.predecessors(ms_id))
                    if not direct_preds:
                        st.warning(f"{ms_code} has no predecessors. No driving path can be identified.")
                    else:
                        path = driving_path_to_activity(G, tasks, rels, ms_id)
                        path_rows = []
                        for i, tid in enumerate(path):
                            t  = task_lookup.get(tid, {})
                            tf = t.get("total_float_days")
                            is_tgt = (tid == ms_id)
                            # Relationship to next step
                            rl, lg = "-", 0
                            if i < len(path) - 1:
                                next_id = path[i+1]
                                if not rels.empty:
                                    rel = rels[
                                        (rels.get("pred_task_id", pd.Series(dtype=str)) == tid) &
                                        (rels.get("succ_task_id", pd.Series(dtype=str)) == next_id)
                                    ]
                                    if not rel.empty:
                                        rl = _rel_label(rel["rel_type"].iloc[0] if "rel_type" in rel.columns else "FS")
                                        lg = safe_float(rel["lag_days"].iloc[0] if "lag_days" in rel.columns else 0, 0)
                            path_rows.append({
                                "Step":            i + 1,
                                "Activity ID":     t.get("task_code", tid),
                                "Activity Name":   t.get("task_name",""),
                                "Start":           format_date(t.get("eff_start")),
                                "Finish":          format_date(t.get("eff_finish")),
                                "Float (d)":       round(float(tf),1) if tf is not None else "-",
                                "Link":            rl if not is_tgt else "-",
                                "Lag (d)":         lg if not is_tgt else "-",
                                "Flag":            _crit_flag(tf),
                                "Milestone":       "TARGET" if is_tgt else "",
                            })

                        path_df = pd.DataFrame(path_rows)
                        st.markdown("**Driving Path**")
                        st.caption("Activities ordered from chain start to the milestone target.")

                        def _path_row_style(row):
                            flag = row.get("Flag","")
                            is_t = row.get("Milestone","") == "TARGET"
                            if is_t:
                                return ["background-color:#0B1F33;color:white;font-weight:700;"] * len(row)
                            cm = {"Critical":"background-color:#fee2e2;",
                                  "Negative Float":"background-color:#fecaca;",
                                  "Near-Critical":"background-color:#fef3c7;"}
                            return [cm.get(flag,"")] * len(row)

                        st.dataframe(path_df.style.apply(_path_row_style, axis=1),
                                     use_container_width=True, hide_index=True)

                        # Mini Gantt for path
                        path_task_ids = [p for p in path if p != ms_id]
                        path_tasks    = tasks[tasks["task_id"].isin(path)].copy()
                        gantt_p = path_tasks.dropna(subset=["eff_start","eff_finish"]).copy() \
                            if "eff_start" in path_tasks.columns else pd.DataFrame()
                        if not gantt_p.empty:
                            gantt_p["Label"] = gantt_p["task_code"].astype(str) + "  " + gantt_p["task_name"].astype(str).str[:35]
                            gantt_p["Type"]  = gantt_p["task_id"].apply(lambda t: "Milestone" if t == ms_id else "Driving Path")
                            fig_dp = px.timeline(
                                gantt_p, x_start="eff_start", x_end="eff_finish", y="Label",
                                color="Type",
                                color_discrete_map={"Milestone":"#0B1F33","Driving Path":"#dc2626"},
                                title=f"Driving Path to {ms_code}",
                            )
                            fig_dp.update_yaxes(autorange="reversed")
                            fig_dp.add_vline(x=datetime.now(), line_dash="dot", line_color="#94A3B8",
                                             annotation_text="Today")
                            fig_dp.update_layout(height=max(250, 50+len(gantt_p)*28),
                                                 margin=dict(l=10,r=10,t=40,b=10),
                                                 plot_bgcolor="#ffffff", paper_bgcolor="#ffffff")
                            st.plotly_chart(fig_dp, use_container_width=True)

            # ---- Successors --------------------------------------------------
            if btn_succ:
                direct_succs = list(G.successors(ms_id))
                if not direct_succs:
                    st.warning(f"{ms_code} has no successors.")
                else:
                    all_succs = trace_successors(G, ms_id)
                    succ_df   = _build_full_trace_df(G, rels, task_lookup, ms_id, all_succs, "succ")
                    # Add WBS
                    if not succ_df.empty and "Activity ID" in succ_df.columns:
                        code_to_wbs = tasks.set_index("task_code")["wbs_path"].to_dict() if "wbs_path" in tasks.columns else {}
                        succ_df.insert(succ_df.columns.get_loc("Activity Name")+1, "WBS",
                                       succ_df["Activity ID"].map(code_to_wbs).fillna("-"))

                    st.markdown(f"**All Successors of {ms_code}** ({len(succ_df)} activities)")

                    def _succ_style(val):
                        return {"Critical":"background-color:#fee2e2;color:#991b1b;font-weight:600;",
                                "Negative Float":"background-color:#fecaca;color:#7f1d1d;font-weight:700;",
                                "Near-Critical":"background-color:#fef3c7;color:#92400e;font-weight:600;",
                                "Float":"background-color:#dcfce7;color:#166534;"}.get(val,"")

                    st.dataframe(
                        succ_df.style.applymap(_succ_style, subset=["Critical Flag"]),
                        use_container_width=True, hide_index=True, height=min(400, 45+len(succ_df)*35),
                    )

            # ---- Individual export -------------------------------------------
            st.divider()
            exp_col, _ = st.columns([1,3])

            exp_detail = pd.DataFrame([{
                "Activity ID":    ms_code,
                "Activity Name":  str(ms.get("task_name","")),
                "WBS":            str(ms.get("wbs_path","")),
                "Forecast Start": format_date(ms.get("eff_start")),
                "Forecast Finish":format_date(ms.get("eff_finish")),
                "Total Float (d)":safe_float(ms.get("total_float_days"), None),
                "Critical":       "Yes" if ms.get("is_critical") else "No",
                "Movement (d)":   move if move is not None else "N/A",
                "Status":         _status_label(str(ms.get("status",""))),
                "Risk Rating":    ms.get("risk_rating",""),
                "In Notes":       "Yes" if ms_code in notes_ids else "No",
                "Risk Keywords":  "Yes" if ms_code in risk_ids else "No",
                "Constraint":     str(ms.get("cstr_type","")).strip() or "None",
            }])

            xls_ms = export_df_to_excel({"Milestone Detail": exp_detail})
            exp_col.download_button(
                label=f"📥 Export {ms_code} Report",
                data=xls_ms,
                file_name=f"milestone_{ms_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"ms_exp_{ms_id}",
            )

    # =========================================================================
    # FULL EXPORT
    # =========================================================================
    st.divider()

    full_exp_rows = []
    for _, ms in milestones.iterrows():
        code = str(ms.get("task_code",""))
        tf   = safe_float(ms.get("total_float_days"), None)
        full_exp_rows.append({
            "Activity ID":      code,
            "Activity Name":    str(ms.get("task_name","")),
            "WBS":              str(ms.get("wbs_path","")),
            "Forecast Start":   format_date(ms.get("eff_start")),
            "Forecast Finish":  format_date(ms.get("eff_finish")),
            "Total Float (d)":  round(tf,1) if tf is not None else "-",
            "Critical":         "Yes" if ms.get("is_critical") else "",
            "Risk Rating":      ms.get("risk_rating",""),
            "Movement (d)":     movement_map.get(code, "-"),
            "Status":           _status_label(str(ms.get("status",""))),
            "In Notes":         "Yes" if code in notes_ids else "",
            "Risk Keywords":    "Yes" if code in risk_ids else "",
            "Constraint":       str(ms.get("cstr_type","")).strip() or "",
        })

    full_df = pd.DataFrame(full_exp_rows)
    high_df = full_df[full_df["Risk Rating"] == "High"]
    med_df  = full_df[full_df["Risk Rating"] == "Medium"]

    xls_full = export_df_to_excel({
        "All Milestones": full_df,
        "High Risk":      high_df if not high_df.empty else pd.DataFrame(columns=["No data"]),
        "Medium Risk":    med_df  if not med_df.empty  else pd.DataFrame(columns=["No data"]),
    })

    dl_col, _ = st.columns([1,3])
    dl_col.download_button(
        label="📥  Export All Milestones to Excel",
        data=xls_full,
        file_name=f"milestones_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Exports All Milestones, High Risk and Medium Risk sheets.",
    )




# -----------------------------------------------------------------------------
# PAGE: RISK & OPPORTUNITY REGISTER
# -----------------------------------------------------------------------------

# Probability and impact scoring for auto-generated items
_PROB_MAP   = {"High": "High", "Medium": "Medium", "Low": "Low"}
_IMPACT_MAP = {"High": "High", "Medium": "Medium", "Low": "Low"}

# RAG colour helpers reused from rest of app
def _rag_colour(priority: str) -> tuple:
    """Return (text_colour, bg_colour, border_colour) for a priority level."""
    return {
        "High":   ("#991b1b", "#fef2f2", "#fca5a5"),
        "Medium": ("#92400e", "#fffbeb", "#fcd34d"),
        "Low":    ("#166534", "#f0fdf4", "#86efac"),
    }.get(priority, ("#374151", "#f9fafb", "#E2E8F0"))


def _ro_card_pill(item_type: str) -> str:
    """Coloured pill for Risk vs Opportunity."""
    if item_type == "Risk":
        return (
            '<span style="background:#dc2626;color:white;padding:2px 10px;'
            'border-radius:10px;font-size:10px;font-weight:700;letter-spacing:0.5px;">'
            'RISK</span>'
        )
    return (
        '<span style="background:#16a34a;color:white;padding:2px 10px;'
        'border-radius:10px;font-size:10px;font-weight:700;letter-spacing:0.5px;">'
        'OPPORTUNITY</span>'
    )


def _ro_row(item_type, priority, act_code, act_name, wbs,
            description, cause, effect, mitigation):
    """Build a single register row dict."""
    return {
        "Type":            item_type,
        "Priority":        priority,
        "Activity ID":     str(act_code),
        "Activity Name":   str(act_name),
        "WBS":             str(wbs),
        "Description":     str(description),
        "Cause":           str(cause),
        "Effect":          str(effect),
        "Mitigation / Action": str(mitigation),
        "Owner":           "",
        "Due Date":        "",
        "Status":          "Open",
    }


def _generate_register(
    tasks: pd.DataFrame,
    rels:  pd.DataFrame,
    near_crit_days: float,
    notes_text: str = "",
) -> pd.DataFrame:
    """
    Generate a draft Risk & Opportunity register from programme data.
    Returns a DataFrame with one row per item.
    """
    rows = []
    now  = datetime.now()
    eight_weeks = now + timedelta(weeks=8)
    tasks = get_critical_threshold(tasks, near_crit_days)

    def _wbs(row):
        w = row.get("wbs_path","") if "wbs_path" in row.index else ""
        return str(w).split(" > ")[0] if w and str(w).strip() not in ("","nan") else "-"

    # -- Predecessor / successor lookup ----------------------------------------
    tasks_with_pred = set()
    tasks_with_succ = set()
    if not rels.empty:
        if "succ_task_id" in rels.columns:
            tasks_with_pred = set(rels["succ_task_id"].dropna())
        if "pred_task_id" in rels.columns:
            tasks_with_succ = set(rels["pred_task_id"].dropna())

    # =========================================================================
    # RISKS
    # =========================================================================

    # R1: Negative float
    if "total_float_days" in tasks.columns:
        neg = tasks[tasks["total_float_days"].apply(lambda f: safe_float(f,0) < 0)]
        for _, t in neg.iterrows():
            tf = round(safe_float(t.get("total_float_days"),0), 1)
            rows.append(_ro_row(
                "Risk", "High",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                f"Activity has {abs(tf)} days negative float.",
                f"The current schedule logic and constraints result in a {abs(tf)}-day overrun on this activity.",
                "The project completion date cannot be achieved on this path without intervention. "
                "Every day of further delay worsens the position.",
                f"Raise with the planner immediately. Develop a recovery plan to recover the {abs(tf)} days. "
                "Consider acceleration, scope reduction, or parallel working.",
            ))

    # R2: Critical activities not started
    if "status" in tasks.columns and "is_critical" in tasks.columns:
        crit_ns = tasks[
            tasks["is_critical"] &
            tasks["status"].apply(lambda s: str(s) in ("TK_NotStart","Not Started"))
        ]
        for _, t in crit_ns.iterrows():
            finish = format_date(t.get("eff_finish"))
            rows.append(_ro_row(
                "Risk", "High",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                f"Critical activity not yet started. Target finish: {finish}.",
                "The activity is on the critical path but mobilisation or enabling works have not commenced.",
                "Any further delay to this activity will directly delay the project finish date.",
                "Confirm the start date and mobilisation plan. If at risk, escalate to the delivery team and PM.",
            ))

    # R3: Near-critical due within 8 weeks
    if "eff_finish" in tasks.columns and "is_near_critical" in tasks.columns:
        nc_soon = tasks[
            tasks["is_near_critical"] &
            tasks["eff_finish"].apply(
                lambda d: d is not None and hasattr(d,"date") and d <= eight_weeks
            )
        ]
        for _, t in nc_soon.iterrows():
            tf = round(safe_float(t.get("total_float_days"),0), 1)
            rows.append(_ro_row(
                "Risk", "Medium",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                f"Near-critical activity due within 8 weeks with only {tf} days float.",
                "Limited schedule buffer combined with imminent completion requirements.",
                "If this activity is delayed, it will move onto the critical path and threaten the project finish.",
                f"Monitor weekly. If float drops below 5 days, treat as critical and escalate.",
            ))

    # R4: No predecessor (open start)
    if "task_id" in tasks.columns and "task_type" in tasks.columns:
        no_pred = tasks[
            ~tasks["task_id"].isin(tasks_with_pred) &
            ~tasks["task_type"].astype(str).str.contains("Milestone|LOE|WBS", na=False)
        ]
        for _, t in no_pred.head(15).iterrows():
            rows.append(_ro_row(
                "Risk", "Medium",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                "Activity has no predecessor. Programme logic is open at the start.",
                "Missing or incomplete logic in the schedule. The activity may start earlier or later than intended.",
                "Float calculations for this activity may be unreliable. "
                "The activity could be delayed without any schedule warning.",
                "Review with the planner. Add a logical predecessor or add a constraint if the start date is fixed.",
            ))

    # R5: No successor (open finish)
    if "task_id" in tasks.columns and "task_type" in tasks.columns:
        no_succ = tasks[
            ~tasks["task_id"].isin(tasks_with_succ) &
            ~tasks["task_type"].astype(str).str.contains("Finish Milestone|LOE|WBS", na=False)
        ]
        for _, t in no_succ.head(15).iterrows():
            rows.append(_ro_row(
                "Risk", "Medium",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                "Activity has no successor. Programme logic is open at the finish.",
                "Missing logic means this activity does not drive any subsequent work in the schedule.",
                "The activity may show artificially high float. "
                "Delays may not cascade correctly through the programme.",
                "Review with the planner. Add a logical successor or confirm the activity is a deliberate end point.",
            ))

    # R6: Excessive lag (> 10 days)
    if not rels.empty and "lag_days" in rels.columns:
        big_lag = rels[rels["lag_days"].apply(lambda l: safe_float(l,0) > 10)]
        for _, r in big_lag.head(10).iterrows():
            code = r.get("succ_task_code", r.get("succ_task_id",""))
            name = r.get("succ_task_name","")
            pred = r.get("pred_task_code", r.get("pred_task_id",""))
            lag  = int(safe_float(r.get("lag_days",0),0))
            match = tasks[tasks["task_code"] == str(code)]
            wbs   = _wbs(match.iloc[0]) if not match.empty else "-"
            rows.append(_ro_row(
                "Risk", "Low",
                code, name, wbs,
                f"Relationship from {pred} has {lag} days lag.",
                f"A {lag}-day lag has been applied instead of modelling the actual work sequence.",
                "Excessive lag disguises schedule risk and inflates float on successor activities. "
                "It may also mask negative float.",
                f"Challenge the {lag}-day lag with the planner. Replace with a properly sequenced "
                "activity or reduce the lag to the minimum justified by contract or site conditions.",
            ))

    # R7: Long duration activities (> 60 days)
    if "orig_dur_days" in tasks.columns:
        long_dur = tasks[tasks["orig_dur_days"].apply(lambda d: safe_float(d,0) > 60)]
        for _, t in long_dur.head(10).iterrows():
            dur = int(safe_float(t.get("orig_dur_days",0),0))
            rows.append(_ro_row(
                "Risk", "Low",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                f"Activity duration is {dur} working days.",
                "Activity is too long to manage or monitor effectively as a single work package.",
                "Problems can go undetected for weeks. Float burn and delays may not surface until it is too late.",
                f"Break the {dur}-day activity into smaller work packages of 20-30 days. "
                "Discuss with the planner to improve schedule resolution.",
            ))

    # R8: Planning notes risk words
    if notes_text and "task_code" in tasks.columns:
        for word in _RISK_WORDS:
            pattern = re.compile(r'\b' + re.escape(word) + r'\b', re.IGNORECASE)
            if not pattern.search(notes_text):
                continue
            for _, t in tasks.iterrows():
                code = str(t.get("task_code",""))
                if not code or code not in notes_text:
                    continue
                idx = notes_text.find(code)
                snippet = notes_text[max(0,idx-300):idx+300]
                if pattern.search(snippet):
                    rows.append(_ro_row(
                        "Risk", "High",
                        code, t.get("task_name",""), _wbs(t),
                        f"Planning notes reference '{word}' against this activity.",
                        f"The planning note indicates a potential {word} issue that is not fully visible in the programme.",
                        "This risk may not be reflected in the current schedule float or critical path.",
                        f"Review the planning note for {code}. Confirm whether the '{word}' item has been "
                        "resolved, raise a formal risk if not, and update the programme if dates are affected.",
                    ))
                    break

    # R9: Activities delayed in comparison (> 10 days slip)
    if "_mi_prev" in st.session_state and "_mi_curr" in st.session_state:
        try:
            prev_t = st.session_state["_mi_prev"]["tasks_df"]
            curr_t = st.session_state["_mi_curr"]["tasks_df"]
            if not prev_t.empty and not curr_t.empty:
                merged_comp = prev_t[["task_code","eff_finish","task_name"]].merge(
                    curr_t[["task_code","eff_finish"]], on="task_code", suffixes=("_p","_c"), how="inner"
                )
                for _, r in merged_comp.iterrows():
                    try:
                        slip = int((pd.Timestamp(r["eff_finish_c"]) - pd.Timestamp(r["eff_finish_p"])).days)
                        if slip > 10:
                            match = tasks[tasks["task_code"] == str(r["task_code"])]
                            wbs   = _wbs(match.iloc[0]) if not match.empty else "-"
                            rows.append(_ro_row(
                                "Risk", "High" if slip > 30 else "Medium",
                                str(r["task_code"]), str(r.get("task_name","")), wbs,
                                f"Activity finish date has slipped {slip} days since the previous programme revision.",
                                f"The activity's forecast finish moved {slip} days later between the two programme versions.",
                                f"A {slip}-day slip on this activity may delay downstream work and impact the project finish date.",
                                f"Investigate the reason for the {slip}-day slip. Agree a recovery programme with the delivery team. "
                                "Update the programme and issue a revised schedule.",
                            ))
                    except Exception:
                        pass
        except Exception:
            pass

    # =========================================================================
    # OPPORTUNITIES
    # =========================================================================

    # O1: Activities pulled earlier in comparison (> 5 days improvement)
    if "_mi_prev" in st.session_state and "_mi_curr" in st.session_state:
        try:
            prev_t = st.session_state["_mi_prev"]["tasks_df"]
            curr_t = st.session_state["_mi_curr"]["tasks_df"]
            if not prev_t.empty and not curr_t.empty:
                merged_opp = prev_t[["task_code","eff_finish","task_name"]].merge(
                    curr_t[["task_code","eff_finish"]], on="task_code", suffixes=("_p","_c"), how="inner"
                )
                for _, r in merged_opp.iterrows():
                    try:
                        gain = int((pd.Timestamp(r["eff_finish_p"]) - pd.Timestamp(r["eff_finish_c"])).days)
                        if gain > 5:
                            match = tasks[tasks["task_code"] == str(r["task_code"])]
                            wbs   = _wbs(match.iloc[0]) if not match.empty else "-"
                            rows.append(_ro_row(
                                "Opportunity", "Medium",
                                str(r["task_code"]), str(r.get("task_name","")), wbs,
                                f"Activity finish date has improved by {gain} days since the previous revision.",
                                f"The activity's forecast finish moved {gain} days earlier between programme versions.",
                                f"This improvement may allow successor activities to start earlier and could reduce the overall project duration.",
                                f"Review whether the {gain}-day improvement can be formally recognised in the programme. "
                                "Check if successor activities can be brought forward accordingly.",
                            ))
                    except Exception:
                        pass
        except Exception:
            pass

    # O2: High float activities that could be resequenced (float > 30 days)
    if "total_float_days" in tasks.columns:
        high_float = tasks[tasks["total_float_days"].apply(lambda f: safe_float(f,0) > 30)]
        for _, t in high_float.head(10).iterrows():
            tf = int(safe_float(t.get("total_float_days"),0))
            rows.append(_ro_row(
                "Opportunity", "Low",
                t.get("task_code",""), t.get("task_name",""), _wbs(t),
                f"Activity has {tf} days total float. There may be scope to resequence.",
                f"Significant float of {tf} days suggests this activity's resources or timing could be optimised.",
                "The float window could be used to smooth resource demand, resolve clashes, or accelerate predecessor activities.",
                f"Review with the planner whether the {tf} days of float can be used to resource-level, "
                "front-load critical activities, or release resource to other packages.",
            ))

    # O3: Float gained in comparison
    if "_mi_prev" in st.session_state and "_mi_curr" in st.session_state:
        try:
            prev_t = st.session_state["_mi_prev"]["tasks_df"]
            curr_t = st.session_state["_mi_curr"]["tasks_df"]
            if not prev_t.empty and not curr_t.empty and \
               "total_float_days" in prev_t.columns and "total_float_days" in curr_t.columns:
                mf = prev_t[["task_code","total_float_days","task_name"]].merge(
                    curr_t[["task_code","total_float_days"]], on="task_code", suffixes=("_p","_c"), how="inner"
                )
                for _, r in mf.iterrows():
                    fp = safe_float(r.get("total_float_days_p"), None)
                    fc = safe_float(r.get("total_float_days_c"), None)
                    if fp is not None and fc is not None and fc - fp > 10:
                        match = tasks[tasks["task_code"] == str(r["task_code"])]
                        wbs   = _wbs(match.iloc[0]) if not match.empty else "-"
                        gain  = round(fc - fp, 1)
                        rows.append(_ro_row(
                            "Opportunity", "Low",
                            str(r["task_code"]), str(r.get("task_name","")), wbs,
                            f"Activity gained {gain} days float since the previous revision.",
                            "Schedule logic or date changes in the current revision have created additional float.",
                            "This float could be used for resource smoothing, risk buffer, or to accommodate other priorities.",
                            f"Review whether the additional {gain} days of float has been deliberately created or is an unintended consequence. "
                            "Confirm with the planner that this aligns with programme strategy.",
                        ))
        except Exception:
            pass

    # O4: Near-critical activities where predecessors have high float
    if "total_float_days" in tasks.columns and not rels.empty:
        nc_acts = tasks[tasks["is_near_critical"]] if "is_near_critical" in tasks.columns else pd.DataFrame()
        if not nc_acts.empty and "pred_task_id" in rels.columns:
            task_float_map = tasks.set_index("task_id")["total_float_days"].to_dict() if "task_id" in tasks.columns else {}
            for _, t in nc_acts.head(10).iterrows():
                tid = t.get("task_id","")
                pred_ids = rels[rels["succ_task_id"] == tid]["pred_task_id"].tolist() if "succ_task_id" in rels.columns else []
                high_float_preds = [
                    p for p in pred_ids
                    if safe_float(task_float_map.get(p), 0) > 20
                ]
                if high_float_preds:
                    tf = round(safe_float(t.get("total_float_days"),0), 1)
                    rows.append(_ro_row(
                        "Opportunity", "Medium",
                        t.get("task_code",""), t.get("task_name",""), _wbs(t),
                        f"Near-critical activity ({tf}d float) has predecessors with high float.",
                        f"One or more predecessor activities have significant float, meaning they could potentially "
                        "be accelerated without impacting the overall programme.",
                        f"Accelerating high-float predecessors could create additional buffer on this near-critical path.",
                        f"Review whether the high-float predecessors can be started earlier or completed faster "
                        "to increase the float buffer on this activity.",
                    ))

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset=["Type","Activity ID","Description"]).reset_index(drop=True)

    # Sort: Risks first, then Opportunities; within each by Priority
    priority_order = {"High": 0, "Medium": 1, "Low": 2}
    type_order     = {"Risk": 0, "Opportunity": 1}
    df["_ts"] = df["Type"].map(type_order)
    df["_ps"] = df["Priority"].map(priority_order)
    df = df.sort_values(["_ts","_ps","Activity ID"]).drop(columns=["_ts","_ps"]).reset_index(drop=True)

    return df


def page_risk_register(data: dict, near_crit_days: float):
    """
    Risk & Opportunity Register page.
    Auto-generates a draft register from programme data with
    editable owner, due date and status fields.
    """
    st.title("⚠️ Risk & Opportunity Register")
    st.caption(
        "Auto-generated draft register based on the uploaded programme. "
        "Edit Owner, Due Date and Status inline. Export when complete."
    )

    tasks = data["tasks_df"]
    rels  = data["relationships_df"]

    if tasks.empty:
        st.warning("No activities found. Please upload a programme first.")
        return

    notes_text = st.session_state.get("_notes_text", "")

    # ---- Generate / cache register ------------------------------------------
    prog_key  = st.session_state.get("_xer_cache_key", "")
    cache_key = f"_ro_register_{prog_key}_{near_crit_days}"

    if st.session_state.get("_ro_register_key") != cache_key:
        with st.spinner("Analysing programme for risks and opportunities..."):
            register_df = _generate_register(tasks, rels, near_crit_days, notes_text)
        st.session_state["_ro_register_df"]  = register_df
        st.session_state["_ro_register_key"] = cache_key
    else:
        register_df = st.session_state["_ro_register_df"]

    if register_df.empty:
        st.success("No risks or opportunities generated from the current programme data.")
        return

    # ---- Summary banner -----------------------------------------------------
    n_total = len(register_df)
    n_risk  = int((register_df["Type"] == "Risk").sum())
    n_opp   = int((register_df["Type"] == "Opportunity").sum())
    n_high  = int((register_df["Priority"] == "High").sum())
    n_med   = int((register_df["Priority"] == "Medium").sum())
    n_low   = int((register_df["Priority"] == "Low").sum())

    # Risk counts by priority
    r_high = int(((register_df["Type"]=="Risk") & (register_df["Priority"]=="High")).sum())
    r_med  = int(((register_df["Type"]=="Risk") & (register_df["Priority"]=="Medium")).sum())
    r_low  = int(((register_df["Type"]=="Risk") & (register_df["Priority"]=="Low")).sum())
    o_high = int(((register_df["Type"]=="Opportunity") & (register_df["Priority"]=="High")).sum())
    o_med  = int(((register_df["Type"]=="Opportunity") & (register_df["Priority"]=="Medium")).sum())
    o_low  = int(((register_df["Type"]=="Opportunity") & (register_df["Priority"]=="Low")).sum())

    st.markdown(
        f"""
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px;">
            <div style="background:#0B1F33;border-radius:12px;padding:16px 20px;">
                <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;margin-bottom:8px;">Risks  ({n_risk})</div>
                <div style="display:flex;gap:12px;">
                    <div style="text-align:center;">
                        <div style="font-size:9px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;">High</div>
                        <div style="font-size:24px;font-weight:800;color:#dc2626;line-height:1.1;">{r_high}</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:9px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;">Medium</div>
                        <div style="font-size:24px;font-weight:800;color:#d97706;line-height:1.1;">{r_med}</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:9px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;">Low</div>
                        <div style="font-size:24px;font-weight:800;color:#94A3B8;line-height:1.1;">{r_low}</div>
                    </div>
                </div>
            </div>
            <div style="background:#0B1F33;border-radius:12px;padding:16px 20px;">
                <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;
                            letter-spacing:1px;margin-bottom:8px;">Opportunities  ({n_opp})</div>
                <div style="display:flex;gap:12px;">
                    <div style="text-align:center;">
                        <div style="font-size:9px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;">High</div>
                        <div style="font-size:24px;font-weight:800;color:#16a34a;line-height:1.1;">{o_high}</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:9px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;">Medium</div>
                        <div style="font-size:24px;font-weight:800;color:#16a34a;line-height:1.1;">{o_med}</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:9px;color:#64748B;text-transform:uppercase;letter-spacing:0.8px;">Low</div>
                        <div style="font-size:24px;font-weight:800;color:#94A3B8;line-height:1.1;">{o_low}</div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Filters ------------------------------------------------------------
    with st.expander("Filter register", expanded=False):
        f1, f2, f3, f4 = st.columns(4)

        f_type   = f1.selectbox("Type",     ["All","Risk","Opportunity"], key="ro_f_type")
        f_pri    = f2.selectbox("Priority", ["All","High","Medium","Low"],  key="ro_f_pri")
        f_status = f3.selectbox("Status",   ["All","Open","In Progress","Closed"], key="ro_f_status")

        all_wbs  = ["All"] + sorted(register_df["WBS"].unique().tolist())
        f_wbs    = f4.selectbox("WBS",      all_wbs, key="ro_f_wbs")

    filtered = register_df.copy()
    if f_type   != "All": filtered = filtered[filtered["Type"]     == f_type]
    if f_pri    != "All": filtered = filtered[filtered["Priority"] == f_pri]
    if f_status != "All": filtered = filtered[filtered["Status"]   == f_status]
    if f_wbs    != "All": filtered = filtered[filtered["WBS"]      == f_wbs]

    st.caption(f"Showing {len(filtered)} of {n_total} items.")

    # ---- Tabs: Risks | Opportunities | Full Register | Export ----------------
    tab_risks, tab_opps, tab_full, tab_export = st.tabs([
        f"Risks ({n_risk})",
        f"Opportunities ({n_opp})",
        "Full Register",
        "Export",
    ])

    EDIT_COLS = [
        "Type","Priority","Activity ID","Activity Name","WBS",
        "Description","Cause","Effect","Mitigation / Action",
        "Owner","Due Date","Status",
    ]

    COL_CONFIG = {
        "Type": st.column_config.SelectboxColumn(
            "Type", options=["Risk","Opportunity"], width="small"
        ),
        "Priority": st.column_config.SelectboxColumn(
            "Priority", options=["High","Medium","Low"], width="small"
        ),
        "Status": st.column_config.SelectboxColumn(
            "Status", options=["Open","In Progress","Closed"], width="small"
        ),
        "Owner":              st.column_config.TextColumn("Owner",    width="small"),
        "Due Date":           st.column_config.TextColumn("Due Date", width="small"),
        "Activity ID":        st.column_config.TextColumn("Activity ID",   width="small"),
        "Activity Name":      st.column_config.TextColumn("Activity Name", width="medium"),
        "WBS":                st.column_config.TextColumn("WBS",           width="medium"),
        "Description":        st.column_config.TextColumn("Description",   width="large"),
        "Cause":              st.column_config.TextColumn("Cause",         width="large"),
        "Effect":             st.column_config.TextColumn("Effect",        width="large"),
        "Mitigation / Action":st.column_config.TextColumn("Mitigation / Action", width="large"),
    }

    def _style_row(row):
        """Row background by Type and Priority."""
        if row.get("Type","") == "Opportunity":
            return ["background-color:#f0fdf4;"] * len(row)
        priority = row.get("Priority","")
        cm = {
            "High":   "background-color:#fef2f2;",
            "Medium": "background-color:#fffbeb;",
        }
        return [cm.get(priority,"")] * len(row)

    def _show_editable(df, key_suffix):
        avail = [c for c in EDIT_COLS if c in df.columns]
        edited = st.data_editor(
            df[avail].style.apply(_style_row, axis=1),
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config=COL_CONFIG,
            key=f"ro_editor_{key_suffix}",
        )
        return edited

    # -- Risks tab --------------------------------------------------------------
    with tab_risks:
        risks_df = filtered[filtered["Type"] == "Risk"].copy()
        if risks_df.empty:
            st.success("No risks match the current filters.")
        else:
            st.markdown(
                '<div style="font-size:12px;font-weight:700;color:#94A3B8;letter-spacing:1px;'
                'text-transform:uppercase;margin-bottom:8px;">Risks</div>',
                unsafe_allow_html=True,
            )
            st.caption("Edit Owner, Due Date and Status directly in the table.")
            edited_risks = _show_editable(risks_df, "risks")

            # Persist edits
            if edited_risks is not None and not edited_risks.empty:
                for col in ["Owner","Due Date","Status","Priority"]:
                    if col in edited_risks.columns:
                        register_df.loc[risks_df.index, col] = edited_risks[col].values
                st.session_state["_ro_register_df"] = register_df

            # High priority cards
            high_risks = risks_df[risks_df["Priority"] == "High"]
            if not high_risks.empty:
                st.markdown("---")
                st.markdown(
                    f'<div style="font-size:12px;font-weight:700;color:#dc2626;letter-spacing:1px;'
                    f'text-transform:uppercase;margin-bottom:10px;">'
                    f'High Priority Risks ({len(high_risks)})</div>',
                    unsafe_allow_html=True,
                )
                for _, item in high_risks.iterrows():
                    tc, bc, brc = _rag_colour("High")
                    st.markdown(
                        f"""
                        <div style="background:{bc};border:1px solid {brc};border-left:5px solid #dc2626;
                                    border-radius:8px;padding:14px 18px;margin-bottom:8px;">
                            <div style="display:flex;align-items:flex-start;justify-content:space-between;
                                        gap:12px;flex-wrap:wrap;">
                                <div style="flex:1;min-width:0;">
                                    {_ro_card_pill("Risk")}
                                    <div style="font-weight:700;color:#0B1F33;font-size:14px;margin-top:6px;">
                                        {item.get("Activity ID","")} - {item.get("Activity Name","")}
                                    </div>
                                    <div style="font-size:12px;color:#64748B;margin-top:2px;">
                                        {item.get("WBS","")}
                                    </div>
                                </div>
                            </div>
                            <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
                                <div>
                                    <div style="font-size:10px;font-weight:700;color:#94A3B8;
                                                text-transform:uppercase;letter-spacing:0.8px;">Description</div>
                                    <div style="font-size:13px;color:#334155;margin-top:3px;">{item.get("Description","")}</div>
                                </div>
                                <div>
                                    <div style="font-size:10px;font-weight:700;color:#94A3B8;
                                                text-transform:uppercase;letter-spacing:0.8px;">Effect</div>
                                    <div style="font-size:13px;color:#334155;margin-top:3px;">{item.get("Effect","")}</div>
                                </div>
                                <div>
                                    <div style="font-size:10px;font-weight:700;color:#94A3B8;
                                                text-transform:uppercase;letter-spacing:0.8px;">Mitigation</div>
                                    <div style="font-size:13px;color:#334155;margin-top:3px;">{item.get("Mitigation / Action","")}</div>
                                </div>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

    # -- Opportunities tab ------------------------------------------------------
    with tab_opps:
        opps_df = filtered[filtered["Type"] == "Opportunity"].copy()
        if opps_df.empty:
            st.info("No opportunities detected from the current programme data.")
            if "_mi_prev" not in st.session_state:
                st.caption(
                    "Note: Comparison-based opportunities (pulled-forward activities, float gained) "
                    "require both previous and current XER files to be uploaded on the "
                    "Programme Comparison page."
                )
        else:
            st.markdown(
                '<div style="font-size:12px;font-weight:700;color:#16a34a;letter-spacing:1px;'
                'text-transform:uppercase;margin-bottom:8px;">Opportunities</div>',
                unsafe_allow_html=True,
            )
            st.caption("Review these opportunities with the planner to see if they can be captured.")
            edited_opps = _show_editable(opps_df, "opps")

            if edited_opps is not None and not edited_opps.empty:
                for col in ["Owner","Due Date","Status","Priority"]:
                    if col in edited_opps.columns:
                        register_df.loc[opps_df.index, col] = edited_opps[col].values
                st.session_state["_ro_register_df"] = register_df

    # -- Full register tab ------------------------------------------------------
    with tab_full:
        st.caption("Complete register -- all risks and opportunities combined.")
        edited_full = _show_editable(filtered, "full")
        if edited_full is not None and not edited_full.empty:
            for col in ["Owner","Due Date","Status","Priority"]:
                if col in edited_full.columns:
                    register_df.loc[filtered.index, col] = edited_full[col].values
            st.session_state["_ro_register_df"] = register_df

    # -- Export tab -------------------------------------------------------------
    with tab_export:
        st.markdown("Download the full Risk & Opportunity Register as a formatted Excel workbook.")

        final_df = st.session_state.get("_ro_register_df", register_df)

        # Summary sheet
        status_counts = final_df.groupby(["Type","Priority"]).size().reset_index(name="Count")
        owner_counts  = final_df[final_df["Owner"] != ""].groupby("Owner").size().reset_index(name="Assigned Items")

        summary_data  = pd.DataFrame({
            "Metric": [
                "Total Items","Total Risks","Total Opportunities",
                "High Priority Risks","Medium Priority Risks","Low Priority Risks",
                "High Priority Opportunities","Medium Priority Opportunities","Low Priority Opportunities",
                "Open Items","In Progress","Closed",
            ],
            "Count": [
                len(final_df),
                int((final_df["Type"]=="Risk").sum()),
                int((final_df["Type"]=="Opportunity").sum()),
                int(((final_df["Type"]=="Risk")&(final_df["Priority"]=="High")).sum()),
                int(((final_df["Type"]=="Risk")&(final_df["Priority"]=="Medium")).sum()),
                int(((final_df["Type"]=="Risk")&(final_df["Priority"]=="Low")).sum()),
                int(((final_df["Type"]=="Opportunity")&(final_df["Priority"]=="High")).sum()),
                int(((final_df["Type"]=="Opportunity")&(final_df["Priority"]=="Medium")).sum()),
                int(((final_df["Type"]=="Opportunity")&(final_df["Priority"]=="Low")).sum()),
                int((final_df["Status"]=="Open").sum()),
                int((final_df["Status"]=="In Progress").sum()),
                int((final_df["Status"]=="Closed").sum()),
            ],
        })

        export_sheets = {
            "Summary":          summary_data,
            "Full Register":    final_df,
            "Risks":            final_df[final_df["Type"]=="Risk"],
            "Opportunities":    final_df[final_df["Type"]=="Opportunity"],
            "High Priority":    final_df[final_df["Priority"]=="High"],
            "Open Items":       final_df[final_df["Status"]=="Open"],
        }
        if not owner_counts.empty:
            export_sheets["By Owner"] = owner_counts

        xls_bytes = export_df_to_excel(export_sheets)

        dl_col, _ = st.columns([1,3])
        dl_col.download_button(
            label="📥  Download Risk & Opportunity Register",
            data=xls_bytes,
            file_name=f"risk_register_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Exports Summary, Full Register, Risks, Opportunities, High Priority and Open Items sheets.",
        )

        st.markdown(
            '<div style="background:#f8fafc;border:1px solid #E2E8F0;border-radius:8px;'
            'padding:14px 18px;margin-top:12px;">'
            '<div style="font-size:12px;font-weight:700;color:#0B1F33;margin-bottom:6px;">'
            'Workbook sheets</div>'
            '<div style="font-size:12px;color:#64748B;line-height:2;">'
            + "".join(f"<strong>{k}</strong> &nbsp;|&nbsp; " for k in export_sheets.keys()) +
            '</div></div>',
            unsafe_allow_html=True,
        )
        st.caption(
            "Note: This is an auto-generated draft register. "
            "Review all items with the project team before issuing formally."
        )



# -----------------------------------------------------------------------------
# PAGE: EXPORT REPORTS
# -----------------------------------------------------------------------------

def page_export_reports(data: dict, near_crit_days: float):
    st.title("📥 Export Reports")
    st.markdown("> Download all schedule data as formatted Excel reports.")

    tasks = data["tasks_df"]
    rels = data["relationships_df"]
    wbs = data["wbs_df"]
    resources = data["resources_df"]

    if tasks.empty:
        st.warning("No data loaded to export.")
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    critical = tasks[tasks["is_critical"]]
    neg_float = tasks[tasks["total_float_days"].apply(lambda f: f is not None and f < 0)]

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Single-Sheet Exports")

        # All activities
        avail = [c for c in ["task_code","task_name","wbs_path","eff_start","eff_finish",
                              "orig_dur_days","rem_dur_days","total_float_days","free_float_days",
                              "status","task_type","is_critical","cstr_type"] if c in tasks.columns]
        xls = export_df_to_excel({"All Activities": tasks[avail]})
        st.download_button("📄 All Activities", xls, "all_activities.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Critical path
        avail_c = [c for c in avail if c in critical.columns]
        xls2 = export_df_to_excel({"Critical Path": critical[avail_c]})
        st.download_button("🔴 Critical Path Activities", xls2, "critical_path.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Relationships
        if not rels.empty:
            xls3 = export_df_to_excel({"Relationships": rels})
            st.download_button("🔗 All Relationships", xls3, "relationships.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col2:
        st.subheader("Multi-Sheet Reports")

        # Full schedule pack
        sheets = {"All Activities": tasks[avail]}
        if not critical.empty:
            sheets["Critical Path"] = critical[avail_c]
        if not neg_float.empty:
            sheets["Negative Float"] = neg_float[[c for c in avail if c in neg_float.columns]]
        if not rels.empty:
            sheets["Relationships"] = rels
        if not wbs.empty:
            sheets["WBS"] = wbs
        if not resources.empty:
            sheets["Resources"] = resources

        xls_full = export_df_to_excel(sheets)
        st.download_button("📦 Full Schedule Data Pack", xls_full, "schedule_data_pack.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # WBS summary
        if "wbs_path" in tasks.columns:
            tasks["wbs_top"] = tasks["wbs_path"].apply(
                lambda x: str(x).split(" > ")[0] if pd.notna(x) and x else "Unknown"
            )
            wbs_summary = tasks.groupby("wbs_top").agg(
                total=("task_id","count"),
                critical=("is_critical","sum"),
                near_critical=("is_near_critical","sum"),
            ).reset_index()
            xls_wbs = export_df_to_excel({"WBS Summary": wbs_summary})
            st.download_button("🌲 WBS Summary", xls_wbs, "wbs_summary.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -----------------------------------------------------------------------------
# PAGE: HOME  (PlanTrace branded landing page)
# -----------------------------------------------------------------------------

def _page_home():
    """
    PlanTrace branded homepage.
    Shown when no XER is loaded, or when the user navigates to Home.
    """

    # ---- Hero section -------------------------------------------------------
    st.markdown(
        """
        <div style="padding: 48px 0 24px 0;">
            <div style="font-size: 13px; font-weight: 700; color: #F5A623;
                        letter-spacing: 2px; text-transform: uppercase;
                        margin-bottom: 10px;">
                Project Programme Intelligence
            </div>
            <div style="font-size: 52px; font-weight: 900; color: #0B1F33;
                        line-height: 1.1; letter-spacing: -1px;">
                PlanTrace
            </div>
            <div style="width: 56px; height: 4px; background: #F5A623;
                        border-radius: 2px; margin: 14px 0 18px 0;"></div>
            <div style="font-size: 20px; font-weight: 400; color: #334155;
                        margin-bottom: 10px;">
                Trace logic. Expose risk. Drive delivery.
            </div>
            <div style="font-size: 15px; color: #64748B; max-width: 680px;
                        line-height: 1.7; margin-bottom: 32px;">
                Project planning intelligence for delivery teams. Upload an XER programme,
                trace predecessors and successors, review critical paths, check programme
                health and understand labour demand &mdash; without opening P6.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Upload prompt -------------------------------------------------------
    st.markdown(
        """
        <div style="background:#0B1F33; border-radius:12px; padding:22px 28px;
                    display:flex; align-items:center; gap:20px; margin-bottom:36px;
                    max-width:560px;">
            <div style="font-size:28px;">📂</div>
            <div>
                <div style="color:#F5A623;font-weight:700;font-size:15px;
                            margin-bottom:4px;">Ready to start</div>
                <div style="color:#CBD5E1;font-size:13px;line-height:1.5;">
                    Upload your <strong style="color:#fff;">.xer file</strong>
                    using the panel on the left to begin analysis.
                    <br>Export from P6 via
                    <strong style="color:#F5A623;">File &rarr; Export &rarr; Primavera P6 XER</strong>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Feature cards -------------------------------------------------------
    st.markdown(
        '<div style="font-size:13px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1.5px;text-transform:uppercase;margin-bottom:16px;">'
        'What PlanTrace does</div>',
        unsafe_allow_html=True,
    )

    CARDS = [
        {
            "icon": "🔗",
            "title": "Logic Trace",
            "body": (
                "See what drives an activity and what it impacts. "
                "Trace full predecessor and successor chains across the network, "
                "with depth levels and relationship types shown at every step."
            ),
        },
        {
            "icon": "🚨",
            "title": "Critical Path",
            "body": (
                "Review the full critical path, near-critical work and negative float. "
                "Identify which activity or milestone is at risk and understand "
                "exactly what is driving it."
            ),
        },
        {
            "icon": "👷",
            "title": "Labour Demand",
            "body": (
                "View labour histograms by week, month, WBS and resource. "
                "Identify peak demand periods and understand resource loading "
                "across the programme."
            ),
        },
        {
            "icon": "🩺",
            "title": "Programme Health",
            "body": (
                "Find missing logic, open ends, constraints, excessive lag and "
                "planning risk before they cause problems. "
                "Eleven automated quality checks with export."
            ),
        },
    ]

    cols = st.columns(4, gap="medium")
    for col, card in zip(cols, CARDS):
        with col:
            st.markdown(
                f"""
                <div class="pt-card">
                    <div class="pt-card-icon">{card["icon"]}</div>
                    <div class="pt-card-accent"></div>
                    <div class="pt-card-title">{card["title"]}</div>
                    <div class="pt-card-body">{card["body"]}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ---- What's in the tool -------------------------------------------------
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:13px;font-weight:700;color:#94A3B8;'
        'letter-spacing:1.5px;text-transform:uppercase;margin-bottom:16px;">'
        'All pages</div>',
        unsafe_allow_html=True,
    )

    PAGE_LIST = [
        ("📊", "Project Summary",          "Activity counts, float distribution, WBS breakdown and schedule span."),
        ("🔍", "Activity Search",           "Search and filter activities. View full detail, dates, float and logic."),
        ("🔗", "Logic Trace",               "Trace predecessors and successors through the network with depth levels."),
        ("🚨", "Critical Path Analysis",    "Full critical path, near-critical and negative float by WBS."),
        ("🎯", "Critical Path to Activity", "Identify the driving chain into any selected activity or milestone."),
        ("👷", "Labour Histogram",          "Weekly and monthly labour demand by resource, WBS and package."),
        ("🩺", "Schedule Health Check",     "Eleven automated quality checks with counts, tables and export."),
        ("📝", "Planning Notes",            "Upload notes, link to activities, keyword search and highlighting."),
        ("📅", "Programme Comparison",      "Compare two XER revisions. See what moved, changed or became critical."),
        ("📥", "Export Reports",            "Download all data as formatted Excel workbooks."),
    ]

    left, right = st.columns(2, gap="large")
    for i, (icon, title, desc) in enumerate(PAGE_LIST):
        col = left if i % 2 == 0 else right
        with col:
            st.markdown(
                f"""
                <div style="display:flex;gap:14px;align-items:flex-start;
                            padding:14px 0;border-bottom:1px solid #E2E8F0;">
                    <div style="font-size:22px;min-width:30px;margin-top:2px;">{icon}</div>
                    <div>
                        <div style="font-weight:700;color:#0B1F33;
                                    font-size:14px;margin-bottom:3px;">{title}</div>
                        <div style="color:#64748B;font-size:13px;
                                    line-height:1.5;">{desc}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ---- Footer -------------------------------------------------------------
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style="border-top:1px solid #E2E8F0;padding-top:18px;
                    display:flex;justify-content:space-between;align-items:center;">
            <div style="font-size:13px;color:#94A3B8;">
                <strong style="color:#0B1F33;">PlanTrace</strong>
                &nbsp;&nbsp;|&nbsp;&nbsp;
                Built for Primavera P6 XER programmes
                &nbsp;&nbsp;|&nbsp;&nbsp;
                No P6 licence required
            </div>
            <div style="font-size:12px;color:#CBD5E1;">
                Upload a .xer file to begin
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# UI COMPONENTS  --  PlanTrace v2 Premium
# Brand: Plan=#2563EB  Trace=#2DD4BF  Risk=#DC2626  OK=#16A34A  Navy=#0A1628
# -----------------------------------------------------------------------------

# -- Brand colours ------------------------------------------------------------
PT_BLUE  = "#071827"   # Primary navy (was blue, now proper navy)
PT_TEAL  = "#F5A623"   # Accent amber (replaces teal)
PT_NAVY  = "#071827"   # Dark sidebar / header
PT_NAVY2 = "#0B2438"   # Secondary navy
PT_AMBER = "#F5A623"   # Amber accent
PT_RED   = "#DC2626"   # Risk / critical
PT_GREEN = "#16A34A"   # Healthy / complete
PT_BG    = "#F3F5F7"   # Page background
PT_CARD  = "#FFFFFF"   # Card background

# -- Navigation definition ----------------------------------------------------
_NAV = [
    ("overview",   "Overview",       ""),
    ("programme",  "Programme",      ""),
    ("logic",      "Logic",          ""),
    ("critical",   "Critical Path",  ""),
    ("labour",     "Labour",         ""),
    ("health",     "Health Check",   ""),
    ("comparison", "Comparison",     ""),
    ("pm_actions", "PM Actions",     ""),
    ("risk",       "Risk Register",  ""),
    ("reports",    "Reports",        ""),
    ("settings",   "Settings",       ""),
]

_NEEDS_PROG = {
    "overview", "programme", "logic", "critical",
    "labour", "health", "pm_actions", "risk", "reports",
}


# -----------------------------------------------------------------------------
# HELPER: embed logo as base64 img tag
# -----------------------------------------------------------------------------
import base64 as _b64
import os as _os

def _logo_tag(width: int = 100) -> str:
    """Return an <img> tag with the PlanTrace logo embedded as base64."""
    logo_path = "/mnt/user-data/uploads/1778872045571_image.png"
    if _os.path.exists(logo_path):
        with open(logo_path, "rb") as _f:
            _data = _b64.b64encode(_f.read()).decode()
        return (
            f'<img src="data:image/png;base64,{_data}" '
            f'width="{width}" style="display:block;" alt="PlanTrace">'
        )
    # Fallback: text-only logo
    return (
        f'<div style="font-size:18px;font-weight:900;line-height:1;letter-spacing:-0.5px;">'
        f'<span style="color:#FFFFFF;">Plan</span>'
        f'<span style="color:#F5A623;">Trace</span>'
        f'</div>'
    )


# -----------------------------------------------------------------------------
# HELPER: KPI card
# -----------------------------------------------------------------------------

def kpi(label: str, value, sub: str = "",
        colour: str = PT_NAVY, accent: str = PT_AMBER) -> str:
    """Premium KPI card HTML."""
    sub_html = (
        f'<div class="kpi-sub">{sub}</div>' if sub else ""
    )
    return (
        f'<div class="kpi-card" style="border-top:3px solid {accent};">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value" style="color:{colour};">{value}</div>'
        f'{sub_html}</div>'
    )


def kpi_row(cards: list):
    """Render a list of (label, value, sub, colour, accent) tuples as equal columns."""
    cols = st.columns(len(cards))
    for col, card in zip(cols, cards):
        args = card if isinstance(card, tuple) else (card,)
        col.markdown(kpi(*args), unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# HELPER: Status chips
# -----------------------------------------------------------------------------

def chip(label: str, style: str = "grey") -> str:
    """Inline HTML status chip. style: red|amber|green|blue|grey|teal"""
    icons = {
        "red": "●", "amber": "●", "green": "●",
        "blue": "●", "grey": "○", "teal": "●",
    }
    icon = icons.get(style, "●")
    return f'<span class="chip chip-{style}">{icon} {label}</span>'


def float_chip(tf) -> str:
    """Colour-coded float value chip."""
    if tf is None:
        return chip("-", "grey")
    try:
        f = float(tf)
    except Exception:
        return chip(str(tf), "grey")
    if f < 0:
        return chip(f"{f}d", "red")
    if f == 0:
        return chip("Critical", "red")
    if f <= 10:
        return chip(f"{f}d", "amber")
    return chip(f"{f}d", "green")


def status_chip(status_str: str) -> str:
    """Status badge for activity status."""
    s = str(status_str).strip()
    mapping = {
        "TK_NotStart": ("Not Started", "grey"),
        "Not Started": ("Not Started", "grey"),
        "TK_Active":   ("In Progress", "blue"),
        "In Progress": ("In Progress", "blue"),
        "TK_Complete": ("Complete",    "green"),
        "Complete":    ("Complete",    "green"),
    }
    label, style = mapping.get(s, (s or "-", "grey"))
    return chip(label, style)


# -----------------------------------------------------------------------------
# HELPER: Page header
# -----------------------------------------------------------------------------

def page_header(title: str, description: str = "", icon: str = ""):
    """Render the dark-navy per-page header with programme status strip."""
    prog_loaded = "programme" in st.session_state

    # Programme status strip
    if prog_loaded:
        prog   = st.session_state["programme"]
        pname  = prog.get("project_info", {}).get("name", "")
        ddate  = prog.get("project_info", {}).get("data_date")
        ntasks = len(prog.get("tasks_df", []))
        nrels  = len(prog.get("relationships_df", []))
        has_res = not prog.get("task_resources_df", pd.DataFrame()).empty
        has_notes = bool(st.session_state.get("_notes_text", ""))
        has_comp  = "_mi_prev" in st.session_state
        dd_str = format_date(ddate) if ddate else "N/A"

        chips_html = (
            chip("Loaded", "green") + "&nbsp;" +
            chip(f"{ntasks:,} activities", "grey") + "&nbsp;" +
            chip(f"{nrels:,} relationships", "grey") + "&nbsp;" +
            (chip("Resources", "green") if has_res else chip("No Resources", "grey")) + "&nbsp;" +
            (chip("Notes", "green") if has_notes else chip("No Notes", "grey")) + "&nbsp;" +
            (chip("Comparison", "amber") if has_comp else chip("No Comparison", "grey"))
        )
        prog_strip = (
            f'<div style="margin-top:10px;display:flex;align-items:center;'
            f'gap:8px;flex-wrap:wrap;">'
            f'<span style="font-size:11px;color:#8FA3B1;font-weight:700;">'
            f'{pname or "Programme"}</span>'
            f'<span style="color:#1E3A4A;">|</span>'
            f'<span style="font-size:11px;color:#4B6275;">Data date: '
            f'<strong style="color:#8FA3B1;">{dd_str}</strong></span>'
            f'<span style="color:#1E3A4A;">|</span>'
            f'{chips_html}'
            f'</div>'
        )
    else:
        prog_strip = (
            f'<div style="margin-top:8px;">'
            f'{chip("No programme loaded", "grey")}'
            f'</div>'
        )

    icon_html = f'<span style="margin-right:8px;font-size:18px;opacity:0.8;">{icon}</span>' if icon else ""
    desc_html = (
        f'<div style="font-size:12px;color:#4B6275;margin-top:4px;line-height:1.5;">'
        f'{description}</div>'
    ) if description else ""

    st.markdown(
        f"""
        <div class="pt-page-header">
            <div style="display:flex;justify-content:space-between;
                        align-items:flex-start;flex-wrap:wrap;gap:12px;">
                <div style="flex:1;min-width:0;">
                    <div style="font-size:11px;font-weight:700;color:#2D4557;
                                text-transform:uppercase;letter-spacing:2px;margin-bottom:6px;">
                        PlanTrace
                    </div>
                    <div style="font-size:22px;font-weight:800;color:#FFFFFF;
                                letter-spacing:-0.3px;line-height:1.1;">
                        {icon_html}{title}
                    </div>
                    {desc_html}
                    {prog_strip}
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# HELPER: Data quality card
# -----------------------------------------------------------------------------

def data_quality_card(data: dict):
    """Compact data quality summary card."""
    tasks  = data.get("tasks_df", pd.DataFrame())
    rels   = data.get("relationships_df", pd.DataFrame())
    res    = data.get("task_resources_df", pd.DataFrame())
    cals   = data.get("calendars_df", pd.DataFrame())
    notes  = bool(st.session_state.get("_notes_text", ""))
    comp   = "_mi_prev" in st.session_state

    items = [
        ("Activities",    len(tasks),         bool(len(tasks))),
        ("Relationships", len(rels),           bool(len(rels))),
        ("Resources",     len(res) if not res.empty else 0, not res.empty),
        ("Calendars",     len(cals) if not cals.empty else 0, not cals.empty),
        ("Notes",         "Yes" if notes else "No", notes),
        ("Comparison",    "Yes" if comp else "No",  comp),
    ]

    good_count = sum(1 for _, _, ok in items if ok)
    if good_count >= 5:
        overall_label, overall_style = "Good", "green"
    elif good_count >= 3:
        overall_label, overall_style = "Partial", "amber"
    else:
        overall_label, overall_style = "Limited", "grey"

    rows_html = ""
    for label, val, ok in items:
        dot_col = PT_GREEN if ok else "#9CA3AF"
        rows_html += (
            f'<div class="dq-row">'
            f'<span style="font-size:12px;color:#374151;">{label}</span>'
            f'<span style="display:flex;align-items:center;gap:5px;">'
            f'<span style="font-size:12px;font-weight:600;color:#0A1628;">{val}</span>'
            f'<span style="color:{dot_col};font-size:10px;">{"●" if ok else "○"}</span>'
            f'</span></div>'
        )

    st.markdown(
        f'<div style="background:#FFFFFF;border-radius:10px;padding:16px 18px;'
        f'border:1px solid #E5E7EB;box-shadow:0 1px 4px rgba(10,22,40,0.07);">'
        f'<div style="display:flex;justify-content:space-between;align-items:center;'
        f'margin-bottom:12px;">'
        f'<span style="font-size:11px;font-weight:700;color:#9CA3AF;text-transform:uppercase;'
        f'letter-spacing:1px;">Data Quality</span>'
        f'{chip(overall_label, overall_style)}'
        f'</div>'
        f'{rows_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# HELPER: PM Attention panel (auto-generates top issues)
# -----------------------------------------------------------------------------

def pm_attention_panel(data: dict, near_crit_days: float):
    """Auto-generate top 5 PM attention items from programme data."""
    tasks = data.get("tasks_df", pd.DataFrame())
    rels  = data.get("relationships_df", pd.DataFrame())
    notes = st.session_state.get("_notes_text", "")

    if tasks.empty:
        return

    tasks = get_critical_threshold(tasks, near_crit_days)
    items = []

    # 1. Negative float
    if "total_float_days" in tasks.columns:
        neg = tasks[tasks["total_float_days"].apply(lambda f: safe_float(f, 0) < 0)]
        if not neg.empty:
            items.append(("red", "Negative Float",
                f"{len(neg)} {'activity' if len(neg)==1 else 'activities'} cannot meet target dates. "
                "Immediate recovery action required."))

    # 2. Critical not started
    if "is_critical" in tasks.columns and "status" in tasks.columns:
        cns = tasks[tasks["is_critical"] & tasks["status"].apply(
            lambda s: str(s) in ("TK_NotStart","Not Started"))]
        if not cns.empty:
            items.append(("red", "Critical Activities Not Started",
                f"{len(cns)} critical {'activity' if len(cns)==1 else 'activities'} not yet started. "
                "Any delay will push out the project finish date."))

    # 3. Open-ended activities
    if not rels.empty and "task_id" in tasks.columns:
        with_pred = set(rels["succ_task_id"].dropna()) if "succ_task_id" in rels.columns else set()
        with_succ = set(rels["pred_task_id"].dropna()) if "pred_task_id" in rels.columns else set()
        open_starts = tasks[~tasks["task_id"].isin(with_pred)]
        open_ends   = tasks[~tasks["task_id"].isin(with_succ)]
        n_open = len(open_starts) + len(open_ends)
        if n_open > 0:
            items.append(("amber", "Open Logic",
                f"{n_open} activities have missing predecessor or successor logic. "
                "Float calculations may be unreliable."))

    # 4. Near-critical due soon
    from datetime import timedelta
    now = datetime.now()
    four_wks = now + timedelta(weeks=4)
    if "is_near_critical" in tasks.columns and "eff_finish" in tasks.columns:
        nc_soon = tasks[
            tasks["is_near_critical"] &
            tasks["eff_finish"].apply(
                lambda d: d is not None and hasattr(d,"date") and d <= four_wks)
        ]
        if not nc_soon.empty:
            items.append(("amber", "Near-Critical Due in 4 Weeks",
                f"{len(nc_soon)} near-critical {'activity is' if len(nc_soon)==1 else 'activities are'} "
                "finishing within 4 weeks with limited float buffer."))

    # 5. Planning notes risk words
    if notes and "task_code" in tasks.columns:
        import re as _re
        risk_acts = []
        for _, t in tasks.head(200).iterrows():
            code = str(t.get("task_code",""))
            if not code or code not in notes:
                continue
            idx = notes.find(code)
            snippet = notes[max(0,idx-200):idx+200]
            for word in _RISK_WORDS:
                if _re.search(r'\b'+_re.escape(word)+r'\b', snippet, _re.IGNORECASE):
                    risk_acts.append(code)
                    break
        if risk_acts:
            items.append(("red", "Risk Keywords in Planning Notes",
                f"{len(risk_acts)} {'activity' if len(risk_acts)==1 else 'activities'} "
                f"mentioned alongside risk keywords (delay, blocked, CE, EWN etc.) in planning notes."))

    # 6. Comparison: major slips
    if "_mi_prev" in st.session_state and "_mi_curr" in st.session_state:
        try:
            prev_t = st.session_state["_mi_prev"]["tasks_df"]
            curr_t = st.session_state["_mi_curr"]["tasks_df"]
            if not prev_t.empty and not curr_t.empty:
                merged = prev_t[["task_code","eff_finish"]].merge(
                    curr_t[["task_code","eff_finish"]], on="task_code",
                    suffixes=("_p","_c"), how="inner")
                slips = 0
                for _, r in merged.iterrows():
                    try:
                        d = int((pd.Timestamp(r["eff_finish_c"]) - pd.Timestamp(r["eff_finish_p"])).days)
                        if d > 14:
                            slips += 1
                    except Exception:
                        pass
                if slips:
                    items.append(("amber", "Major Date Slippage Detected",
                        f"{slips} activities have slipped more than 14 days compared to the previous revision."))
        except Exception:
            pass

    if not items:
        return

    # Show top 5
    items = items[:5]
    st.markdown(
        '<div class="section-label" style="margin-bottom:12px;">PM Attention Required</div>',
        unsafe_allow_html=True,
    )
    for style, title, body in items:
        cls = "attention-item" if style == "red" else (
              "attention-item attention-item-amber" if style == "amber"
              else "attention-item attention-item-blue")
        st.markdown(
            f'<div class="{cls}">'
            f'<div style="font-weight:700;color:#0A1628;font-size:13px;margin-bottom:3px;">'
            f'{title}</div>'
            f'<div style="font-size:12px;color:#374151;line-height:1.5;">{body}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
    st.markdown("<br>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# HELPER: Empty state card
# -----------------------------------------------------------------------------

def empty_state(icon: str, title: str, body: str, action: str = ""):
    """Premium empty-state card."""
    action_html = (
        f'<div style="margin-top:14px;display:inline-block;background:#071827;'
        f'color:white;border-radius:6px;padding:8px 18px;font-size:12px;font-weight:700;'
        f'letter-spacing:0.5px;border-left:3px solid #F5A623;">'
        f'{action}</div>'
    ) if action else ""
    icon_html = f'<div style="font-size:32px;margin-bottom:10px;">{icon}</div>' if icon else ""
    st.markdown(
        f'<div class="empty-state">'
        f'{icon_html}'
        f'<div style="font-size:16px;font-weight:700;color:#071827;margin-bottom:8px;">'
        f'{title}</div>'
        f'<div style="font-size:13px;color:#6B7280;max-width:400px;margin:0 auto;'
        f'line-height:1.6;">{body}</div>'
        f'{action_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# LANDING PAGE  (shown when no XER uploaded)
# -----------------------------------------------------------------------------

def _landing_page():
    """Professional control centre landing page shown when no programme is loaded."""

    # Control bar header
    st.markdown(
        f"""
        <div style="background:linear-gradient(180deg,#071827 0%,#0B2438 100%);
                    margin:-0px -28px 0 -28px;padding:18px 28px 16px 28px;
                    border-bottom:2px solid #F5A623;">
            <div style="display:flex;align-items:flex-end;justify-content:space-between;">
                <div>
                    <div style="font-size:11px;color:#2D4557;font-weight:700;
                                text-transform:uppercase;letter-spacing:2px;margin-bottom:4px;">
                        PlanTrace
                    </div>
                    <div style="font-size:26px;font-weight:800;color:#FFFFFF;
                                letter-spacing:-0.5px;line-height:1;">
                        Control Centre
                    </div>
                    <div style="font-size:13px;color:#4B6275;margin-top:5px;">
                        Planning intelligence for project delivery teams.
                    </div>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:10px;color:#2D4557;font-weight:700;
                                text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">
                        System Status
                    </div>
                    <div style="display:flex;align-items:center;gap:6px;">
                        <div style="width:6px;height:6px;border-radius:50%;background:#16A34A;"></div>
                        <span style="font-size:11px;color:#4B6275;font-weight:600;">Online</span>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Status strip
    st.markdown(
        f"""
        <div style="background:#0B2438;margin:-0px -28px 0 -28px;
                    padding:8px 28px;border-bottom:1px solid #102A43;
                    display:flex;align-items:center;gap:0;flex-wrap:wrap;">
            <div style="display:flex;align-items:center;gap:8px;padding-right:20px;
                        border-right:1px solid #102A43;margin-right:20px;">
                <div style="width:5px;height:5px;border-radius:50%;background:#374151;"></div>
                <span style="font-size:9px;color:#2D4557;font-weight:700;
                             text-transform:uppercase;letter-spacing:0.8px;">Programme</span>
                <span style="font-size:11px;color:#3A5060;font-weight:600;margin-left:4px;">Not Loaded</span>
            </div>
            <div style="display:flex;align-items:center;gap:8px;padding-right:20px;
                        border-right:1px solid #102A43;margin-right:20px;">
                <span style="font-size:9px;color:#2D4557;font-weight:700;
                             text-transform:uppercase;letter-spacing:0.8px;">Relationships</span>
                <span style="font-size:11px;color:#3A5060;font-weight:600;margin-left:4px;">Waiting</span>
            </div>
            <div style="display:flex;align-items:center;gap:8px;padding-right:20px;
                        border-right:1px solid #102A43;margin-right:20px;">
                <span style="font-size:9px;color:#2D4557;font-weight:700;
                             text-transform:uppercase;letter-spacing:0.8px;">Resources</span>
                <span style="font-size:11px;color:#3A5060;font-weight:600;margin-left:4px;">Waiting</span>
            </div>
            <div style="display:flex;align-items:center;gap:8px;padding-right:20px;
                        border-right:1px solid #102A43;margin-right:20px;">
                <span style="font-size:9px;color:#2D4557;font-weight:700;
                             text-transform:uppercase;letter-spacing:0.8px;">Notes</span>
                <span style="font-size:11px;color:#3A5060;font-weight:600;margin-left:4px;">Not Loaded</span>
            </div>
            <div style="display:flex;align-items:center;gap:8px;">
                <span style="font-size:9px;color:#2D4557;font-weight:700;
                             text-transform:uppercase;letter-spacing:0.8px;">Comparison</span>
                <span style="font-size:11px;color:#3A5060;font-weight:600;margin-left:4px;">Not Loaded</span>
            </div>
        </div>
        <div style="margin-bottom:24px;"></div>
        """,
        unsafe_allow_html=True,
    )

    # Upload CTA + description
    cta_col, desc_col = st.columns([1, 2], gap="large")
    with cta_col:
        st.markdown(
            f"""
            <div style="background:#071827;border-radius:8px;padding:22px;
                        border:1px solid #0B2438;border-left:3px solid #F5A623;">
                <div style="font-size:10px;color:#F5A623;font-weight:700;
                            text-transform:uppercase;letter-spacing:1.5px;margin-bottom:8px;">
                    Upload Programme
                </div>
                <div style="font-size:13px;color:#8FA3B1;line-height:1.6;margin-bottom:14px;">
                    Upload a Primavera P6 XER file using the <strong style="color:#CBD5E1;">
                    sidebar on the left</strong> to activate the control centre.
                </div>
                <div style="font-size:11px;color:#2D4557;line-height:1.7;">
                    Export from P6:<br>
                    <span style="color:#4B6275;">File &rarr; Export &rarr; Primavera P6 XER</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with desc_col:
        st.markdown(
            f"""
            <div style="padding:4px 0;">
                <div style="font-size:14px;color:#111827;font-weight:600;margin-bottom:8px;">
                    Planning intelligence for project delivery teams
                </div>
                <div style="font-size:13px;color:#6B7280;line-height:1.7;margin-bottom:16px;">
                    Upload a Primavera P6 XER programme to interrogate logic, critical paths,
                    labour demand, schedule quality and programme movement — without opening P6.
                </div>
                <div style="display:flex;gap:20px;flex-wrap:wrap;">
                    <div style="text-align:center;">
                        <div style="font-size:20px;font-weight:800;color:#071827;">11</div>
                        <div style="font-size:10px;color:#6B7280;text-transform:uppercase;
                                    letter-spacing:0.8px;font-weight:600;">Health Checks</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:20px;font-weight:800;color:#071827;">9</div>
                        <div style="font-size:10px;color:#6B7280;text-transform:uppercase;
                                    letter-spacing:0.8px;font-weight:600;">Analysis Modules</div>
                    </div>
                    <div style="text-align:center;">
                        <div style="font-size:20px;font-weight:800;color:#071827;">0</div>
                        <div style="font-size:10px;color:#6B7280;text-transform:uppercase;
                                    letter-spacing:0.8px;font-weight:600;">P6 Licence Required</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='margin-bottom:24px;'></div>", unsafe_allow_html=True)

    # Section label
    st.markdown(
        '<div class="section-label">Analysis Modules</div>',
        unsafe_allow_html=True,
    )

    # Module cards - row 1
    c1, c2, c3, c4 = st.columns(4, gap="small")
    modules_row1 = [
        (c1, "Logic Trace",
         "Trace predecessor and successor chains through the programme network.",
         ["Predecessors", "Successors", "Full chain", "Export"]),
        (c2, "Critical Path",
         "Identify critical work, near-critical activities and negative float conditions.",
         ["Critical work", "Near-critical", "Negative float", "Driving path"]),
        (c3, "Labour Demand",
         "Labour histograms by week, month, WBS and resource type.",
         ["Weekly histogram", "Monthly histogram", "WBS view", "Resource view"]),
        (c4, "Programme Health",
         "Eleven automated schedule quality checks against DCMA and industry standards.",
         ["Open logic", "Constraints", "Long durations", "Float issues"]),
    ]
    for col, title, desc, outputs in modules_row1:
        outputs_html = "".join(f"<li>{o}</li>" for o in outputs)
        col.markdown(
            f'<div class="module-card">'
            f'<div class="module-label">Module</div>'
            f'<div class="module-title">{title}</div>'
            f'<div class="module-desc">{desc}</div>'
            f'<div style="font-size:9px;color:#9CA3AF;text-transform:uppercase;'
            f'letter-spacing:1px;font-weight:700;margin-bottom:5px;">Outputs</div>'
            f'<ul class="module-outputs">{outputs_html}</ul>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='margin-bottom:12px;'></div>", unsafe_allow_html=True)

    # Row 2
    st.markdown('<div class="section-label">Additional Modules</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3, gap="small")
    modules_row2 = [
        (c1, "Programme Comparison",
         "Intelligence on movement between two XER revisions. Identify slippage, gains and changes.",
         ["Date movement", "Slip analysis", "Gained activities", "Risk & opportunity"]),
        (c2, "PM Actions",
         "Auto-generated prioritised action list based on programme analysis and risk indicators.",
         ["Prioritised actions", "Risk items", "Float warnings", "Export"]),
        (c3, "Reports",
         "Export all programme data, analysis results and health checks to formatted Excel workbooks.",
         ["Activity export", "Critical path", "Health check", "Full workbook"]),
    ]
    for col, title, desc, outputs in modules_row2:
        outputs_html = "".join(f"<li>{o}</li>" for o in outputs)
        col.markdown(
            f'<div class="module-card">'
            f'<div class="module-label">Module</div>'
            f'<div class="module-title">{title}</div>'
            f'<div class="module-desc">{desc}</div>'
            f'<div style="font-size:9px;color:#9CA3AF;text-transform:uppercase;'
            f'letter-spacing:1px;font-weight:700;margin-bottom:5px;">Outputs</div>'
            f'<ul class="module-outputs">{outputs_html}</ul>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Footer
    st.markdown(
        f'<div style="margin-top:32px;padding:16px 0;border-top:1px solid #E5E7EB;">'
        f'<span style="font-size:11px;color:#9CA3AF;">'
        f'<strong style="color:#071827;">PlanTrace</strong>'
        f' &nbsp;&mdash;&nbsp; Planning intelligence for project delivery'
        f' &nbsp;&mdash;&nbsp; No P6 licence required'
        f'</span></div>',
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# SIDEBAR
# -----------------------------------------------------------------------------

def _sidebar() -> tuple:
    """
    Render the PlanTrace branded sidebar.
    Returns (active_page_key: str, near_crit_days: int).
    """
    with st.sidebar:

        # -- Brand header ------------------------------------------------------
        logo_html = _logo_tag(width=80)
        st.markdown(
            f'<div style="padding:16px 14px 12px 14px;border-bottom:1px solid #0B2438;">'
            f'<div style="display:flex;align-items:center;gap:10px;">'
            f'{logo_html}'
            f'<div>'
            f'<div style="font-size:16px;font-weight:900;line-height:1;letter-spacing:-0.3px;">'
            f'<span style="color:#FFFFFF;">Plan</span>'
            f'<span style="color:#F5A623;">Trace</span>'
            f'</div>'
            f'<div style="font-size:9px;color:#2D4557;text-transform:uppercase;'
            f'letter-spacing:2px;margin-top:3px;font-weight:700;">Planning Intelligence</div>'
            f'</div></div></div>',
            unsafe_allow_html=True,
        )

        # -- XER Upload --------------------------------------------------------
        st.markdown(
            '<div style="padding:12px 14px 0 14px;">'
            '<div style="font-size:9px;color:#2D4557;font-weight:700;'
            'letter-spacing:1.5px;text-transform:uppercase;margin-bottom:5px;">'
            'Programme File</div>',
            unsafe_allow_html=True,
        )
        xer_file = st.file_uploader(
            "Upload XER", type=["xer"],
            label_visibility="collapsed",
            help="Export from P6: File > Export > Primavera P6 XER",
            key="sidebar_xer_upload",
        )
        st.markdown("</div>", unsafe_allow_html=True)

        # Parse on new upload
        if xer_file is not None:
            ck = f"xer_{xer_file.name}_{xer_file.size}"
            if st.session_state.get("_xer_cache_key") != ck:
                with st.spinner("Parsing programme..."):
                    try:
                        parsed = parse_xer(xer_file.read())
                        st.session_state["programme"]      = parsed
                        st.session_state["_xer_cache_key"] = ck
                        st.session_state["_xer_filename"]  = xer_file.name
                    except Exception as e:
                        st.error(f"Parse error: {e}")
                        st.session_state.pop("programme", None)

        # -- Programme status card ---------------------------------------------
        if "programme" in st.session_state:
            prog   = st.session_state["programme"]
            fname  = st.session_state.get("_xer_filename", "")
            ntasks = len(prog.get("tasks_df", []))
            nrels  = len(prog.get("relationships_df", []))
            ddate  = prog.get("project_info", {}).get("data_date")
            pname  = prog.get("project_info", {}).get("name", "")
            dd_s   = format_date(ddate) if ddate else "N/A"
            pname_line = (
                f'<div style="font-size:10px;color:#4B5563;margin-bottom:3px;'
                f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"'
                f' title="{pname}">{pname}</div>'
                if pname else ""
            )
            st.markdown(
                f'<div style="margin:8px 10px;background:#0B2438;'
                f'border:1px solid #102A43;border-radius:6px;padding:10px 12px;">'
                f'<div style="display:flex;align-items:center;gap:5px;margin-bottom:5px;">'
                f'<div style="width:5px;height:5px;border-radius:50%;'
                f'background:{PT_GREEN};flex-shrink:0;"></div>'
                f'<div style="font-size:9px;font-weight:700;color:{PT_GREEN};'
                f'letter-spacing:1px;text-transform:uppercase;">Loaded</div>'
                f'</div>'
                f'<div style="font-size:11px;font-weight:700;color:#CBD5E1;margin-bottom:2px;'
                f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"'
                f' title="{fname}">{fname}</div>'
                f'{pname_line}'
                f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-top:6px;">'
                f'<div style="background:#071827;border-radius:3px;padding:4px 8px;">'
                f'<div style="font-size:8px;color:#2D4557;text-transform:uppercase;'
                f'letter-spacing:0.8px;font-weight:700;">Activities</div>'
                f'<div style="font-size:14px;font-weight:800;color:#F5A623;">{ntasks:,}</div>'
                f'</div>'
                f'<div style="background:#071827;border-radius:3px;padding:4px 8px;">'
                f'<div style="font-size:8px;color:#2D4557;text-transform:uppercase;'
                f'letter-spacing:0.8px;font-weight:700;">Rels</div>'
                f'<div style="font-size:14px;font-weight:800;color:#F5A623;">{nrels:,}</div>'
                f'</div></div>'
                f'<div style="font-size:10px;color:#2D4557;margin-top:5px;">'
                f'Data date: <span style="color:#4B6275;">{dd_s}</span></div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div style="margin:8px 10px;background:#0B2438;'
                f'border:1px solid #0f2236;border-radius:6px;padding:10px 12px;">'
                f'<div style="display:flex;align-items:center;gap:5px;margin-bottom:4px;">'
                f'<div style="width:5px;height:5px;border-radius:50%;'
                f'background:#2D4557;flex-shrink:0;"></div>'
                f'<div style="font-size:9px;font-weight:700;color:#2D4557;'
                f'text-transform:uppercase;letter-spacing:1px;">No Programme Loaded</div>'
                f'</div>'
                f'<div style="font-size:11px;color:#1E3A4A;line-height:1.5;">'
                f'Upload a .xer file above.</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        # -- Navigation --------------------------------------------------------
        st.markdown(
            '<div style="margin:12px 10px 6px 10px;">'
            '<div style="font-size:9px;color:#1E3A4A;font-weight:700;'
            'letter-spacing:1.5px;text-transform:uppercase;margin-bottom:6px;'
            'border-top:1px solid #0B2438;padding-top:10px;">Navigation</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        prog_loaded = "programme" in st.session_state

        if "nav_page" not in st.session_state:
            st.session_state["nav_page"] = "overview"
        current = st.session_state["nav_page"]

        for key, label, icon in _NAV:
            is_active   = (current == key)
            is_disabled = (not prog_loaded) and (key in _NEEDS_PROG)

            if is_active:
                bg_style  = f"background:#0B2438!important;"
                txt_style = f"color:#F5A623!important;"
                border    = f"border-left:3px solid #F5A623!important;"
            elif is_disabled:
                bg_style  = "background:transparent!important;"
                txt_style = "color:#1A2E3A!important;"
                border    = "border-left:3px solid transparent!important;"
            else:
                bg_style  = "background:transparent!important;"
                txt_style = "color:#4B6275!important;"
                border    = "border-left:3px solid transparent!important;"

            opacity = "opacity:0.3;" if is_disabled else ""

            st.markdown(
                f'<div style="margin:1px 0;{bg_style}{border}{opacity}">',
                unsafe_allow_html=True,
            )
            if st.button(f"{label}", key=f"nav_{key}",
                         use_container_width=True, disabled=is_disabled):
                st.session_state["nav_page"] = key
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # -- Settings ----------------------------------------------------------
        st.markdown(
            '<div style="margin:10px 10px 4px 10px;">'
            '<div style="font-size:9px;color:#1E3A4A;font-weight:700;'
            'letter-spacing:1.5px;text-transform:uppercase;margin-bottom:5px;'
            'border-top:1px solid #0B2438;padding-top:10px;">Settings</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        near_crit_days = st.slider(
            "Near-Critical Float (days)",
            min_value=1, max_value=30, value=10, step=1,
        )

        # -- Footer ------------------------------------------------------------
        st.markdown(
            '<div style="padding:10px 14px;border-top:1px solid #0B2438;margin-top:12px;">'
            '<div style="font-size:9px;color:#1E3A4A;line-height:1.7;font-weight:600;">'
            'P6 Export: File &rarr; Export &rarr; P6 XER</div>'
            '<div style="font-size:9px;color:#152233;margin-top:2px;">'
            'v3.0 &nbsp;|&nbsp; Planning Intelligence</div>'
            '</div>',
            unsafe_allow_html=True,
        )

    return st.session_state["nav_page"], near_crit_days


# -----------------------------------------------------------------------------
# GUARD
# -----------------------------------------------------------------------------

def _guard() -> bool:
    """Show empty state and return True if no programme loaded."""
    if "programme" not in st.session_state:
        empty_state(
            "",
            "No Programme Loaded",
            "Upload a .xer file using the sidebar on the left to unlock this page. "
            "Export from Primavera P6 via File &rarr; Export &rarr; Primavera P6 XER.",
            "Upload in Sidebar",
        )
        return True
    return False


# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

def main():
    active, near_crit_days = _sidebar()
    prog_loaded = "programme" in st.session_state
    data = st.session_state.get("programme", {})

    # -- OVERVIEW ------------------------------------------------------------
    if active == "overview":
        if not prog_loaded:
            _landing_page()
        else:
            page_header(
                "Overview",
                "Programme summary, KPI indicators and top issues requiring your attention.",
                "",
            )
            # Data quality card on the right, KPI row on the left
            tasks = data.get("tasks_df", pd.DataFrame())
            if not tasks.empty:
                tasks_kpi = get_critical_threshold(tasks, near_crit_days)
                n_crit  = int(tasks_kpi["is_critical"].sum()) if "is_critical" in tasks_kpi.columns else 0
                n_nc    = int(tasks_kpi["is_near_critical"].sum()) if "is_near_critical" in tasks_kpi.columns else 0
                n_neg   = int(tasks_kpi["total_float_days"].apply(lambda f: safe_float(f,0) < 0).sum()) if "total_float_days" in tasks_kpi.columns else 0
                n_rels  = len(data.get("relationships_df", pd.DataFrame()))

                kpi_row([
                    ("Total Activities",    len(tasks),       "",                   PT_NAVY,  PT_AMBER),
                    ("Critical",            n_crit,           "float ≤ 0",          PT_RED,   PT_RED),
                    ("Near-Critical",       n_nc,             f"float ≤ {near_crit_days}d", "#92400E", PT_AMBER),
                    ("Negative Float",      n_neg,            "beyond target date", PT_RED,   PT_RED) if n_neg > 0
                        else ("Negative Float", n_neg,        "all on track",       PT_GREEN, PT_GREEN),
                    ("Relationships",       n_rels,           "",                   PT_NAVY,  PT_AMBER),
                ])
                st.markdown("<br>", unsafe_allow_html=True)

            # Attention panel + data quality side by side
            left_col, right_col = st.columns([3, 1], gap="large")
            with left_col:
                pm_attention_panel(data, near_crit_days)
            with right_col:
                data_quality_card(data)

            # Existing overview page content below
            page_project_summary(data, near_crit_days)

    # -- PROGRAMME ------------------------------------------------------------
    elif active == "programme":
        page_header("Programme", "Search activities, plan lookahead, track milestones and review planning notes.", "")
        if _guard():
            return
        p_tabs = st.tabs(["Activity Search", "Lookahead", "Milestones", "Planning Notes"])
        with p_tabs[0]: page_activity_search(data, near_crit_days)
        with p_tabs[1]: page_lookahead(data, near_crit_days)
        with p_tabs[2]: page_milestone_tracker(data, near_crit_days)
        with p_tabs[3]: page_planning_notes(data)

    # -- LOGIC ----------------------------------------------------------------
    elif active == "logic":
        page_header("Logic", "Trace predecessor and successor chains through the programme network.", "")
        if _guard():
            return
        l_tabs = st.tabs(["Logic Trace", "Path to Selected Activity"])
        with l_tabs[0]: page_logic_trace(data, near_crit_days)
        with l_tabs[1]: page_critical_path_to_activity(data, near_crit_days)

    # -- CRITICAL PATH --------------------------------------------------------
    elif active == "critical":
        page_header("Critical Path", "Full critical path, near-critical activities and negative float analysis.", "")
        if _guard():
            return
        page_critical_path(data, near_crit_days)

    # -- LABOUR ---------------------------------------------------------------
    elif active == "labour":
        page_header("Labour", "Labour demand histograms by week, month, resource and WBS.", "")
        if _guard():
            return
        task_res = data.get("task_resources_df", pd.DataFrame())
        if task_res.empty:
            empty_state(
                "",
                "No Resource Data Found",
                "This XER file does not contain resource loading. "
                "Resources must be assigned in P6 and included in the export. "
                "You can also upload a separate CSV file with columns: "
                "task_code, rsrc_name, target_qty, target_start, target_finish.",
            )
        page_labour_histogram(data)

    # -- HEALTH CHECK ---------------------------------------------------------
    elif active == "health":
        page_header("Health Check", "Automated schedule quality checks covering logic, float, constraints and durations.", "")
        if _guard():
            return
        page_health_check(data, near_crit_days)

    # -- COMPARISON -----------------------------------------------------------
    elif active == "comparison":
        page_header("Comparison", "Programme movement intelligence between two XER revisions.", "")
        c_tabs = st.tabs(["Programme Movement", "Risk & Opportunity"])
        with c_tabs[0]:
            page_programme_comparison()
        with c_tabs[1]:
            if _guard():
                pass
            else:
                page_risk_register(data, near_crit_days)

    # -- PM ACTIONS -----------------------------------------------------------
    elif active == "pm_actions":
        page_header("PM Actions", "Auto-generated prioritised action list based on programme analysis.", "")
        if _guard():
            return
        page_pm_actions(data, near_crit_days)

    # -- RISK REGISTER --------------------------------------------------------
    elif active == "risk":
        page_header("Risk Register", "Auto-generated risk and opportunity register from programme data.", "")
        if _guard():
            return
        page_risk_register(data, near_crit_days)

    # -- REPORTS --------------------------------------------------------------
    elif active == "reports":
        page_header("Reports", "Export all programme data and analysis to formatted Excel workbooks.", "")
        if _guard():
            return
        page_export_reports(data, near_crit_days)

    # -- SETTINGS -------------------------------------------------------------
    elif active == "settings":
        page_header("Settings", "App configuration and loaded programme details.", "")
        st.markdown("### Configuration")
        st.info(
            f"Near-critical float threshold: **{near_crit_days} days**. "
            "Adjust using the slider in the sidebar."
        )
        if prog_loaded:
            st.markdown("### Loaded Programme")
            proj = data.get("project_info", {})
            st.json({
                "name":          proj.get("name", ""),
                "data_date":     format_date(proj.get("data_date")),
                "activities":    len(data.get("tasks_df", [])),
                "relationships": len(data.get("relationships_df", [])),
                "parse_method":  data.get("parse_method", ""),
            })
            st.markdown("### Data Quality")
            data_quality_card(data)


if __name__ == "__main__":
    main()