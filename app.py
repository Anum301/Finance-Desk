# ===== ANUM'S FINANCEDESK - V49 - SAME 2-COL STYLE, JUST SMALLER =====
MY_NAME = "Anum"
MY_OFFICE = "HSS - Finance Operations, Lahore"
DATA_FILE = "FinanceDesk_Data.xlsx"
# ================================================================

import streamlit as st
import pandas as pd
from datetime import date, datetime
import io, os
st.set_page_config(page_title=f"FinanceDesk - {MY_NAME}", page_icon="🟩", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
.stApp, .block-container, .h-card, .h-label, .h-num, .h-title, .h-sub, .h-text, .h-small, p, h1, h2, h3, label, input, select, textarea, button { font-family: 'Times New Roman', Times, serif !important; }
.stApp { background-color: #F3F4F1 !important; }
section[data-testid="stSidebar"] { background-color: #101E14 !important; }
section[data-testid="stSidebar"] * { color: #E8EDE8 !important; font-size: 14px !important; }
section[data-testid="stSidebar"] h3 { color: #fff !important; font-size: 20px !important; font-weight: 800 !important; }
section[data-testid="stSidebar"] [data-testid="stCaption"] { color: #8A9A8E !important; font-size: 11px !important; letter-spacing: 1px !important; font-weight: 700 !important; }
.block-container { padding-top: 70px !important; padding-left: 32px !important; padding-right: 32px !important; }
header[data-testid="stHeader"] { background: transparent !important; height: 45px !important; }
div[data-testid="stToolbar"] { top: 8px !important; right: 16px !important; }
div[data-testid="stDecoration"] { display: none !important; }
/* V49 - SAME STYLE, JUST SMALLER - 2 COLUMNS KEPT! */
div[data-testid="stDialog"] { display: flex !important; align-items: center !important; justify-content: center !important; }
div[data-testid="stDialog"] > div { width: 680px !important; max-width: 92vw !important; min-height: 88vh !important; max-height: 92vh !important; margin: auto !important; padding: 12px 18px !important; overflow-y: auto !important; overflow-x: hidden !important; border-radius: 14px !important; }
div[data-testid="stDialog"] > div > div { gap: 4px !important; }
div[data-testid="stDialog"] label p { font-family: 'Times New Roman', Times, serif !important; font-size: 12px !important; font-weight: 700 !important; color: #101E14 !important; margin-bottom: 1px !important; }
div[data-testid="stDialog"] input, div[data-testid="stDialog"] select { font-family: 'Times New Roman', Times, serif !important; font-size: 13px !important; height: 32px !important; min-height: 32px !important; padding: 4px 10px !important; border-radius: 8px !important; }
div[data-testid="stDialog"] textarea { font-family: 'Times New Roman', Times, serif !important; font-size: 13px !important; padding: 6px 10px !important; border-radius: 8px !important; }
div[data-testid="stDialog"] div[data-testid="stVerticalBlock"] { gap: 4px !important; }
div[data-testid="stDialog"] div[data-testid="stHorizontalBlock"] { gap: 8px !important; }
div[data-testid="stDialog"] button { height: 34px !important; font-size: 13px !important; }
div[data-testid="stFileUploader"] section > div > span, div[data-testid="stFileUploader"] section button span, div[data-testid="stFileUploader"] section svg { display: none !important; }
div[data-testid="stFileUploader"] section { padding: 8px 10px !important; min-height: 0 !important; }
div[data-testid="stVerticalBlockBorderWrapper"]:has(.reg-marker), div[data-testid="stContainer"]:has(.reg-marker), section:has(.reg-marker) { background: #D9EAD3 !important; }
div[data-testid="stVerticalBlockBorderWrapper"]:has(.reg-marker) > div, div[data-testid="stVerticalBlockBorderWrapper"]:has(.reg-marker) div[data-testid="stVerticalBlock"], div[data-testid="stVerticalBlockBorderWrapper"]:has(.reg-marker) div[data-testid="stHorizontalBlock"] { background: #D9EAD3 !important; }
div[data-testid="stVerticalBlockBorderWrapper"] { background: #D9EAD3 !important; border: 1px solid #B8D0B0 !important; border-radius: 14px !important; padding: 18px !important; }
div[data-testid="stVerticalBlockBorderWrapper"] > div { background: #D9EAD3 !important; }
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"] { background: #D9EAD3 !important; }
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"] { background: #D9EAD3 !important; }
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stTextInput"] > div,
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stSelectbox"] > div,
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stNumberInput"] > div,
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stDateInput"] > div { background: #fff !important; }
div[data-testid="stVerticalBlockBorderWrapper"] input,
div[data-testid="stVerticalBlockBorderWrapper"] select { background: #fff !important; }
.h-card { background: #fff; border: 1px solid #E9ECE8; border-radius: 14px; padding: 16px 18px; transition: all 0.2s ease !important; }
.h-label { font-size: 11px !important; letter-spacing: 1px; color: #8A9A8E !important; font-weight: 700; text-transform: uppercase; }
.h-num { font-size: 25px !important; font-weight: 800 !important; color: #101E14 !important; margin: 7px 0 3px 0; }
.h-title { font-size: 26px !important; font-weight: 800 !important; color: #101E14 !important; }
.h-sub { font-size: 16px !important; font-weight: 700 !important; color: #101E14 !important; margin: 4px 0 3px 0; }
.h-text { font-size: 13.5px !important; color: #5C6B60 !important; }
.h-small { font-size: 12.5px !important; color: #8A9A8E !important; }
button[kind="primary"] { background: #1A2E22 !important; color: white !important; border: none !important; font-size: 13.5px !important; height: 38px !important; border-radius: 9px !important; }
button[kind="secondary"] { font-size: 13.5px !important; height: 38px !important; border-radius: 9px !important; background: white !important; color: #1A2E22 !important; }
.reg-head { background: #E9EFE7 !important; border-top: 1px solid #D0DAD0 !important; border-bottom: 1px solid #D0DAD0 !important; padding: 10px 16px !important; display: flex; font-size: 11.5px !important; color: #1A2E22 !important; font-weight: 800 !important; }
div[data-testid="stSidebarCollapsedControl"], div[data-testid="collapsedControl"],
div[data-testid="stSidebarCollapsedControl"] *, div[data-testid="collapsedControl"] *,
button[data-testid="stSidebarCollapseButton"], button[data-testid="stSidebarCollapseButton"] *,
section[data-testid="stSidebar"] button[aria-label="Close sidebar"],
section[data-testid="stSidebar"] button[aria-label="Close sidebar"] * { font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important; font-size: 24px !important; color: #101E14 !important; display: flex !important; visibility: visible !important; opacity: 1 !important; }
section[data-testid="stSidebar"] button[aria-label="Close sidebar"] *, section[data-testid="stSidebar"] button[aria-label="Close sidebar"] { color: #fff !important; }
div[data-testid="stSidebarCollapsedControl"] { width: 44px !important; height: 44px !important; background: transparent !important; }
div[data-testid="stSidebarCollapsedControl"] button { width: 44px !important; height: 44px !important; background: transparent !important; border: none !important; }
button { transition: all 0.2s ease !important; }
div[data-testid="stButton"] { transition: all 0.2s ease !important; }
div[data-testid="stButton"] button { transition: all 0.2s ease !important; }
div[data-testid="stButton"]:hover { transform: translateY(-3px) !important; }
div[data-testid="stButton"] button:hover { transform: translateY(-3px) !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; }
div[data-testid="stButton"] button[kind="primary"]:hover { transform: translateY(-3px) !important; background: #1A2E22 !important; color: white !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.30) !important; }
div[data-testid="stDownloadButton"] { transition: all 0.2s ease !important; }
div[data-testid="stDownloadButton"] button { transition: all 0.2s ease !important; background: #fff !important; color: #1A2E22 !important; border: 1px solid #D0DAD0 !important; font-weight: 700 !important; }
div[data-testid="stDownloadButton"]:hover { transform: translateY(-3px) !important; }
div[data-testid="stDownloadButton"]:hover button { transform: translateY(-3px) !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #F7FAF5 !important; }
div[data-testid="stFormSubmitButton"] button:hover { transform: translateY(-3px) !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; }
.h-card:hover { transform: translateY(-3px) !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #F7FAF5 !important; }
.no-move:hover { transform: none !important; border: 1px solid #E9ECE8 !important; box-shadow: none !important; background: #fff !important; }
.inner-box { transition: all 0.2s ease !important; }
.inner-box:hover { transform: translateY(-3px) !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #FFFFFF !important; }
section[data-testid="stSidebar"] label:hover { background: #1A2E22 !important; border-radius: 8px !important; padding: 6px 10px !important; }
section[data-testid="stSidebar"] label:hover p { color: white !important; font-weight: 700 !important; }
.st-open { background: #FFF3CD; color: #856404; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 12px; }
.st-overdue { background: #F8D7DA; color: #721C24; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 12px; }
.st-pending { background: #D1ECF1; color: #0C5460; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 12px; }
.st-payment { background: #D4EDDA; color: #155724; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 12px; }
.st-completed { background: #E2E3E5; color: #383D41; padding: 3px 10px; border-radius: 20px; font-weight: 700; font-size: 12px; }
.reg-row { transition: all 0.2s ease !important; border-radius: 10px !important; padding: 6px 10px !important; margin: 2px 0 !important; }
.reg-row:hover { transform: translateY(-3px) !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #F7FAF5 !important; }
.last-box { transition: all 0.2s ease !important; }
.last-box:hover { transform: translateY(-3px) !important; border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #F7FAF5 !important; }
div[data-testid="stTextInput"], div[data-testid="stSelectbox"], div[data-testid="stNumberInput"], div[data-testid="stDateInput"] { transition: all 0.2s ease !important; }
div[data-testid="stTextInput"] > div, div[data-testid="stSelectbox"] > div, div[data-testid="stNumberInput"] > div, div[data-testid="stDateInput"] > div { transition: all 0.2s ease !important; border-radius: 9px !important; background: #fff !important; }
div[data-testid="stTextInput"]:hover, div[data-testid="stSelectbox"]:hover, div[data-testid="stNumberInput"]:hover, div[data-testid="stDateInput"]:hover { transform: translateY(-3px) !important; }
div[data-testid="stTextInput"]:hover > div, div[data-testid="stSelectbox"]:hover > div, div[data-testid="stNumberInput"]:hover > div, div[data-testid="stDateInput"]:hover > div { border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #F7FAF5 !important; }
div[data-testid="stTextInput"]:focus-within > div, div[data-testid="stSelectbox"]:focus-within > div, div[data-testid="stNumberInput"]:focus-within > div, div[data-testid="stDateInput"]:focus-within > div { border: 2px solid #1A2E22 !important; box-shadow: 0 8px 20px rgba(26,46,34,0.18) !important; background: #F7FAF5 !important; }
input:hover, select:hover { background: #F7FAF5 !important; }
</style>
""", unsafe_allow_html=True)
def load_tasks():
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_excel(DATA_FILE)
            return df.to_dict("records")
        except: return []
    return []
def save_tasks(tasks_list):
    try: pd.DataFrame(tasks_list).to_excel(DATA_FILE, index=False)
    except: pass
def make_excel_bytes(tasks_list):
    buf = io.BytesIO()
    df = pd.DataFrame(tasks_list) if tasks_list else pd.DataFrame(columns=["Received","Task","Requestor","Category","Reference","Budget Head","Amount","Status","Pending With","Follow-up","End Date","Priority","Proof","Assigned To","Due Date","Reason","Any Date","Invoice"])
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Finance Register", index=False)
        ws = w.sheets["Finance Register"]
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.oddHeader.center.text = f"FinanceDesk - {MY_NAME} - {MY_OFFICE}"; ws.oddFooter.center.text = f"Prepared by {MY_NAME} - {date.today().strftime('%d %b %Y')} - Page &P of &N"
    buf.seek(0); return buf.getvalue()
def make_csv_bytes(tasks_list):
    if tasks_list: return pd.DataFrame(tasks_list).to_csv(index=False).encode('utf-8')
    return pd.DataFrame(columns=["Received","Task","Requestor","Category","Reference","Budget Head","Amount","Status","Pending With","Follow-up","End Date","Priority","Proof","Assigned To","Due Date","Reason","Any Date","Invoice"]).to_csv(index=False).encode('utf-8')
def clear_filters_for(title_cat):
    for k in list(st.session_state.keys()):
        if k.startswith(f"search_{title_cat}") or k.startswith(f"status_{title_cat}") or k.startswith(f"type_{title_cat}") or k.startswith(f"pri_{title_cat}") or k.startswith(f"person_{title_cat}") or k.startswith(f"budget_{title_cat}"):
            del st.session_state[k]
if "tasks" not in st.session_state: st.session_state.tasks = load_tasks()
if "show_form" not in st.session_state: st.session_state.show_form = False
def GAP(): st.markdown('<div style="height:18px;"></div>', unsafe_allow_html=True)
def SMALL_GAP(): st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)
def status_pill(s):
    m = {"Open":"st-open","Overdue":"st-overdue","Pending External":"st-pending","Payment Stage":"st-payment","Completed":"st-completed","Received":"st-open"}
    return f'<span class="{m.get(s,"st-open")}">{s}</span>'
# ===== V49 - SAME 2-COLUMN STYLE AS PHOTO - JUST SMALLER! =====
@st.dialog("Create task record", width="medium")
def create_task_dialog():
    st.markdown('<div style="font-family:Times New Roman; font-size:10px; letter-spacing:2px; color:#8A9A8E; font-weight:700;">FINANCE TASK DETAILS</div><div style="font-family:Times New Roman; font-size:20px; font-weight:800; color:#101E14; margin:0;">Create task record</div><div style="font-family:Times New Roman; font-size:12px; color:#5C6B60; margin-bottom:6px;">Capture the request, current owner and next action.</div>', unsafe_allow_html=True)
    task_name = st.text_input("Task name / description", placeholder="e.g. Dr. Amna Waqar — reimbursement")
    d1, d2 = st.columns(2)
    with d1: received = st.date_input("Date task received", value=date.today())
    with d2: requestor = st.text_input("Requestor / faculty", placeholder="Name of faculty / staff")
    e1, e2 = st.columns(2)
    with e1: category = st.selectbox("Task type", ["Reimbursement", "PRs & procurement", "Procurement", "Travel claims", "Tax reviews", "Advances", "Grants", "Local", "RA / TA"])
    with e2: ref = st.text_input("PR / memo / claim reference", placeholder="e.g. 3400078219")
    f1, f2 = st.columns(2)
    with f1: amount = st.number_input("Amount (PKR)", min_value=0, step=1000, value=0)
    with f2: budget = st.text_input("Budget head", placeholder="e.g. FIF-1117, FWC-157")
    g1, g2 = st.columns(2)
    with g1: status = st.selectbox("Current status", ["Received", "Open", "Overdue", "Pending External", "Payment Stage", "Completed"])
    with g2: priority = st.selectbox("Priority", ["Normal", "Low", "Medium", "High", "Urgent"])
    h1, h2 = st.columns(2)
    with h1: pending_with = st.text_input("Pending with Person", placeholder="e.g. Miss Madiha / Finance")
    with h2: any_date = st.date_input("Any date", value=date.today())
    i1, i2 = st.columns(2)
    with i1: due_date = st.date_input("Due date", value=date.today())
    with i2: follow = st.date_input("Next follow-up date", value=date.today())
    end = st.date_input("End / completion date", value=date.today())
    reason = st.text_area("Reason for delay", placeholder="If delayed, why? e.g. Awaiting approval, missing bill, etc.", height=45)
    invoice_file = st.file_uploader("Upload Invoice (optional)", type=["pdf","jpg","png","xlsx"])
    invoice_name = invoice_file.name if invoice_file else ""
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("Cancel", use_container_width=True):
            st.session_state.show_form = False; st.rerun()
    with c2:
        if st.button("Save Task", use_container_width=True, type="primary"):
            if not task_name:
                st.error("Please enter Task name / description")
            else:
                st.session_state.tasks.append({"Received": str(received), "Task": task_name, "Requestor": requestor, "Category": category, "Reference": ref, "Budget Head": budget, "Amount": amount, "Status": status, "Pending With": pending_with, "Follow-up": str(follow), "End Date": str(end), "Priority": priority, "Proof": invoice_name, "Assigned To": "", "Due Date": str(due_date), "Reason": reason, "Any Date": str(any_date), "Invoice": invoice_name})
                save_tasks(st.session_state.tasks); st.session_state.show_form = False; st.success(f"Saved, {MY_NAME}!"); st.rerun()
if st.session_state.show_form:
    create_task_dialog()
_hour = datetime.now().hour
if _hour < 12: _greet = "Good Morning"
elif _hour < 17: _greet = "Good Afternoon"
else: _greet = "Good Evening"
st.markdown(f'''<div style="display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; width:100%; box-sizing:border-box; overflow:visible; background:#fff; border:1px solid #E9ECE8; border-radius:12px; padding:14px 18px; margin-top:10px; margin-bottom:16px;">
<div style="flex:1; min-width:220px;"><div style="font-size:11px; letter-spacing:1px; color:#8A9A8E; font-weight:700;">HSS / FINANCE OPERATIONS</div><div style="font-size:18px; font-weight:800; white-space:normal;">{_greet}, {MY_NAME} 👋</div></div>
<div style="flex:1; min-width:220px; text-align:right; white-space:normal; overflow:visible;"><div style="font-size:13px; font-weight:700; white-space:normal; line-height:1.4;">{MY_OFFICE}</div><div style="font-size:12px; color:#8A9A8E; white-space:normal; line-height:1.4;">Today: {date.today().strftime("%d %b %Y")} · Prepared by {MY_NAME}</div></div></div>''', unsafe_allow_html=True)
with st.sidebar:
    st.markdown(f"### 🟩 FinanceDesk")
    st.caption(f"{MY_NAME.upper()} · WORKSPACE")
    menu = st.radio("Go to", ["🏠 Overview", "📋 All tasks", "➕ New task", "✏️ Edit / Delete", "📊 Dashboard", "🖨️ Print View", "📁 Audit / Export"])
    st.caption("FINANCE OPERATIONS")
    cat = st.radio("Category", ["All Categories", "PRs & procurement", "Procurement", "Reimbursements", "Travel claims", "Tax reviews", "Advances", "Grants", "Local", "RA / TA"], label_visibility="collapsed")
    st.divider()
    st.markdown(f"✅ **Keep every follow-up visible, {MY_NAME}.**")
    st.caption("Track the next action, the person holding it, and the date it needs attention.")
    st.caption(f"{MY_OFFICE}")
tasks = st.session_state.tasks
is_category_page = (cat != "All Categories")
tasks_view = [t for t in tasks if t["Category"] == cat] if is_category_page else tasks
open_tasks = len([t for t in tasks_view if t["Status"] != "Completed"])
overdue = len([t for t in tasks_view if t["Status"] == "Overdue"])
pending_ext = len([t for t in tasks_view if t["Status"] == "Pending External"])
payments = sum([t["Amount"] for t in tasks_view if t["Status"] == "Payment Stage"])
total_val = sum([t["Amount"] for t in tasks_view])
completed_val = sum([t["Amount"] for t in tasks_view if t["Status"] == "Completed"])
today_str = str(date.today())
due_today = [t for t in tasks_view if t.get("Follow-up") == today_str and t["Status"] != "Completed"]
_top1, _top2, _top3 = st.columns([4, 1.2, 1])
with _top2:
    st.download_button("⬇ Export CSV", make_csv_bytes(tasks_view), f"FinanceDesk_{MY_NAME}_{date.today()}.csv", "text/csv", use_container_width=True, key="top_export")
with _top3:
    if st.button("+ Add task", use_container_width=True, type="primary", key="top_add"):
        st.session_state.show_form = True; st.rerun()
if due_today:
    st.markdown(f'''<div style="background:#F8D7DA; border:1px solid #F5C6CB; border-radius:12px; padding:12px 16px; margin-bottom:14px;"><b>🔴 {MY_NAME}, you have {len(due_today)} follow-up(s) DUE TODAY:</b> {", ".join([t["Task"] for t in due_today])}</div>''', unsafe_allow_html=True)
else:
    st.markdown(f'''<div style="background:#D4EDDA; border:1px solid #C3E6CB; border-radius:12px; padding:12px 16px; margin-bottom:14px;">🟢 <b>All clear, {MY_NAME}!</b> No follow-ups due today.</div>''', unsafe_allow_html=True)
if is_category_page:
    st.markdown(f'<div class="h-label">HSS / FINANCE OPERATIONS / {cat.upper()} · {MY_NAME}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="h-title">{cat} page.</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="h-text">Showing only {cat} requests, approvals, follow-ups and payments.</div>', unsafe_allow_html=True)
else:
    st.markdown(f'<div class="h-label">HSS / FINANCE OPERATIONS · {MY_NAME}</div>', unsafe_allow_html=True)
    st.markdown('<div class="h-title">Task control centre.</div>', unsafe_allow_html=True)
    st.markdown('<div class="h-text">A clear view of requests, approvals, follow-ups and payments.</div>', unsafe_allow_html=True)
def finance_box(cat_name, total_v, comp_v, pay_v):
    st.markdown(f'''<div class="h-card no-move"><div class="h-label">Financial Snapshot - {cat_name} · {MY_NAME}</div><div class="h-sub">Finance at a glance</div><div class="h-small">Amounts are calculated from task records with an entered amount.</div><div style="display:flex; gap:12px; margin-top:14px;"><div class="inner-box" style="flex:1; background:#F7FAF5; border:1px solid #E9ECE8; border-radius:10px; padding:14px;"><div class="h-label">Total Tracked Value</div><div class="h-num">PKR {total_v:,}</div><div class="h-small">All recorded task amounts</div></div><div class="inner-box" style="flex:1; background:#F7FAF5; border:1px solid #E9ECE8; border-radius:10px; padding:14px;"><div class="h-label">Completed Value</div><div class="h-num">PKR {comp_v:,}</div><div class="h-small">Records marked completed</div></div><div class="inner-box" style="flex:1; background:#EDF2EB; border:1px solid #E9ECE8; border-radius:10px; padding:14px;"><div class="h-label">Payment Pipeline</div><div class="h-num">PKR {pay_v:,}</div><div class="h-small">Currently at payment stage</div></div></div><div class="h-small" style="margin-top:12px;">Add task amounts to see the financial breakdown here.</div></div>''', unsafe_allow_html=True)
def register_box(title_cat, show_export=True):
    GAP()
    with st.container(border=True):
        st.markdown('<div class="reg-marker" style="display:none;"></div>', unsafe_allow_html=True)
        _h1, _h2, _h3 = st.columns([3, 1.2, 1])
        with _h1:
            st.markdown(f'<div class="h-label">Finance Register · {MY_NAME}</div><div class="h-sub">{title_cat} - Task register</div><div class="h-small">{len(tasks_view)} records · {open_tasks} currently open · Managed by {MY_NAME}</div>', unsafe_allow_html=True)
        with _h2:
            if st.button("Clear Filters", use_container_width=True, key=f"reg_clear_{title_cat}"):
                clear_filters_for(title_cat); st.rerun()
        with _h3:
            if st.button("+ New Record", use_container_width=True, type="primary", key=f"reg_new_{title_cat}"):
                st.session_state.show_form = True; st.rerun()
        SMALL_GAP()
        s1,s2,s3,s4,s5,s6 = st.columns([2,1,1,1,1,1])
        with s1: search = st.text_input("Search", placeholder="Search name, PR, person, budget head...", label_visibility="collapsed", key=f"search_{title_cat}")
        with s2: status_f = st.selectbox("Status", ["All statuses", "Open", "Overdue", "Pending External", "Payment Stage", "Completed"], label_visibility="collapsed", key=f"status_{title_cat}")
        with s3: type_f = st.selectbox("Type", ["All task types", "PRs & procurement", "Procurement", "Reimbursements", "Travel claims", "Tax reviews", "Advances", "Grants", "Local", "RA / TA"], label_visibility="collapsed", key=f"type_{title_cat}")
        with s4: pri_f = st.selectbox("Priority", ["All priorities", "Low", "Medium", "High", "Urgent"], label_visibility="collapsed", key=f"pri_{title_cat}")
        with s5: person_f = st.text_input("Pending With", placeholder="Pending with person...", label_visibility="collapsed", key=f"person_{title_cat}")
        with s6: budget_f = st.text_input("Budget Head", placeholder="Budget head...", label_visibility="collapsed", key=f"budget_{title_cat}")
        filtered = tasks_view
        if search: filtered = [t for t in filtered if search.lower() in str(t).lower()]
        if status_f != "All statuses": filtered = [t for t in filtered if t["Status"] == status_f]
        if pri_f != "All priorities": filtered = [t for t in filtered if t["Priority"] == pri_f]
        if type_f != "All task types": filtered = [t for t in filtered if t["Category"] == type_f]
        if person_f: filtered = [t for t in filtered if person_f.lower() in t.get("Pending With","").lower()]
        if budget_f: filtered = [t for t in filtered if budget_f.lower() in t.get("Budget Head","").lower()]
        SMALL_GAP()
        st.markdown('<div class="reg-head"><span style="flex:1;">RECEIVED</span><span style="flex:1.5;">TASK / REQUESTOR</span><span style="flex:1.5;">TYPE & REFERENCE</span><span style="flex:1.5;">AMOUNT / BUDGET HEAD</span><span style="flex:1.5;">STATUS / PENDING WITH</span><span style="flex:1;">FOLLOW-UP</span><span style="flex:1;">END DATE</span></div>', unsafe_allow_html=True)
        if not filtered:
            st.markdown('<div style="text-align:center;padding:28px 20px 12px;"><div style="font-size:14px;font-weight:700;">Your finance register is ready</div><div style="font-size:12px;color:#8A9A8E;margin-top:4px;">Add your first finance task to start tracking requests and follow-ups.</div></div>', unsafe_allow_html=True)
            b1,b2,b3 = st.columns([3,1,3])
            with b2:
                if st.button("+ Add first task", use_container_width=True, key=f"first_{title_cat}"): st.session_state.show_form = True; st.rerun()
        else:
            for t in filtered:
                st.markdown('<div class="reg-row">', unsafe_allow_html=True)
                c1,c2,c3,c4,c5,c6,c7 = st.columns([1,1.5,1.5,1.5,1.5,1,1])
                c1.write(t["Received"]); c2.write(f'**{t["Task"]}**\n{t["Requestor"]}'); c3.write(f'{t["Category"]}\n{t["Reference"]}'); c4.write(f'PKR {t["Amount"]:,}\n{t["Budget Head"]}'); c5.markdown(status_pill(t["Status"]) + f'<br>{t["Pending With"]}', unsafe_allow_html=True); c6.write(t["Follow-up"]); c7.write(t["End Date"])
                st.markdown('</div>', unsafe_allow_html=True)
                st.divider()
        st.divider()
        st.caption(f"Showing **{len(filtered)}** of **{len(tasks_view)}** records · Prepared by {MY_NAME} · Currency: PKR")
if menu == "📁 Audit / Export":
    GAP()
    st.markdown(f'<div class="h-sub">📁 Audit / Export - {MY_NAME}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="h-text">For Audit: Download the full register as Excel, or upload a backup Excel to restore.</div>', unsafe_allow_html=True)
    GAP()
    with st.container(border=True):
        st.markdown(f'<div class="h-label">AUDIT SUMMARY</div><div class="h-sub">Prepared by {MY_NAME} · {MY_OFFICE}</div><div class="h-small">Date: {date.today().strftime("%d %b %Y")} · Total Records: {len(tasks)} · Total Value: PKR {total_val:,} · Completed: PKR {completed_val:,} · Payments: PKR {payments:,}</div>', unsafe_allow_html=True)
        SMALL_GAP()
        a1,a2 = st.columns(2)
        with a1: st.download_button("📥 Download FULL Excel for Audit (.xlsx)", make_excel_bytes(tasks), f"FinanceDesk_Audit_{MY_NAME}_{date.today()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="audit_xlsx")
        with a2: st.download_button("📄 Download CSV for Audit", make_csv_bytes(tasks), f"FinanceDesk_Audit_{MY_NAME}_{date.today()}.csv", "text/csv", use_container_width=True, key="audit_csv")
        st.caption("Tip: File goes to your Downloads folder. Forward that file on WhatsApp / Email to Audit.")
        st.divider()
        st.markdown('<div class="h-sub">Restore / Upload old Excel</div><div class="h-small">If you have an old FinanceDesk_Data.xlsx, upload it here to bring back all tasks.</div>', unsafe_allow_html=True)
        up = st.file_uploader("Upload Excel to Restore", type=["xlsx","csv"])
        if up:
            try:
                df_up = pd.read_excel(up) if up.name.endswith(".xlsx") else pd.read_csv(up)
                if st.button("✅ Restore these tasks", type="primary", key="restore_btn"):
                    st.session_state.tasks = df_up.to_dict("records"); save_tasks(st.session_state.tasks); st.success(f"Restored {len(df_up)} tasks, {MY_NAME}!"); st.rerun()
            except Exception as e: st.error(f"Could not read file: {e}")
    register_box("All")
elif menu == "🏠 Overview" or is_category_page:
    GAP()
    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f'<div class="h-card"><div class="h-label">Open Tasks</div><div class="h-num">{open_tasks}</div><div class="h-small">Items not yet completed</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="h-card"><div class="h-label">Overdue / Follow-up</div><div class="h-num">{overdue}</div><div class="h-small">Past due or follow-up date reached</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="h-card"><div class="h-label">Pending External</div><div class="h-num">{pending_ext}</div><div class="h-small">With Finance, vendor or another office</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="h-card" style="background:#EDF2EB;"><div class="h-label">Payments in Process</div><div class="h-num">PKR {payments:,}</div><div class="h-small">Value of tasks at payment stage</div></div>', unsafe_allow_html=True)
    GAP(); finance_box(cat if is_category_page else "All", total_val, completed_val, payments)
    if not is_category_page:
        GAP()
        w1,w2 = st.columns(2)
        with w1: st.markdown('<div class="h-card last-box"><div class="h-label">Workload</div><div class="h-sub">Tasks by status</div><div class="h-small">Add tasks to see breakdown by status.</div></div>', unsafe_allow_html=True)
        with w2: st.markdown(f'<div class="h-card last-box"><div class="h-label">Next Actions · {MY_NAME}</div><div class="h-sub">Follow-ups to watch</div><div class="h-small">{"No follow-ups due." if not due_today else f"{len(due_today)} due today - check above!"}</div></div>', unsafe_allow_html=True)
    register_box(cat if is_category_page else "All")
elif menu == "📊 Dashboard":
    GAP()
    st.markdown(f'<div class="h-sub">📊 Dashboard - {MY_NAME}</div>', unsafe_allow_html=True)
    if tasks_view:
        df = pd.DataFrame(tasks_view)
        d1,d2 = st.columns(2)
        with d1: st.markdown('<div class="h-label">Tasks by Status</div>', unsafe_allow_html=True); st.bar_chart(df["Status"].value_counts())
        with d2: st.markdown('<div class="h-label">Amount by Category (PKR)</div>', unsafe_allow_html=True); st.bar_chart(df.groupby("Category")["Amount"].sum())
        GAP(); finance_box("All", total_val, completed_val, payments)
    else: st.info(f"No data yet, {MY_NAME}. Add a task to see charts.")
    register_box("All")
elif menu == "✏️ Edit / Delete":
    GAP()
    st.markdown(f'<div class="h-sub">✏️ Edit / Delete - {MY_NAME}</div>', unsafe_allow_html=True)
    if not tasks: st.info("No tasks yet.")
    else:
        options = [f"{i+1}. {t['Task']} - {t['Category']} - {t['Status']}" for i,t in enumerate(tasks)]
        sel = st.selectbox("Select task to edit / delete", options)
        idx = int(sel.split(".")[0])-1
        t = tasks[idx]
        with st.form("edit"):
            e1,e2,e3 = st.columns(3)
            nt = e1.text_input("Task / Title", value=t["Task"]); nr = e2.text_input("Requestor", value=t["Requestor"]); ncat = e3.selectbox("Category", ["PRs & procurement", "Procurement", "Reimbursements", "Travel claims", "Tax reviews", "Advances", "Grants", "Local", "RA / TA"], index=["PRs & procurement", "Procurement", "Reimbursements", "Travel claims", "Tax reviews", "Advances", "Grants", "Local", "RA / TA"].index(t["Category"]) if t["Category"] in ["PRs & procurement", "Procurement", "Reimbursements", "Travel claims", "Tax reviews", "Advances", "Grants", "Local", "RA / TA"] else 0)
            f1,f2,f3 = st.columns(3)
            namt = f1.number_input("Amount (PKR)", value=int(t["Amount"])); nstat = f2.selectbox("Status", ["Open", "Overdue", "Pending External", "Payment Stage", "Completed"], index=["Open", "Overdue", "Pending External", "Payment Stage", "Completed"].index(t["Status"]) if t["Status"] in ["Open", "Overdue", "Pending External", "Payment Stage", "Completed"] else 0); npend = f3.text_input("Pending With", value=t["Pending With"])
            up = st.text_input("Upload / Proof File Name (optional)", value=t.get("Proof",""), placeholder="e.g. PR-123.pdf")
            c_save, c_del = st.columns(2)
            with c_save:
                if st.form_submit_button("💾 Save Changes", use_container_width=True, type="primary"):
                    tasks[idx].update({"Task": nt, "Requestor": nr, "Category": ncat, "Amount": namt, "Status": nstat, "Pending With": npend, "Proof": up}); save_tasks(tasks); st.success(f"Saved, {MY_NAME}!"); st.rerun()
            with c_del:
                if st.form_submit_button("🗑️ Delete Task", use_container_width=True):
                    tasks.pop(idx); save_tasks(tasks); st.success("Deleted!"); st.rerun()
elif menu == "🖨️ Print View":
    GAP()
    st.markdown(f'<div style="background:#fff; border:1px solid #E9ECE8; border-radius:14px; padding:24px;"><h2 style="font-family:Times New Roman;">FinanceDesk - Print Report</h2><p>Prepared by: <b>{MY_NAME}</b><br>Office: {MY_OFFICE}<br>Date: {date.today().strftime("%d %b %Y")}</p><p>Total Tracked: PKR {total_val:,} | Completed: PKR {completed_val:,} | Payments: PKR {payments:,}</p><p>Open: {open_tasks} | Overdue: {overdue} | Pending External: {pending_ext}</p><p style="color:#8A9A8E; font-size:12px;">Press Ctrl + P in browser to print this page.</p></div>', unsafe_allow_html=True)
    register_box("All", show_export=False)
elif menu in ["📋 All tasks", "➕ New task"] or st.session_state.show_form:
    if menu == "➕ New task":
        st.session_state.show_form = True; st.rerun()
    register_box("All")
st.caption(f"FinanceDesk · {MY_NAME} · {MY_OFFICE} · Auto-saved to {DATA_FILE}")
