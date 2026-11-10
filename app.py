import streamlit as st
import pandas as pd
from datetime import datetime, date
from supabase import create_client

# --- YOUR SUPABASE - PASTE YOUR KEY HERE ---
SUPABASE_URL = "https://opikmenodygpotuuoxdi.supabase.co"
SUPABASE_KEY = "sb_publishable_3ts27IR5LdfBpQYrLxn7MA_XwfP4X6_"
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

st.set_page_config(page_title="Finance Desk - Anum", layout="wide")
st.title("💰 Finance Desk - Anum")

# --- LOAD TASKS FROM SUPABASE ---
def load_tasks():
    try:
        res = supabase.table("tasks").select("*").order("id").execute()
        tasks = []
        for row in res.data:
            d = row["data"]
            d["_id"] = row["id"]
            tasks.append(d)
        return tasks
    except Exception as e:
        st.error(f"Load error: {e}")
        return []

def save_task(task):
    supabase.table("tasks").insert({"data": task}).execute()

def delete_task(task_id):
    supabase.table("tasks").delete().eq("id", task_id).execute()

def update_task(task_id, task):
    supabase.table("tasks").update({"data": task}).eq("id", task_id).execute()

if "tasks" not in st.session_state:
    st.session_state.tasks = load_tasks()

# --- ADD TASK FORM ---
st.subheader("➕ Add New Task")
with st.form("add_task", clear_on_submit=True):
    c1, c2, c3 = st.columns(3)
    received = c1.text_input("Received From")
    task_text = c2.text_input("Task / Description")
    amount = c3.number_input("Amount (Rs)", min_value=0.0, step=100.0)
    c4, c5, c6 = st.columns(3)
    task_date = c4.date_input("Date", value=date.today())
    status = c5.selectbox("Status", ["Pending", "In Progress", "Done", "Paid"])
    remarks = c6.text_input("Remarks")
    submitted = st.form_submit_button("💾 Save Task", use_container_width=True)
    if submitted and task_text:
        new_task = {
            "Received": received,
            "Task": task_text,
            "Amount": float(amount),
            "Date": str(task_date),
            "Status": status,
            "Remarks": remarks,
        }
        save_task(new_task)
        st.session_state.tasks = load_tasks()
        st.success("Saved to Supabase! ✅")
        st.rerun()

# --- SHOW TASKS ---
st.subheader(f"📋 All Tasks ({len(st.session_state.tasks)}) - Saved in Supabase ☁️")
if st.session_state.tasks:
    df = pd.DataFrame(st.session_state.tasks)
    # Hide internal id for display
    show_cols = [c for c in df.columns if c != "_id"]
    st.dataframe(df[show_cols], use_container_width=True, hide_index=True)
    total_pending = df[df["Status"] != "Paid"]["Amount"].sum() if "Amount" in df.columns else 0
    total_paid = df[df["Status"] == "Paid"]["Amount"].sum() if "Amount" in df.columns else 0
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Tasks", len(df))
    col2.metric("Pending Amount", f"Rs {total_pending:,.0f}")
    col3.metric("Paid Amount", f"Rs {total_paid:,.0f}")

    # Delete
    st.markdown("---")
    st.subheader("🗑️ Delete Task")
    to_delete = st.selectbox("Select task to delete", options=range(len(st.session_state.tasks)), format_func=lambda i: f"{st.session_state.tasks[i].get('Task','')} - {st.session_state.tasks[i].get('Received','')}")
    if st.button("Delete Selected"):
        delete_task(st.session_state.tasks[to_delete]["_id"])
        st.session_state.tasks = load_tasks()
        st.rerun()
else:
    st.info("No tasks yet. Add your first task above! Your data will now save FOREVER in Supabase!")

st.caption("V52 - Supabase Cloud Save | Finance-Desk-Anum")
