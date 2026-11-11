import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Finance Desk - Anum", layout="wide")
st.title("💰 Finance Desk - Anum")
st.write("Your private work desk - Upload Excel, no copy-paste!")

uploaded = st.file_uploader("Upload your Excel file", type=["xlsx", "csv"])

if uploaded is not None:
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)

    # --- FIX: Auto-fix column names ---
    # Remove spaces, make it easy to find
    df.columns = [str(c).strip() for c in df.columns]
    # Show what columns we found - for debugging
    st.write("Columns found in your Excel:", list(df.columns))

    # Find the real column names automatically
    def find_col(keywords):
        for c in df.columns:
            low = c.lower()
            for k in keywords:
                if k in low:
                    return c
        return None

    col_name = find_col(["name"])
    col_title = find_col(["title"])
    col_date = find_col(["date"])
    col_amount = find_col(["amount", "amt", "rs", "price"])
    col_pr = find_col(["pr"])
    col_budget = find_col(["budget"])
    col_sap = find_col(["sap"])
    col_approver = find_col(["approver", "approve"])
    col_status = find_col(["status"])

    # --- FIX AMOUNT ---
    def clean_amount(x):
        try:
            if pd.isna(x):
                return 0
            s = str(x).replace(",", "").replace(" ", "")
            if "*" in s:
                parts = s.split("*")
                val = 1
                for p in parts:
                    val *= float(p)
                return val
            return float(s)
        except:
            return 0

    if col_amount:
        df["Amount_Clean"] = df[col_amount].apply(clean_amount)
    else:
        df["Amount_Clean"] = 0
        st.warning("Amount column not found - check your Excel header")

    # --- TOP SUMMARY ---
    st.subheader("📊 Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Records", len(df))
    c2.metric("Total Amount", f"{df['Amount_Clean'].sum():,.0f}")
    if col_status:
        c3.metric("Processed", (df[col_status].astype(str).str.lower() == 'processed').sum())
        c4.metric("Pending / Others", (df[col_status].astype(str).str.lower() != 'processed').sum())
    else:
        c3.metric("Processed", 0)
        c4.metric("Pending", 0)

    st.divider()

    # --- FILTERS ---
    st.subheader("🔍 Search & Filter")
    f1, f2, f3 = st.columns(3)
    search = f1.text_input("Search by Name or Title")
    status_filter = "All"
    if col_status:
        status_list = ["All"] + sorted(df[col_status].dropna().astype(str).unique().tolist())
        status_filter = f2.selectbox("Filter by Status", status_list)
    approver_filter = "All"
    if col_approver:
        approver_list = ["All"] + sorted(df[col_approver].dropna().astype(str).unique().tolist())
        approver_filter = f3.selectbox("Filter by PR Approver", approver_list)

    f4, f5 = st.columns(2)
    sap_filter = "All"
    if col_sap:
        sap_list = ["All"] + sorted(df[col_sap].dropna().astype(str).unique().tolist())
        sap_filter = f4.selectbox("Filter by SAP Holder", sap_list)
    budget_filter = "All"
    if col_budget:
        budget_list = ["All"] + sorted(df[col_budget].dropna().astype(str).unique().tolist())
        budget_filter = f5.selectbox("Filter by Budget Head", budget_list)

    # Apply filters
    filtered = df.copy()
    if search:
        mask = pd.Series(False, index=filtered.index)
        if col_name:
            mask = mask | filtered[col_name].astype(str).str.contains(search, case=False, na=False)
        if col_title:
            mask = mask | filtered[col_title].astype(str).str.contains(search, case=False, na=False)
        filtered = filtered[mask]
    if status_filter != "All" and col_status:
        filtered = filtered[filtered[col_status].astype(str) == status_filter]
    if approver_filter != "All" and col_approver:
        filtered = filtered[filtered[col_approver].astype(str) == approver_filter]
    if sap_filter != "All" and col_sap:
        filtered = filtered[filtered[col_sap].astype(str) == sap_filter]
    if budget_filter != "All" and col_budget:
        filtered = filtered[filtered[col_budget].astype(str) == budget_filter]

    st.write(f"Showing **{len(filtered)}** records | Total of filtered: **{filtered['Amount_Clean'].sum():,.0f}**")

    # --- MAIN TABLE ---
    st.subheader("📋 All Records")
    show_cols = [c for c in [col_name, col_title, col_date, col_amount, col_pr, col_budget, col_sap, col_approver, col_status] if c is not None]
    st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)

    # --- DOWNLOAD ---
    st.subheader("⬇️ Download")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered[show_cols].to_excel(writer, index=False, sheet_name='Filtered Data')
    st.download_button(
        label="Download Filtered Data as Excel",
        data=output.getvalue(),
        file_name="Finance_Desk_Filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- DETAIL VIEW ---
    st.divider()
    st.subheader("👤 Detail View")
    if col_name:
        selected_name = st.selectbox("Select Name to see full detail", sorted(df[col_name].dropna().astype(str).unique().tolist()))
        detail = df[df[col_name].astype(str) == selected_name][show_cols]
        st.table(detail)

else:
    st.info("👆 Please upload your Excel file to see your data here")