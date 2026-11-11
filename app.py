import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Finance Desk - Anum", layout="wide")
st.title("💰 Finance Desk - Anum")
st.write("Your private work desk - Upload Excel, no copy-paste!")

# --- UPLOAD BLOCK ---
uploaded = st.file_uploader("Upload your Excel file", type=["xlsx", "csv"])

if uploaded is not None:
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)

    df.columns = df.columns.str.strip()

    # --- FIX AMOUNT: handles 40,000 and 14,000*6 ---
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

    df["Amount_Clean"] = df["Amount"].apply(clean_amount)

    # --- TOP SUMMARY BLOCKS ---
    st.subheader("📊 Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Records", len(df))
    c2.metric("Total Amount", f"{df['Amount_Clean'].sum():,.0f}")
    c3.metric("Processed", (df['Status'].astype(str).str.lower() == 'processed').sum())
    c4.metric("Pending / Others", (df['Status'].astype(str).str.lower() != 'processed').sum())

    st.divider()

    # --- FILTER BLOCKS ---
    st.subheader("🔍 Search & Filter")
    f1, f2, f3 = st.columns(3)
    search = f1.text_input("Search by Name or Title")
    status_list = ["All"] + sorted(df['Status'].dropna().astype(str).unique().tolist())
    status_filter = f2.selectbox("Filter by Status", status_list)
    
    # PR Approver Filter - YOU ASKED FOR THIS
    approver_list = ["All"] + sorted(df['PR Approver'].dropna().astype(str).unique().tolist())
    approver_filter = f3.selectbox("Filter by PR Approver", approver_list)

    f4, f5 = st.columns(2)
    sap_list = ["All"] + sorted(df['SAP Holder'].dropna().astype(str).unique().tolist())
    sap_filter = f4.selectbox("Filter by SAP Holder", sap_list)
    # Budget head filter if column exists
    if 'Budget head' in df.columns:
        budget_list = ["All"] + sorted(df['Budget head'].dropna().astype(str).unique().tolist())
        budget_filter = f5.selectbox("Filter by Budget Head", budget_list)
    else:
        budget_filter = "All"

    # Apply filters
    filtered = df.copy()
    if search:
        filtered = filtered[
            filtered['Name'].astype(str).str.contains(search, case=False, na=False) |
            filtered['Title'].astype(str).str.contains(search, case=False, na=False)
        ]
    if status_filter != "All":
        filtered = filtered[filtered['Status'].astype(str) == status_filter]
    if approver_filter != "All":
        filtered = filtered[filtered['PR Approver'].astype(str) == approver_filter]
    if sap_filter != "All":
        filtered = filtered[filtered['SAP Holder'].astype(str) == sap_filter]
    if budget_filter != "All":
        filtered = filtered[filtered['Budget head'].astype(str) == budget_filter]

    st.write(f"Showing **{len(filtered)}** records | Total of filtered: **{filtered['Amount_Clean'].sum():,.0f}**")

    # --- MAIN TABLE BLOCK - Excel direct in website ---
    st.subheader("📋 All Records")
    show_cols = [c for c in ['Name','Title','Date','Amount','PR','Budget head','SAP Holder','PR Approver','Status'] if c in df.columns]
    st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)

    # --- DOWNLOAD BUTTON - YOU ASKED FOR THIS ---
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

    # --- DETAIL VIEW BLOCK ---
    st.divider()
    st.subheader("👤 Detail View")
    selected_name = st.selectbox("Select Name to see full detail", sorted(df['Name'].dropna().astype(str).unique().tolist()))
    detail = df[df['Name'].astype(str) == selected_name][show_cols]
    st.table(detail)

else:
    st.info("👆 Please upload your Excel file to see your data here")
    st.write("Your Excel should have columns: Name, Title, Date, Amount, PR, Budget head, SAP Holder, PR Approver, Status")