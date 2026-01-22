import streamlit as st
import pandas as pd
from io import BytesIO
import re

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="QRC Dashboard", layout="wide")
st.title("📊 QRC Dashboard")

# =========================
# LOAD MASTER DATA
# =========================
master_df = pd.read_excel("Service_TAT_Annexure.xlsx")

master_df.columns = master_df.columns.str.strip().str.lower()

master_df.rename(columns={
    'query sub-category': 'sub_category',
    'tat': 'tat_days',
    'tat (days)': 'tat_days',
    'qrc type': 'qrc_type',
    'process': 'process'
}, inplace=True)

# =========================
# UPLOAD REPORT FILE
# =========================
uploaded_file = st.file_uploader(
    "📤 Upload Report File (.xls / .xlsx)",
    type=["xls", "xlsx"]
)

if uploaded_file:

    # =========================
    # READ REPORT FILE
    # =========================
    report_df = pd.read_excel(uploaded_file)
    report_df.columns = report_df.columns.str.strip().str.lower()

    report_df.rename(columns={
        'query sub-category': 'sub_category',
        'query category (cs, escalation & gro)': 'query_category',
        'agent': 'agent',
        'name type': 'type',
        'created ti': 'created_ti',
        'resolved t': 'resolved_t',
        'description': 'description'
    }, inplace=True)

    # =========================
    # NORMALIZE QUERY CATEGORY
    # =========================
    report_df['query_category'] = (
        report_df['query_category']
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # =========================
    # SOURCE FILTER – EMAIL ONLY
    # =========================
    if 'source' in report_df.columns:
        report_df['source'] = report_df['source'].astype(str).str.strip().str.lower()
        report_df = report_df[report_df['source'] == 'email']
    else:
        st.warning("⚠ 'Source' column not found in uploaded file")

    # =========================
    # DATETIME CONVERSION
    # =========================
    report_df['created_ti'] = pd.to_datetime(report_df['created_ti'], errors='coerce')
    report_df['resolved_t'] = pd.to_datetime(report_df['resolved_t'], errors='coerce')

    # =========================
    # MERGE WITH MASTER (OLD LOGIC)
    # =========================
    df = report_df.merge(
        master_df[['sub_category', 'tat_days', 'qrc_type']],
        on='sub_category',
        how='left'
    )

    # =========================
    # BLANK SUB-CATEGORY HANDLING
    # =========================
    mask_blank_subcat = (
        df['sub_category'].isna()
        | (df['sub_category'].astype(str).str.strip() == "")
    )
    df.loc[mask_blank_subcat, ['tat_days', 'qrc_type']] = pd.NA

    # =========================
    # COLLECTION ISSUE → COMPLAINT OVERRIDE (OLD)
    # =========================
    mask_collection_issue = (
        (df['query_category'] == 'collection issue') &
        (
            df['sub_category'].isna()
            | (df['sub_category'].astype(str).str.strip() == "")
        )
    )
    df.loc[mask_collection_issue, 'qrc_type'] = 'Complaint'

    # =====================================================
    # 🔥 NEW LOGIC – ONLY FOR "Cancellation of loan after disbursal"
    # =====================================================
    def override_qrc_and_tat(row):

        if str(row.get('sub_category', '')).strip().lower() != 'cancellation of loan after disbursal':
            return row['qrc_type'], row['tat_days']   # 👈 PURANA LOGIC SAFE

        mail = str(row.get('description', '')).lower()

        # STRONG COMPLAINT SIGNALS
        complaint_keywords = [
            'fraud', 'fraudulent', 'unauthorized', 'without consent',
            'mis-selling', 'misselling', 'cheated', 'scam',
            'legal', 'legal notice', 'lawyer', 'court',
            'consumer', 'ombudsman', 'rbi', 'fir', 'police',
            'dispute', 'disputed', 'harassment', 'illegal', 'fake'
        ]

        for word in complaint_keywords:
            if word in mail:
                return 'Complaint', 15   # ✅ Complaint = 15 days

        
                return 'Request', 2      # ✅ Request = 2 days

        
    # 👇 ONLY THIS LINE IS DIFFERENT
    df[['qrc_type', 'tat_days']] = df.apply(
        lambda x: pd.Series(override_qrc_and_tat(x)),
        axis=1
    )

    # =========================
    # ACTUAL TAT (DAYS)
    # =========================
    df['actual_tat_days'] = (
        (df['resolved_t'] - df['created_ti'])
        .dt.total_seconds() / 86400
    ).round(2)

    # =========================
    # TAT STATUS LOGIC
    # =========================
    def tat_status_logic(row):
        if pd.isna(row['resolved_t']):
            return "Unresolved"
        elif pd.notna(row['tat_days']) and row['actual_tat_days'] <= row['tat_days']:
            return "Within TAT"
        else:
            return "Out of TAT"

    df['tat_status'] = df.apply(tat_status_logic, axis=1)

    # =========================
    # DASHBOARD 1 – OVERALL
    # =========================
    total = len(df)
    within = (df['tat_status'] == "Within TAT").sum()
    out = (df['tat_status'] == "Out of TAT").sum()
    unresolved = (df['tat_status'] == "Unresolved").sum()
    resolved_total = within + out

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cases", total)
    c2.metric("Within TAT", f"{within} ({round((within / resolved_total) * 100, 2)}%)" if resolved_total else "0 (0%)")
    c3.metric("Out of TAT", f"{out} ({round((out / resolved_total) * 100, 2)}%)" if resolved_total else "0 (0%)")
    c4.metric("Unresolved", f"{unresolved} ({round((unresolved / total) * 100, 2)}%)" if total else "0 (0%)")

    st.divider()

    # =========================
    # DASHBOARD 2 – AGENT WISE
    # =========================
    st.subheader("👤 Agent-wise TAT")

    agent_df = (
        df.groupby('agent')
        .agg(
            total_cases=('tat_status', 'count'),
            within_tat=('tat_status', lambda x: (x == "Within TAT").sum()),
            out_of_tat=('tat_status', lambda x: (x == "Out of TAT").sum()),
            unresolved=('tat_status', lambda x: (x == "Unresolved").sum())
        )
        .reset_index()
    )

    agent_df['within_tat_%'] = round(
        (agent_df['within_tat'] / (agent_df['within_tat'] + agent_df['out_of_tat'])) * 100,
        2
    )

    st.dataframe(agent_df, use_container_width=True)

    st.divider()

    # =========================
    # DASHBOARD 3 – TYPE vs QRC
    # =========================
    st.subheader("🔍 Type vs QRC Type Match")

    df['type_match_status'] = df.apply(
        lambda x: "Match"
        if str(x['type']).strip().lower() == str(x['qrc_type']).strip().lower()
        else "Mismatch",
        axis=1
    )

    match_df = df[['sub_category', 'type', 'qrc_type', 'type_match_status']]
    st.dataframe(match_df, use_container_width=True)

    st.divider()

    # =========================
    # DASHBOARD 4 – SUB CATEGORY
    # =========================
    st.subheader("📌 Query Sub-category Summary")

    subcat_df = (
        df.groupby('sub_category', dropna=False)
        .agg(
            total_cases=('tat_status', 'count'),
            within_tat=('tat_status', lambda x: (x == "Within TAT").sum()),
            out_of_tat=('tat_status', lambda x: (x == "Out of TAT").sum()),
            unresolved=('tat_status', lambda x: (x == "Unresolved").sum())
        )
        .reset_index()
    )

    subcat_df['%_of_total'] = round((subcat_df['total_cases'] / total) * 100, 2)
    st.dataframe(subcat_df, use_container_width=True)

    st.divider()

    # =========================
    # EXCEL DOWNLOAD
    # =========================
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Full_Data", index=False)
        agent_df.to_excel(writer, sheet_name="Agent_Wise_TAT", index=False)
        match_df.to_excel(writer, sheet_name="Type_vs_QRC", index=False)
        subcat_df.to_excel(writer, sheet_name="Sub_Category_Summary", index=False)

    output.seek(0)

    st.download_button(
        "⬇ Download Complete TAT Excel Report",
        data=output,
        file_name="TAT_Complete_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

