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
master_df = pd.read_excel(
    "Service_TAT_Annexure.xlsx"
)

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

    report_df = pd.read_excel(uploaded_file)
    report_df.columns = report_df.columns.str.strip().str.lower()

    # =========================
    # REMOVE DUPLICATE COLUMN NAMES
    # =========================
    report_df = report_df.loc[:, ~report_df.columns.duplicated()]

    # =========================
    # HANDLE DUPLICATE COLUMN (AE, AF)
    # =========================
    cols = list(report_df.columns)

    # AE = Query Category
    if len(cols) >= 31:
        report_df['query category (cs, escalation & gro)'] = report_df.iloc[:, 30]

    # AF = Query Sub-category
    if len(cols) >= 32:
        report_df['query sub-category'] = report_df.iloc[:, 31]

    # =========================
    # RENAME
    # =========================
    report_df.rename(columns={
        'query sub-category': 'sub_category',
        'query category (cs, escalation & gro)': 'query_category',
        'agent': 'agent',
        'name type': 'type',
        'created time': 'created_time',
        'resolved time': 'resolved_time',
        'description': 'description',
        'ticket id': 'ticket_id'
    }, inplace=True)

    # =========================
    # REMOVE DUPLICATE TICKET ID
    # =========================
    if 'ticket_id' in report_df.columns:
        report_df = report_df.drop_duplicates(subset=['ticket_id'])

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
    report_df['created_time'] = pd.to_datetime(report_df['created_time'], errors='coerce')
    report_df['resolved_time'] = pd.to_datetime(report_df['resolved_time'], errors='coerce')

    # =========================
    # MERGE WITH MASTER
    # =========================
    df = report_df.merge(
        master_df[['sub_category','qrc_type','tat_days']],
        on='sub_category',
        how='left',
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
    # COLLECTION ISSUE → COMPLAINT OVERRIDE
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
    # NEW LOGIC – Cancellation of loan after disbursal
    # =====================================================
    def override_qrc_and_tat(row):

        if str(row.get('sub_category', '')).strip().lower() != 'cancellation of loan after disbursal':
            return row['qrc_type'], row['tat_days']

        mail = str(row.get('description', '')).lower()

        # IGNORE DISCLAIMER PART
        mail = re.split(r'disclaimer', mail, flags=re.IGNORECASE)[0]

        complaint_keywords = [
            'fraud', 'fraudulent', 'without consent',
            'mis-selling', 'misselling of course','misselling of loan',
            'fir', 'police','fake promisses', 'job gurantee',
            'dpd','delayed payment','undue recovery practice',
            'collection harassment','agent misbehaviour','abuse',
            'dispute','harassment','fake',
            'cibil rectification','cibil dispute','cibil issue'
        ]

        for word in complaint_keywords:
            if word in mail:
                return 'Complaint', 15

        return 'Request', 2

    df[['qrc_type', 'tat_days']] = df.apply(
        lambda x: pd.Series(override_qrc_and_tat(x)),
        axis=1
    )

    # =========================
    # ACTUAL TAT (DAYS)
    # =========================
    df['actual_tat_days'] = (
        (df['resolved_time'] - df['created_time'])
        .dt.total_seconds() / 86400
    ).round(2)

    # =========================
    # TAT STATUS LOGIC
    # =========================
    def tat_status_logic(row):
        if pd.isna(row['resolved_time']):
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



