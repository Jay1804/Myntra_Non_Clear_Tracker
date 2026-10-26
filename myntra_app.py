import os
from io import BytesIO

import pandas as pd
import pymysql
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

st.set_page_config(page_title="Myntra Case Report", layout="wide")

CLIENT_EXTERNAL_ID = 6970

QUERY = """
select
Distinct ecm.case_ars_no,
ecff.case_flex_field1,
concat_ws(' ', ecc1.first_name, ecc1.middle_name, ecc1.last_name) as Candidate_Name,
checkpoint_live.fn_case_status(case_status) as case_status,

-- Education Verification
 (select coalesce(check_severity,'NA')
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id and
            check_id IN (25,83) and check_status <> 9
             limit 1) as `Education Verification severity` ,

  (select checkpoint_live.fn_check_status(Check_status)
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id and
            check_id IN (25,83) and check_status <> 9 limit 1) `Education Verification check status`,

(select group_concat(concat_ws(' -> ', ecc2.check_name, ecc2.check_severity, ecc2.closure_comments) order by ecc2.check_name separator ',')
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id
        and check_id IN (25,83) and check_status <> 9
        ) as `Education Verification Non-Green Comments`,

 (select coalesce(check_severity,'NA')
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id and
            check_id IN (58) and check_status <> 9
             limit 1) as `Current Employment Verification severity` ,

  (select checkpoint_live.fn_check_status(Check_status)
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id and
            check_id IN (58) and check_status <> 9 limit 1) `Current Employment Verification check status`,

(select group_concat(concat_ws(' -> ', ecc2.check_name, ecc2.check_severity, ecc2.closure_comments) order by ecc2.check_name separator ',')
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id
        and check_id IN (58) and check_status <> 9
        ) as `Current Employment Verification Non-Green Comments`,

-- Previous Employment Verification
 (select coalesce(check_severity,'NA')
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id and
            check_id IN (41) and check_status <> 9
             limit 1) as `Previous Employment Verification severity` ,

  (select checkpoint_live.fn_check_status(Check_status)
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id and
            check_id IN (41) and check_status <> 9 limit 1) `Previous Employment Verification check status`,

(select group_concat(concat_ws(' -> ', ecc2.check_name, ecc2.check_severity, ecc2.closure_comments) order by ecc2.check_name separator ',')
    from ec_case_checks ecc2
        where ecm.case_id = ecc2.case_id
        and check_id IN (41) and check_status <> 9
        ) as `Previous Employment Verification Non-Green Comments`,

 DATE_FORMAT((SELECT ecr.report_sent_on FROM ec_case_reports ecr
    WHERE ecm.case_id = ecr.case_id AND ecr.report_type = 1
        ORDER BY ecr.report_sent_on DESC LIMIT 1),'%Y/%m/%d %h:%i %p') AS Final_Report_Sent


FROM ec_case_master ecm
left join ec_case_candidates ecc1 on ecc1.candidate_id=ecm.candidate_id
left join ec_client ec1 on ecm.client_id = ec1.client_id
LEFT JOIN ec_case_fields ecff ON ecm.case_id=ecff.case_id


where
ec1.client_external_id in (6970)
AND ecc1.first_name NOT LIKE '%test%'
    AND ecc1.first_name NOT LIKE '%dummy%'
    and ecm.received_date > '2025-01-01'
Order by ecm.case_ars_no
"""


# Output column headers, in the same order as the SELECT list
COLUMN_NAMES = [
    "case_ars_no",
    "EMP ID",
    "candidate_name",
    "BGV Current STATUS",
    "education verification severity",
    "education verification check status",
    "education verification non-green comments",
    "current employment verification severity",
    "current employment verification check status",
    "current employment verification non-green comments",
    "previous employment verification severity",
    "previous employment verification check status",
    "previous employment verification non-green comments",
    "Final completion date (GREEN/ RED)",
]


def get_connection():
    return pymysql.connect(
        host=os.getenv("DB_HOST"),
        port=int(os.getenv("DB_PORT", "3306")),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME", "checkpoint_live"),
        charset="utf8mb4",
        connect_timeout=30,
    )


@st.cache_data(show_spinner="Querying database...", ttl=600)
def load_data() -> pd.DataFrame:
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(QUERY)
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
    finally:
        conn.close()
    df = pd.DataFrame(rows, columns=cols)
    df.columns = [COLUMN_NAMES[i] for i in range(len(cols))]
    return clean_emp_id(df)


def clean_emp_id(df: pd.DataFrame) -> pd.DataFrame:
    """Store numeric EMP IDs as numbers (no 'number stored as text' error) and blanks as empty."""
    df = df.copy()
    ids = df["EMP ID"].astype(str).str.strip()
    df["EMP ID"] = [int(v) if v.isdigit() else (v or None) for v in ids]
    return df


def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Myntra")
        ws = writer.sheets["Myntra"]
        side = Side(style="thin", color="000000")
        border = Border(left=side, right=side, top=side, bottom=side)
        header_fill = PatternFill("solid", fgColor="D9D9D9")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                if cell.row > 1:
                    cell.alignment = Alignment(vertical="top", horizontal="left")
        for i, col in enumerate(df.columns, start=1):
            longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
            ws.column_dimensions[get_column_letter(i)].width = min(max(longest + 2, 12), 50)
        ws.freeze_panes = "A2"
    return buf.getvalue()


st.title("Myntra Case Verification Report")
st.caption(f"Client external ID {CLIENT_EXTERNAL_ID} · cases received after 2025-01-01")

if st.sidebar.button("Refresh data"):
    load_data.clear()

try:
    df = load_data()
except Exception as e:
    st.error(f"Database query failed: {e}")
    st.stop()

# Filters
status_col = "BGV Current STATUS"
search = st.sidebar.text_input("Search (ARS no / candidate name)")
if status_col:
    statuses = sorted(df[status_col].dropna().astype(str).unique())
    chosen = st.sidebar.multiselect("Case status", statuses)
    if chosen:
        df = df[df[status_col].astype(str).isin(chosen)]
if search:
    mask = (
        df["case_ars_no"].astype(str).str.contains(search, case=False, na=False)
        | df["candidate_name"].astype(str).str.contains(search, case=False, na=False)
    )
    df = df[mask]

c1, c2 = st.columns(2)
c1.metric("Cases", len(df))
c2.metric("Final report sent", int(df["Final completion date (GREEN/ RED)"].notna().sum()))

styled = df.style.set_properties(**{"border": "1px solid #000000"}).format({"EMP ID": lambda v: "" if pd.isna(v) else str(int(v)) if isinstance(v, (int, float)) else str(v)})
st.dataframe(styled, use_container_width=True, hide_index=True)

d1, d2 = st.columns(2)
d1.download_button("Download CSV", df.to_csv(index=False).encode("utf-8"),
                   "myntra_report.csv", "text/csv")
try:
    d2.download_button("Download Excel", to_excel(df), "myntra_report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
except ImportError:
    pass
