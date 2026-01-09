import streamlit as st
import pandas as pd
import io
from datetime import datetime

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors


# ================= CONFIG =================
SALES_SHEET = "MAIN_COPY"
TARGET_SHEET = "MARKETING TARGET"

USERS = {
    "admin":   {"password": "admin@123",   "marketing": "ALL"},
    "ashok":   {"password": "ashok@123",   "marketing": "Ashok Marketing"},
    "suresh":  {"password": "suresh@123",  "marketing": "Suresh - Marketing"},
    "ho":      {"password": "ho@123",      "marketing": "H O - Marketing"},
}

MONTH_MAP = {
    "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7,
    "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11,
    "DEC": 12, "JAN": 1, "FEB": 2, "MAR": 3
}
# =========================================

st.set_page_config(page_title="Marketing Sales Dashboard", layout="wide")

# ============ SHARED CACHE ============
@st.cache_data(show_spinner="Loading Excel data...")
def load_excel_cached(file_bytes):
    sales_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SALES_SHEET)
    target_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=TARGET_SHEET)
    return sales_df, target_raw

# ============ PDF ============
def generate_pdf(marketing_name, df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20)
    styles = getSampleStyleSheet()
    elements = []

    # ---------- HEADER ----------
    elements.append(
        Paragraph(
            f"<b>SALES PERFORMANCE REPORT (2025–2026)</b><br/><br/>"
            f"<b>Marketing Person:</b> {marketing_name}",
            styles["Title"]
        )
    )

    elements.append(Paragraph("<br/>", styles["Normal"]))

    # ---------- SUMMARY ----------
    total_target = df["Target"].sum()
    total_sales = df["sales"].sum()
    not_achieved = total_target - total_sales
    achievement_pct = (total_sales / total_target * 100) if total_target > 0 else 0

    summary_table = Table(
        [
            ["TARGET", "ACHIEVED", "NOT ACHIEVED", "%"],
            [
                f"₹ {total_target:,.0f}",
                f"₹ {total_sales:,.0f}",
                f"₹ {not_achieved:,.0f}",
                f"{achievement_pct:.1f} %"
            ],
        ],
        colWidths=[110, 110, 110, 60]
    )

    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (1, 1), colors.lightgreen),
        ("BACKGROUND", (2, 1), (2, 1), colors.salmon),
        ("BACKGROUND", (3, 1), (3, 1), colors.khaki),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))

    elements.append(summary_table)
    elements.append(Paragraph("<br/><br/>", styles["Normal"]))

    # ---------- DETAIL TABLE ----------
    table_data = [["Month", "Target", "Sales", "Achievement %"]]

    for _, r in df.iterrows():
        table_data.append([
            r["Month_Text"],
            f"₹ {r['Target']:,.0f}",
            f"₹ {r['sales']:,.0f}",
            f"{r['Achievement_%']} %"
        ])

    detail_table = Table(table_data, repeatRows=1)
    detail_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]))

    elements.append(detail_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============ LOGIN ============
def login():
    st.title("🔐 Marketing Login")
    user = st.text_input("Username")
    pwd = st.text_input("Password", type="password")

    if st.button("Login"):
        if user in USERS and USERS[user]["password"] == pwd:
            st.session_state["user"] = user
            st.session_state["marketing"] = USERS[user]["marketing"]
            st.rerun()
        else:
            st.error("Invalid username or password")

# ============ DASHBOARD ============
def dashboard():
    marketing = st.session_state["marketing"]
    is_admin = (marketing == "ALL")

    st.sidebar.success(f"Logged in as: {marketing}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()

    # ---------- ADMIN UPLOAD ----------
    if is_admin:
        uploaded_file = st.file_uploader("📤 Admin: Upload Marketing Excel File", type=["xlsx"])
        if uploaded_file:
            st.session_state["file_bytes"] = uploaded_file.getvalue()
            st.session_state["last_updated"] = datetime.now()

    if "file_bytes" not in st.session_state:
        st.warning("⚠️ Admin has not uploaded the Excel file yet")
        st.stop()

    sales_df, target_raw = load_excel_cached(st.session_state["file_bytes"])

    if "last_updated" in st.session_state:
        st.caption(f"📅 Data last updated on: {st.session_state['last_updated'].strftime('%d-%m-%Y %H:%M:%S')}")

    # ---------- TARGET ----------
    target_raw.columns = target_raw.columns.str.strip()
    target_df = target_raw.melt(
        id_vars=["Marketing Person"],
        var_name="Month",
        value_name="Target"
    ).rename(columns={"Marketing Person": "MARK"})

    target_df["Target"] = (
        target_df["Target"].astype(str)
        .str.replace("₹", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    target_df["Target"] = pd.to_numeric(target_df["Target"], errors="coerce").fillna(0)

    target_df["Month_No"] = target_df["Month"].map(MONTH_MAP)
    target_df["Year"] = target_df["Month_No"].apply(lambda x: 2025 if x >= 4 else 2026)
    target_df["YearMonth"] = target_df["Year"] * 100 + target_df["Month_No"]

    # ---------- CLEAN & FILTER ----------
    sales_df["MARK"] = sales_df["MARK"].astype(str).str.strip().str.upper()
    target_df["MARK"] = target_df["MARK"].astype(str).str.strip().str.upper()

    sales_df = sales_df[sales_df["HELPER"].isin(["NOFILL", "GREEN"])]
    valid_marketers = target_df["MARK"].unique()
    sales_df = sales_df[sales_df["MARK"].isin(valid_marketers)]

    selected_marketing = "ALL"
    if is_admin:
        selected_marketing = st.selectbox("📌 Select Marketing Person", ["ALL"] + sorted(valid_marketers))
        if selected_marketing != "ALL":
            sales_df = sales_df[sales_df["MARK"] == selected_marketing]
            target_df = target_df[target_df["MARK"] == selected_marketing]
    else:
        m = marketing.strip().upper()
        sales_df = sales_df[sales_df["MARK"] == m]
        target_df = target_df[target_df["MARK"] == m]

    # ---------- CALCULATION ----------
    monthly_sales = sales_df.groupby(["MARK", "YearMonth", "Month_Text"], as_index=False)["sales"].sum()

    monthly_report = pd.merge(
        monthly_sales,
        target_df[["MARK", "YearMonth", "Target"]],
        on=["MARK", "YearMonth"],
        how="left"
    )

    monthly_report["Target"] = monthly_report["Target"].fillna(0)
    monthly_report["Achievement_%"] = (monthly_report["sales"] / monthly_report["Target"] * 100).round(1)
    monthly_report = monthly_report.sort_values(["MARK", "YearMonth"])

    # ================= UI =================
    st.title("📊 SALES PERFORMANCE DASHBOARD")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Target", f"₹ {monthly_report['Target'].sum():,.0f}")
    col2.metric("Total Sales", f"₹ {monthly_report['sales'].sum():,.0f}")
    col3.metric("Achievement %",
        f"{(monthly_report['sales'].sum() / monthly_report['Target'].sum() * 100):.1f} %"
        if monthly_report["Target"].sum() > 0 else "0 %"
    )

    st.subheader("📊 Month-wise Target vs Sales")
    if is_admin and selected_marketing == "ALL":
        chart_df = monthly_report.groupby("Month_Text", as_index=False)[["Target", "sales"]].sum().set_index("Month_Text")
    else:
        chart_df = monthly_report.set_index("Month_Text")[["Target", "sales"]]
    st.bar_chart(chart_df)

    st.subheader("📋 Month-wise Sales Performance")
    st.dataframe(
        monthly_report[["MARK", "Month_Text", "Target", "sales", "Achievement_%"]]
        .rename(columns={"MARK": "Marketing Person", "sales": "Sales Achieved"}),
        use_container_width=True
    )

    # ---------- PDF ----------
    pdf_name = selected_marketing if is_admin and selected_marketing != "ALL" else marketing
    pdf = generate_pdf(pdf_name, monthly_report)

    st.download_button(
        f"📄 Download {pdf_name} – Sales Report",
        pdf,
        file_name=f"{pdf_name}_Sales_Report.pdf",
        mime="application/pdf"
    )

# ============ MAIN ============
if "user" not in st.session_state:
    login()
else:
    dashboard()
