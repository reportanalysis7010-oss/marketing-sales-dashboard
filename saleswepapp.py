import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ================= FONT =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR, "fonts")

pdfmetrics.registerFont(TTFont("DejaVu", os.path.join(FONT_DIR, "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf")))

# ================= CONFIG =================
SALES_SHEET = "MAIN_COPY"
TARGET_SHEET = "MARKETING TARGET"
MAKE_TARGET_SHEET = "MAKE TARGET"
NEW_CUSTOMER_SHEET = "Merge1"

USERS = {
    "admin": {"password": "admin@123", "marketing": "ALL"},
    "ashok": {"password": "ashok@123", "marketing": "Ashok Marketing"},
    "suresh": {"password": "suresh@123", "marketing": "Suresh - Marketing"},
    "ho": {"password": "ho@123", "marketing": "H O - Marketing"},
}

MONTH_MAP = {
    "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7,
    "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11,
    "DEC": 12, "JAN": 1, "FEB": 2, "MAR": 3
}

st.set_page_config(page_title="Marketing Sales Dashboard", layout="wide")

# ================= LOAD EXCEL =================
@st.cache_data(show_spinner="Loading Excel data...")
def load_excel_cached(file_bytes):
    sales_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SALES_SHEET)
    target_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=TARGET_SHEET)
    make_target_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=MAKE_TARGET_SHEET)
    new_customer_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=NEW_CUSTOMER_SHEET)
    return sales_df, target_raw, make_target_df, new_customer_df

# ================= PDF =================
def generate_pdf(marketing_name, df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    styles["Title"].fontName = "DejaVu-Bold"
    styles["Normal"].fontName = "DejaVu"

    elements = []

    elements.append(
        Paragraph(
            f"<b>SALES PERFORMANCE REPORT (2025–2026)</b><br/><br/>"
            f"<b>Marketing Person:</b> {marketing_name}",
            styles["Title"]
        )
    )

    total_target = df["Target"].sum()
    total_sales = df["sales"].sum()
    not_achieved = total_target - total_sales
    achievement_pct = (total_sales / total_target * 100) if total_target else 0

    summary = Table([
        ["TARGET", "ACHIEVED", "NOT ACHIEVED", "%"],
        [
            f"₹ {total_target:,.0f}",
            f"₹ {total_sales:,.0f}",
            f"₹ {not_achieved:,.0f}",
            f"{achievement_pct:.1f} %"
        ]
    ])

    summary.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B3C91")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (1, 1), colors.HexColor("#2ECC71")),
        ("BACKGROUND", (2, 1), (2, 1), colors.HexColor("#E74C3C")),
        ("BACKGROUND", (3, 1), (3, 1), colors.HexColor("#F4D03F")),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER")
    ]))

    elements.append(summary)
    elements.append(Paragraph("<br/>", styles["Normal"]))

    table_data = [["Month", "Target", "Sales", "Achievement %"]]
    for _, r in df.iterrows():
        table_data.append([
            r["Month_Text"],
            f"₹ {r['Target']:,.0f}",
            f"₹ {r['sales']:,.0f}",
            f"{r['Achievement_%']} %"
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 1, colors.black)
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ================= LOGIN =================
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

# ================= DASHBOARD =================
def dashboard():
    marketing = st.session_state["marketing"]
    is_admin = marketing == "ALL"

    st.sidebar.success(f"Logged in as: {marketing}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()

    uploaded_file = st.file_uploader("📤 Upload Marketing Excel File", type=["xlsx"])
    if uploaded_file:
        st.session_state["file_bytes"] = uploaded_file.getvalue()
        st.session_state["last_updated"] = datetime.now()

    if "file_bytes" not in st.session_state:
        st.warning("Upload Excel file to continue")
        st.stop()

    sales_df, target_raw, make_target_df, new_customer_df = load_excel_cached(
        st.session_state["file_bytes"]
    )

    # ================= EXISTING LOGIC (UNCHANGED) =================
    target_df = target_raw.melt(
        id_vars=["Marketing Person"],
        var_name="Month",
        value_name="Target"
    ).rename(columns={"Marketing Person": "MARK"})

    target_df["Target"] = (
        target_df["Target"].astype(str)
        .str.replace("₹", "").str.replace(",", "").astype(float)
    )

    target_df["Month_No"] = target_df["Month"].map(MONTH_MAP)
    target_df["Year"] = target_df["Month_No"].apply(lambda x: 2025 if x >= 4 else 2026)
    target_df["YearMonth"] = target_df["Year"] * 100 + target_df["Month_No"]

    sales_df = sales_df[sales_df["HELPER"].isin(["NOFILL", "GREEN"])]
    sales_df["MARK"] = sales_df["MARK"].str.upper().str.strip()
    target_df["MARK"] = target_df["MARK"].str.upper().str.strip()

    if not is_admin:
        m = marketing.upper()
        sales_df = sales_df[sales_df["MARK"] == m]
        target_df = target_df[target_df["MARK"] == m]

    monthly_sales = sales_df.groupby(
        ["MARK", "YearMonth", "Month_Text"], as_index=False
    )["sales"].sum()

    monthly_report = pd.merge(
        monthly_sales,
        target_df[["MARK", "YearMonth", "Target"]],
        on=["MARK", "YearMonth"],
        how="left"
    )

    monthly_report["Target"] = monthly_report["Target"].fillna(0)
    monthly_report["Achievement_%"] = (
        monthly_report["sales"] / monthly_report["Target"] * 100
    ).round(1)

    # ================= BRAND-WISE LOGIC (USING `make`) =================
    sales_df["make"] = sales_df["make"].astype(str).str.strip().str.upper()
    make_target_df["Make"] = make_target_df["Make"].astype(str).str.strip().str.upper()

    start_date = pd.Timestamp("2025-04-01")
    end_date = pd.Timestamp("2026-01-31")

    month_count = (
        (end_date.year - start_date.year) * 12 +
        (end_date.month - start_date.month) + 1
    )

    brand_sales = sales_df.groupby("make", as_index=False)["sales"].sum()
    make_target_df["Total_Target"] = make_target_df["Target"] * month_count

    brand_report = pd.merge(
        brand_sales,
        make_target_df,
        left_on="make",
        right_on="Make",
        how="left"
    )

    brand_report["Achievement_%"] = (
        brand_report["sales"] / brand_report["Total_Target"] * 100
    ).round(1)

    # ================= NEW CUSTOMER LOGIC =================
    new_customer_df["CUSTOMER NAME"] = new_customer_df["CUSTOMER NAME"].astype(str).str.strip()
    new_customer_count = new_customer_df["CUSTOMER NAME"].nunique()

    new_customer_sales = sales_df[
        sales_df["CUSTOMER NAME"].isin(new_customer_df["CUSTOMER NAME"])
    ]["sales"].sum()

    # ================= UI =================
    st.title("📊 SALES PERFORMANCE DASHBOARD")

    st.subheader("📋 Month-wise Sales Performance")
    st.dataframe(monthly_report, use_container_width=True)

    st.subheader("🆕 New Customer Summary")
    st.metric("Number of New Customers", new_customer_count)
    st.metric("New Customer Sales Value", f"₹ {new_customer_sales:,.0f}")

    st.subheader("🏷️ Brand Wise Sales (Apr 2025 – Jan 2026)")
    st.dataframe(
        brand_report[["make", "sales", "Total_Target", "Achievement_%"]]
        .rename(columns={
            "make": "Brand",
            "sales": "Sales",
            "Total_Target": "Target"
        }),
        use_container_width=True
    )

    pdf = generate_pdf(marketing, monthly_report)
    st.download_button(
        "📄 Download Sales Report PDF",
        pdf,
        file_name=f"{marketing}_Sales_Report.pdf",
        mime="application/pdf"
    )

# ================= MAIN =================
if "user" not in st.session_state:
    login()
else:
    dashboard()
