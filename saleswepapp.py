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

# ================= CACHE =================
@st.cache_data(show_spinner="Loading Excel...")
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

    elements = [
        Paragraph(
            f"<b>SALES PERFORMANCE REPORT (2025–2026)</b><br/><br/>"
            f"<b>Marketing Person:</b> {marketing_name}",
            styles["Title"]
        )
    ]

    total_target = df["Target"].sum()
    total_sales = df["sales"].sum()

    summary = Table([
        ["TARGET", "ACHIEVED", "%"],
        [f"₹ {total_target:,.0f}", f"₹ {total_sales:,.0f}",
         f"{(total_sales/total_target*100):.1f} %" if total_target > 0 else "0 %"]
    ])

    summary.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B3C91")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 1, colors.black)
    ]))

    elements.append(summary)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ================= LOGIN =================
def login():
    st.title("🔐 Marketing Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u in USERS and USERS[u]["password"] == p:
            st.session_state["user"] = u
            st.session_state["marketing"] = USERS[u]["marketing"]
            st.rerun()
        else:
            st.error("Invalid login")

# ================= DASHBOARD =================
def dashboard():
    marketing = st.session_state["marketing"]
    is_admin = marketing == "ALL"

    st.sidebar.success(f"Logged in as: {marketing}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()

    if is_admin:
        f = st.file_uploader("📤 Upload Excel", type="xlsx")
        if f:
            st.session_state["file_bytes"] = f.getvalue()
            st.session_state["updated"] = datetime.now()

    if "file_bytes" not in st.session_state:
        st.warning("⚠️ Admin has not uploaded file")
        st.stop()

    sales_df, target_raw, make_target_df, new_customer_df = load_excel_cached(
        st.session_state["file_bytes"]
    )

    # ================= EXISTING LOGIC (UNCHANGED) =================
    sales_df["MARK"] = sales_df["MARK"].astype(str).str.upper().str.strip()
    sales_df["make"] = sales_df["make"].astype(str).str.upper().str.strip()
    sales_df = sales_df[sales_df["HELPER"].isin(["NOFILL", "GREEN"])]

    if not is_admin:
        sales_df = sales_df[sales_df["MARK"] == marketing.upper()]

    # ================= BRAND WISE (NEW FEATURE – ADDED ONLY) =================
    make_target_df["Make"] = make_target_df["Make"].astype(str).str.upper().str.strip()

    start_date = pd.Timestamp("2025-04-01")
    end_date = pd.Timestamp("2026-01-31")
    month_count = 10

    brand_rows = []

    for _, r in make_target_df.iterrows():
        make_name = r["Make"]
        monthly_target = r["Target"]

        sales_val = sales_df[
            (sales_df["Date"] >= start_date) &
            (sales_df["Date"] <= end_date) &
            (sales_df["make"].str.contains(make_name, na=False))
        ]["sales"].sum()

        total_target = monthly_target * month_count
        ach_pct = (sales_val / total_target * 100) if total_target > 0 else 0

        brand_rows.append({
            "Brand": make_name,
            "Sales": round(sales_val, 2),
            "Target": round(total_target, 2),
            "Achievement_%": round(ach_pct, 1)
        })

    brand_report = pd.DataFrame(brand_rows)

    # ================= NEW CUSTOMER (NEW FEATURE) =================
    new_customer_df["Date"] = pd.to_datetime(new_customer_df["Date"])
    new_customer_df = new_customer_df[
        (new_customer_df["Date"] >= start_date) &
        (new_customer_df["Date"] <= end_date)
    ]

    new_customer_count = new_customer_df["CUSTOMER NAME"].nunique()
    new_customer_sales = new_customer_df["sales"].sum()

    # ================= UI =================
    st.title("📊 SALES PERFORMANCE DASHBOARD")

    st.subheader("🆕 New Customers (Apr 2025 – Jan 2026)")
    st.metric("Number of New Customers", new_customer_count)
    st.metric("New Customer Sales", f"₹ {new_customer_sales:,.0f}")

    st.divider()

    st.subheader("🏷️ Brand Wise Sales (Apr 2025 – Jan 2026)")
    st.dataframe(brand_report, use_container_width=True)

    pdf = generate_pdf(marketing, sales_df)
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
