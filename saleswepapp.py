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
@st.cache_data(show_spinner="Loading Excel data...")
def load_excel_cached(file_bytes):
    sales_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SALES_SHEET)
    target_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=TARGET_SHEET)

    try:
        make_target_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=MAKE_TARGET_SHEET)
    except:
        make_target_df = pd.DataFrame(columns=["Make", "Target"])

    try:
        new_customer_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=NEW_CUSTOMER_SHEET)
    except:
        new_customer_df = pd.DataFrame()

    return sales_df, target_raw, make_target_df, new_customer_df

# ================= PDF =================
def generate_pdf(marketing_name, df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    styles["Title"].fontName = "DejaVu-Bold"
    styles["Normal"].fontName = "DejaVu"

    elements = []

    elements.append(Paragraph(
        f"<b>SALES PERFORMANCE REPORT (2025–2026)</b><br/><br/>"
        f"<b>Marketing Person:</b> {marketing_name}",
        styles["Title"]
    ))

    total_target = df["Target"].sum()
    total_sales = df["sales"].sum()
    not_achieved = total_target - total_sales
    pct = (total_sales / total_target * 100) if total_target else 0

    summary = Table(
        [
            ["TARGET", "ACHIEVED", "NOT ACHIEVED", "%"],
            [
                f"₹ {total_target:,.0f}",
                f"₹ {total_sales:,.0f}",
                f"₹ {not_achieved:,.0f}",
                f"{pct:.1f} %"
            ]
        ],
        colWidths=[120, 120, 120, 60]
    )

    summary.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B3C91")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("BACKGROUND", (0,1), (1,1), colors.HexColor("#2ECC71")),
        ("BACKGROUND", (2,1), (2,1), colors.HexColor("#E74C3C")),
        ("BACKGROUND", (3,1), (3,1), colors.khaki),
        ("GRID", (0,0), (-1,-1), 1, colors.black),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
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

    detail = Table(table_data, repeatRows=1)
    detail.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 1, colors.black),
        ("ALIGN", (1,1), (-1,-1), "RIGHT"),
    ]))

    elements.append(detail)
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
            st.session_state["last_updated"] = datetime.now()

    if "file_bytes" not in st.session_state:
        st.warning("⚠️ Admin has not uploaded the Excel file yet")
        st.stop()

    sales_df, target_raw, make_target_df, new_customer_df = load_excel_cached(
        st.session_state["file_bytes"]
    )

    # ================= FIX SALES COLUMN =================
    sales_df["sales"] = sales_df["Value"].fillna(0)

    # ================= MONTHLY TARGET =================
    target_df = target_raw.melt(
        id_vars=["Marketing Person"],
        var_name="Month",
        value_name="Target"
    ).rename(columns={"Marketing Person": "MARK"})

    target_df["Target"] = (
        target_df["Target"].astype(str)
        .str.replace("₹", "").str.replace(",", "").str.strip()
    ).astype(float)

    target_df["Month_No"] = target_df["Month"].map(MONTH_MAP)
    target_df["Year"] = target_df["Month_No"].apply(lambda x: 2025 if x >= 4 else 2026)
    target_df["YearMonth"] = target_df["Year"] * 100 + target_df["Month_No"]

    # ================= CLEAN =================
    sales_df["MARK"] = sales_df["MARK"].str.upper().str.strip()
    sales_df["make"] = sales_df["make"].astype(str).str.upper()
    target_df["MARK"] = target_df["MARK"].str.upper().str.strip()

    sales_df = sales_df[sales_df["HELPER"].isin(["NOFILL", "GREEN"])]

    # ================= MARKETING FILTER =================
    selected_marketing = "ALL"
    if is_admin:
        selected_marketing = st.selectbox(
            "📌 Select Marketing Person",
            ["ALL"] + sorted(target_df["MARK"].unique())
        )
        if selected_marketing != "ALL":
            sales_df = sales_df[sales_df["MARK"] == selected_marketing]
            target_df = target_df[target_df["MARK"] == selected_marketing]
    else:
        m = marketing.upper()
        sales_df = sales_df[sales_df["MARK"] == m]
        target_df = target_df[target_df["MARK"] == m]

    # ================= MONTHLY REPORT =================
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

    monthly_report = monthly_report.sort_values(["MARK", "YearMonth"])

    # ================= SALES PERFORMANCE REPORT =================
    st.subheader("📊 Sales Performance Report")

    col1, col2, col3 = st.columns(3)

    col1.metric("Total Target", f"₹ {monthly_report['Target'].sum():,.0f}")
    col2.metric("Total Sales", f"₹ {monthly_report['sales'].sum():,.0f}")

    achievement = (
        monthly_report['sales'].sum() / monthly_report['Target'].sum() * 100
        if monthly_report['Target'].sum() > 0 else 0
    )
    col3.metric("Achievement %", f"{achievement:.1f} %")

    st.subheader("📊 Month-wise Target vs Sales")

    chart_df = (
        monthly_report
        .groupby("Month_Text", as_index=False)[["Target", "sales"]]
        .sum()
        .set_index("Month_Text")
    )
    st.bar_chart(chart_df)

    st.subheader("📋 Month-wise Sales Performance")
    st.dataframe(
        monthly_report[["MARK", "Month_Text", "Target", "sales", "Achievement_%"]]
        .rename(columns={
            "MARK": "Marketing Person",
            "sales": "Sales Achieved",
            "Achievement_%": "Achievement %"
        }),
        use_container_width=True
    )

    # ================= NEW CUSTOMER REPORT =================
    st.subheader("🆕 New Customer Report")

    new_customer_df["CUSTOMER NAME"] = new_customer_df["CUSTOMER NAME"].astype(str).str.upper().str.strip()
    sales_df["CUSTOMER NAME"] = sales_df["CUSTOMER NAME"].astype(str).str.upper().str.strip()

    new_customer_sales_df = sales_df[sales_df["CUSTOMER NAME"].isin(new_customer_df["CUSTOMER NAME"])]

    new_customer_count = new_customer_sales_df["CUSTOMER NAME"].nunique()
    new_customer_sales = new_customer_sales_df["sales"].sum()

    col_nc1, col_nc2 = st.columns(2)
    col_nc1.metric("New Customers", new_customer_count)
    col_nc2.metric("New Customer Sales", f"₹ {new_customer_sales:,.0f}")

    # ================= BRAND WISE =================
    st.subheader("🏷️ Brand Wise Sales")

    make_target_df["Make"] = make_target_df["Make"].astype(str).str.upper()

    brand_rows = []
    months_count = sales_df["YearMonth"].nunique()

    for _, r in make_target_df.iterrows():
        mk = r["Make"]
        m_target = r["Target"] * months_count

        mk_sales = sales_df[sales_df["make"].str.contains(mk, na=False)]["sales"].sum()
        pct = (mk_sales / m_target * 100) if m_target else 0

        brand_rows.append({
            "Brand": mk,
            "Sales": round(mk_sales),
            "Target": round(m_target),
            "Achievement_%": round(pct, 1)
        })

    brand_df = pd.DataFrame(brand_rows)
    st.dataframe(brand_df, use_container_width=True)

    # ================= PROJECTION (NEW FEATURE) =================
    st.subheader("📈 Target Projection (Dynamic Remaining Months)")

    max_month_no = sales_df["Month No"].max()

    fiscal_index = max_month_no - 3
    if fiscal_index < 1:
        fiscal_index += 12

    remaining_months = 12 - fiscal_index

    yearly_target = target_df["Target"].sum()
    achieved_sales = sales_df["sales"].sum()
    remaining_sales = yearly_target - achieved_sales
    monthly_need = (remaining_sales / remaining_months) if remaining_months > 0 else 0

    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
    col_p1.metric("Yearly Target", f"₹ {yearly_target:,.0f}")
    col_p2.metric("Achieved Sales", f"₹ {achieved_sales:,.0f}")
    col_p3.metric("Needed to Achieve", f"₹ {remaining_sales:,.0f}")
    col_p4.metric("Required Per Month", f"₹ {monthly_need:,.0f}")

    # ================= PDF =================
    pdf_name = selected_marketing if selected_marketing != "ALL" else marketing
    pdf = generate_pdf(pdf_name, monthly_report)

    st.download_button(
        f"📄 Download {pdf_name} Sales Report",
        pdf,
        file_name=f"{pdf_name}_Sales_Report.pdf",
        mime="application/pdf"
    )

# ================= MAIN =================
if "user" not in st.session_state:
    login()
else:
    dashboard()
