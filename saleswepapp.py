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



# ================= PDF GENERATOR =================
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
    total_sales = df["Value"].sum()
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
        ("GRID", (0,0), (-1,-1), 1, colors.black),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))

    elements.append(summary)
    elements.append(Paragraph("<br/>", styles["Normal"]))

    # Detail Rows
    table_data = [["Month", "Target", "Sales", "Achievement %"]]
    for _, r in df.iterrows():
        table_data.append([
            r["Month_Text"],
            f"₹ {r['Target']:,.0f}",
            f"₹ {r['Value']:,.0f}",
            f"{r['Achievement_%']} %"
        ])

    detail = Table(table_data, repeatRows=1)
    detail.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 1, colors.black),
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



# ================= MAIN DASHBOARD =================
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

    if "file_bytes" not in st.session_state:
        st.warning("⚠️ Admin has not uploaded Excel")
        st.stop()

    sales_df, target_raw, make_target_df, new_customer_df = load_excel_cached(
        st.session_state["file_bytes"]
    )


    # ================= FIX Month_Text =================
    month_cols = ["Month_Text", "MONTH_TEXT", "Month Text", "monthtext"]
    found = False
    for col in month_cols:
        if col in sales_df.columns:
            sales_df["Month_Text"] = sales_df[col].astype(str)
            found = True
            break

    if not found:
        st.error("❌ Month_Text column missing")
        st.stop()


    # ================= CLEAN SALES DF =================
    sales_df["MARK"] = sales_df["MARK"].str.upper().str.strip()
    sales_df["make"] = sales_df["make"].astype(str).str.upper()

    sales_df = sales_df[sales_df["HELPER"].isin(["NOFILL", "GREEN"])]

    # Convert VALUE to numeric
    sales_df["Value"] = pd.to_numeric(sales_df["Value"], errors="coerce").fillna(0)


    # ================= TARGET DF =================
    target_df = target_raw.melt(
        id_vars=["Marketing Person"],
        var_name="Month",
        value_name="Target"
    ).rename(columns={"Marketing Person": "MARK"})

    target_df["Target"] = (
        target_df["Target"].astype(str).str.replace("₹","").str.replace(",","")
    ).astype(float)

    target_df["Month_No"] = target_df["Month"].map(MONTH_MAP)
    target_df["Year"] = target_df["Month_No"].apply(lambda x: 2025 if x >= 4 else 2026)
    target_df["YearMonth"] = target_df["Year"]*100 + target_df["Month_No"]
    target_df["MARK"] = target_df["MARK"].str.upper().str.strip()


    # ================= MARKETING FILTER =================
    selected_marketing = "ALL"
    if is_admin:
        selected_marketing = st.selectbox(
            "Select Marketing Person",
            ["ALL"] + sorted(target_df["MARK"].unique())
        )
        if selected_marketing != "ALL":
            sales_df = sales_df[sales_df["MARK"] == selected_marketing]
            target_df = target_df[target_df["MARK"] == selected_marketing]
    else:
        marketing_upper = marketing.upper()
        sales_df = sales_df[sales_df["MARK"] == marketing_upper]
        target_df = target_df[target_df["MARK"] == marketing_upper]


    # ================= MONTHLY REPORT =================
    if "YearMonth" not in sales_df.columns:
        st.error("❌ YearMonth missing in MAIN_COPY sheet")
        st.stop()

    monthly_sales = sales_df.groupby(
        ["MARK", "YearMonth", "Month_Text"], as_index=False
    )["Value"].sum()

    monthly_report = pd.merge(
        monthly_sales,
        target_df[["MARK","YearMonth","Target"]],
        on=["MARK","YearMonth"],
        how="left"
    )

    monthly_report["Target"] = monthly_report["Target"].fillna(0)
    monthly_report["Achievement_%"] = (
        monthly_report["Value"]/monthly_report["Target"]*100
    ).round(1)


    # ================= SALES PERFORMANCE REPORT =================
    st.subheader("📊 Sales Performance Report")

    col1, col2, col3 = st.columns(3)
    total_target = monthly_report["Target"].sum()
    total_sales = monthly_report["Value"].sum()

    col1.metric("Total Target", f"₹ {total_target:,.0f}")
    col2.metric("Total Sales", f"₹ {total_sales:,.0f}")
    col3.metric("Achievement %", f"{(total_sales/total_target*100):.1f} %" if total_target else "0 %")


    # ================= SALES PROJECTION =================
    st.subheader("📈 Sales Projection (To Achieve Full Target)")

    # Completed months based on VALUE, not sales column
    completed_months = sales_df[sales_df["Value"] > 0]["YearMonth"].nunique()

    remaining_target = max(total_target - total_sales, 0)
    remaining_months = max(12 - completed_months, 0)

    required_monthly_sales = (remaining_target / remaining_months) if remaining_months > 0 else 0

    colp1, colp2, colp3, colp4 = st.columns(4)
    colp1.metric("Completed Months", completed_months)
    colp2.metric("Remaining Target", f"₹ {remaining_target:,.0f}")
    colp3.metric("Months Left", remaining_months)
    colp4.metric("Required Monthly Sales", f"₹ {required_monthly_sales:,.0f}")


    # ================= CHARTS =================
    st.subheader("📊 Month-wise Target vs Sales")

    chart_df = monthly_report.groupby("Month_Text", as_index=False)[["Target","Value"]].sum()
    chart_df = chart_df.set_index("Month_Text")

    st.bar_chart(chart_df)


    # ================= MONTHLY TABLE =================
    st.subheader("📋 Month-wise Sales Performance")
    st.dataframe(monthly_report.rename(columns={"Value":"Sales Value"}), use_container_width=True)


    # ================= NEW CUSTOMER REPORT =================
    st.subheader("🆕 New Customer Report")

    sales_df["CUSTOMER NAME"] = sales_df["CUSTOMER NAME"].astype(str).str.upper().str.strip()
    new_customer_df["CUSTOMER NAME"] = new_customer_df["CUSTOMER NAME"].astype(str).str.upper().str.strip()

    new_customer_sales_df = sales_df[
        sales_df["CUSTOMER NAME"].isin(new_customer_df["CUSTOMER NAME"])
    ]

    nc_count = new_customer_sales_df["CUSTOMER NAME"].nunique()
    nc_sales = new_customer_sales_df["Value"].sum()

    colnc1, colnc2 = st.columns(2)
    colnc1.metric("New Customers", nc_count)
    colnc2.metric("New Customer Sales", f"₹ {nc_sales:,.0f}")


    # ================= BRAND WISE SALES =================
    st.subheader("🏷️ Brand Wise Sales")

    make_target_df["Make"] = make_target_df["Make"].astype(str).str.upper()

    brand_rows = []
    months_count = sales_df["YearMonth"].nunique()

    for _, r in make_target_df.iterrows():
        mk = r["Make"]
        m_target = r["Target"] * months_count

        mk_sales = sales_df[sales_df["make"].str.contains(mk, na=False)]["Value"].sum()
        pct = (mk_sales / m_target * 100) if m_target > 0 else 0

        brand_rows.append({
            "Brand": mk,
            "Sales": mk_sales,
            "Target": m_target,
            "Achievement_%": round(pct, 1)
        })

    st.dataframe(pd.DataFrame(brand_rows), use_container_width=True)


    # ================= PDF DOWNLOAD =================
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
