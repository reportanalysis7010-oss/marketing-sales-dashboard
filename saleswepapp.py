import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, black, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ================= FONT LOADING (GITHUB SAFE) =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR, "fonts")

pdfmetrics.registerFont(TTFont("DejaVu", os.path.join(FONT_DIR, "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf")))

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

st.set_page_config(page_title="Marketing Sales Dashboard", layout="wide")

# ================= CACHE =================
@st.cache_data(show_spinner="Loading Excel data...")
def load_excel_cached(file_bytes):
    sales_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SALES_SHEET)
    target_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=TARGET_SHEET)
    return sales_df, target_df

# ================= PDF GENERATION =================
def generate_pdf(marketing_name, df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    BLUE   = HexColor("#0B3C91")
    GREEN  = HexColor("#2ECC71")
    RED    = HexColor("#E74C3C")
    YELLOW = HexColor("#F4D03F")
    GRAY   = HexColor("#F2F2F2")

    # ---------- HEADER ----------
    c.setFont("DejaVu-Bold", 20)
    c.drawCentredString(width / 2, height - 40, "SALES PERFORMANCE REPORT (2025–2026)")

    c.setFont("DejaVu-Bold", 15)
    c.drawCentredString(width / 2, height - 75, f"Marketing Person: {marketing_name.upper()}")

    # ---------- KPI CALC ----------
    total_target = df["Target"].sum()
    total_sales = df["sales"].sum()
    not_achieved = total_target - total_sales
    achievement_pct = (total_sales / total_target * 100) if total_target else 0

    # ---------- KPI BOXES ----------
    box_y = height - 140
    box_w = (width - 80) / 4
    box_h = 45
    x_start = 40

    kpis = [
        ("TARGET", total_target, GREEN),
        ("ACHIEVED", total_sales, GREEN),
        ("NOT ACHIEVED", not_achieved, RED),
        ("%", f"{achievement_pct:.1f} %", YELLOW),
    ]

    for i, (label, value, color) in enumerate(kpis):
        x = x_start + i * box_w

        c.setFillColor(BLUE)
        c.rect(x, box_y + box_h, box_w, 20, fill=1)

        c.setFillColor(color)
        c.rect(x, box_y, box_w, box_h, fill=1)

        c.setFillColor(white)
        c.setFont("DejaVu-Bold", 11)
        c.drawCentredString(x + box_w / 2, box_y + box_h + 5, label)

        c.setFillColor(black)
        c.setFont("DejaVu-Bold", 13)
        c.drawCentredString(
            x + box_w / 2,
            box_y + 15,
            f"₹ {value:,.0f}" if isinstance(value, (int, float)) else value
        )

    # ---------- TABLE ----------
    table_y = box_y - 40
    col_x = [40, 160, 310, 460]

    c.setFillColor(GRAY)
    c.rect(40, table_y, width - 80, 22, fill=1)

    c.setFillColor(black)
    c.setFont("DejaVu-Bold", 11)
    c.drawString(col_x[0], table_y + 6, "Month")
    c.drawString(col_x[1], table_y + 6, "Target")
    c.drawString(col_x[2], table_y + 6, "Sales")
    c.drawString(col_x[3], table_y + 6, "Achievement %")

    c.setFont("DejaVu", 11)
    y = table_y - 18

    for _, r in df.iterrows():
        c.drawString(col_x[0], y, r["Month_Text"])
        c.drawRightString(col_x[1] + 90, y, f"₹ {r['Target']:,.0f}")
        c.drawRightString(col_x[2] + 90, y, f"₹ {r['sales']:,.0f}")
        c.drawRightString(col_x[3] + 70, y, f"{r['Achievement_%']} %")

        y -= 18
        if y < 60:
            c.showPage()
            c.setFont("DejaVu", 11)
            y = height - 60

    c.save()
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

    if is_admin:
        uploaded = st.file_uploader("📤 Upload Marketing Excel File", type="xlsx")
        if uploaded:
            st.session_state["file_bytes"] = uploaded.getvalue()
            st.session_state["last_updated"] = datetime.now()

    if "file_bytes" not in st.session_state:
        st.warning("⚠️ Admin has not uploaded the Excel file")
        st.stop()

    sales_df, target_raw = load_excel_cached(st.session_state["file_bytes"])

    # ---------- TARGET ----------
    target_df = target_raw.melt(
        id_vars=["Marketing Person"],
        var_name="Month",
        value_name="Target"
    ).rename(columns={"Marketing Person": "MARK"})

    target_df["Target"] = (
        target_df["Target"].astype(str)
        .str.replace("₹", "", regex=False)
        .str.replace(",", "", regex=False)
        .astype(float)
    )

    target_df["Month_No"] = target_df["Month"].map(MONTH_MAP)
    target_df["Year"] = target_df["Month_No"].apply(lambda x: 2025 if x >= 4 else 2026)
    target_df["YearMonth"] = target_df["Year"] * 100 + target_df["Month_No"]

    # ---------- SALES ----------
    sales_df["MARK"] = sales_df["MARK"].str.upper().str.strip()
    target_df["MARK"] = target_df["MARK"].str.upper().str.strip()

    sales_df = sales_df[sales_df["HELPER"].isin(["NOFILL", "GREEN"])]

    if not is_admin:
        sales_df = sales_df[sales_df["MARK"] == marketing.upper()]
        target_df = target_df[target_df["MARK"] == marketing.upper()]

    monthly_sales = sales_df.groupby(
        ["MARK", "YearMonth", "Month_Text"], as_index=False
    )["sales"].sum()

    report = pd.merge(
        monthly_sales,
        target_df[["MARK", "YearMonth", "Target"]],
        on=["MARK", "YearMonth"],
        how="left"
    )

    report["Achievement_%"] = (report["sales"] / report["Target"] * 100).round(1)

    # ---------- UI ----------
    st.title("📊 SALES PERFORMANCE DASHBOARD")

    st.dataframe(report, use_container_width=True)

    pdf = generate_pdf(marketing, report)
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
