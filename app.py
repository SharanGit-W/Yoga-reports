"""
Yoga Center Attendance Generator - Web Dashboard
"""
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple, List

import streamlit as st
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Required for running Matplotlib on a web server
import matplotlib.pyplot as plt

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak

# =============================================================================
# Configuration & Data Models
# =============================================================================
PAGE_SIZE = A4
LEFT_MARGIN, RIGHT_MARGIN = 14 * mm, 14 * mm
TOP_MARGIN, BOTTOM_MARGIN = 16 * mm, 16 * mm

BRAND_RED = colors.HexColor("#8A1E1E")
BRAND_DARK = colors.HexColor("#222222")
BRAND_GREY = colors.HexColor("#666666")
BRAND_LIGHT = colors.HexColor("#F5F5F5")
CHART_RED = "#8A1E1E"
CHART_DARK = "#222222"
DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

@dataclass
class ReportMeta:
    center_name: str
    report_month: str
    report_period: str
    month_key: str

# =============================================================================
# Core Logic (Same proven logic from v2.4)
# =============================================================================
def safe_str(x) -> str:
    return "" if pd.isna(x) else str(x).strip()

def normalize_col(col) -> str:
    if isinstance(col, pd.Timestamp): return col.strftime("%d-%b-%Y")
    if hasattr(col, "year") and hasattr(col, "month") and hasattr(col, "day"):
        try: return pd.Timestamp(col).strftime("%d-%b-%Y")
        except: return str(col).strip()
    return str(col).strip()

def clean_center_name(raw_center: str) -> str:
    txt = safe_str(raw_center)
    if not txt: return "UNKNOWN CENTER"
    if "-" in txt: txt = txt.split("-", 1)[-1].strip()
    txt = re.sub(r"\s*\(.*?\)\s*", "", txt).strip()
    lower = txt.lower()
    replacements = {"vijayanagar(blore)": "Vijayanagara", "vijayanagar": "Vijayanagara", "sadashivanagar": "Sadashivanagara"}
    for k, v in replacements.items():
        if lower.replace(" ", "") == k.replace(" ", ""): return v.upper()
    return txt.upper()

def find_header_row(path: str) -> int:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in range(1, min(20, ws.max_row) + 1):
        values = [safe_str(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 12) + 1)]
        if "SlNo" in values and "StudentId" in values: return r - 1
    raise ValueError("Could not find header row containing 'SlNo' and 'StudentId'.")

def attendance_cell_to_bool(series: pd.Series) -> pd.Series:
    if series.dtype == "O":
        s = series.fillna("").astype(str).str.strip().str.lower()
        return s.isin({"present", "p", "x", "1", "yes", "y", "true"})
    return (~series.isna()) & (series.astype(str).str.strip().str.lower() != "0")

def extract_metadata(raw: pd.DataFrame) -> str:
    for _, row in raw.head(5).iterrows():
        vals = [safe_str(v) for v in row.tolist()[:6]]
        joined = " | ".join(v for v in vals if v)
        if any(k in joined.lower() for k in ["vijayanagar", "sadashivanagar", "jayanagar", "rysri"]):
            for v in vals:
                if v and not re.search(r"^\d{1,2}[-/][A-Za-z]{3}[-/]\d{2,4}$", v):
                    if any(k in v.lower() for k in ["vijayanagar", "sadashivanagar", "jayanagar", "rysri"]): return v
    return ""

def read_and_clean(input_file: str):
    header_row = find_header_row(input_file)
    raw = pd.read_excel(input_file, header=header_row)
    raw.columns = [normalize_col(c) for c in raw.columns]
    raw = raw.dropna(how="all").reset_index(drop=True)

    status_idx = raw.columns.get_loc("Status")
    date_cols, date_index = [], []
    for c in raw.columns[status_idx + 1:]:
        if safe_str(c).strip().lower() in {"present", "absent"}: continue
        dt = pd.to_datetime(c, format="%d-%b-%Y", errors="coerce")
        if pd.notna(dt):
            date_cols.append(c)
            date_index.append(dt)

    sort_idx = np.argsort(pd.to_datetime(date_index))
    date_cols = [date_cols[i] for i in sort_idx]
    date_index = pd.DatetimeIndex([date_index[i] for i in sort_idx])

    df = raw.copy()
    for col in ["Batch", "Center", "Status", "StudentName", "StudentId"]:
        if col in df.columns: df[col] = df[col].map(safe_str)

    presence = pd.DataFrame({c: attendance_cell_to_bool(df[c]) for c in date_cols}, index=df.index)
    df["Present_Count_Calc"] = presence.sum(axis=1).astype(int)
    
    center_raw = df["Center"].mode().iloc[0] if "Center" in df.columns and df["Center"].str.strip().ne("").any() else extract_metadata(raw)
    date_start, date_end = date_index.min(), date_index.max()

    meta = ReportMeta(
        center_name=clean_center_name(center_raw),
        report_month=date_start.strftime("%B %Y"),
        report_period=f"{date_start.strftime('%d-%b-%Y')} to {date_end.strftime('%d-%b-%Y')}",
        month_key=date_start.strftime("%b_%Y")
    )
    return df, date_cols, date_index, meta, presence

def build_analytics(df, date_cols, date_index, presence):
    daily_totals = pd.DataFrame({"Date": date_index, "Attendance": presence.sum(axis=0).values.astype(int)})
    daily_totals["DayOfWeek"] = daily_totals["Date"].dt.day_name()
    daily_totals["ShortDate"] = daily_totals["Date"].dt.strftime("%d-%b")
    dow_totals = daily_totals.groupby("DayOfWeek", as_index=False)["Attendance"].sum().set_index("DayOfWeek").reindex(DOW_ORDER).fillna(0).reset_index()
    
    batch_analysis = df.groupby("Batch", as_index=False).agg(Members=("StudentId", "count"), Total_Present=("Present_Count_Calc", "sum")).sort_values("Total_Present", ascending=False)

    df["Segment"] = pd.cut(df["Present_Count_Calc"], bins=[-1, 0, 5, 11, 35], labels=["Inactive (0 visits)", "Occasional (1-5 visits)", "Regular (6-11 visits)", "Dedicated (12+ visits)"])
    segment_counts = df["Segment"].value_counts().reindex(["Inactive (0 visits)", "Occasional (1-5 visits)", "Regular (6-11 visits)", "Dedicated (12+ visits)"]).fillna(0).reset_index()
    segment_counts.columns = ["Member Segment", "Count"]

    top_attendees = df[["StudentName", "Batch", "Present_Count_Calc"]].sort_values(["Present_Count_Calc", "StudentName"], ascending=[False, True]).reset_index(drop=True)
    least_active = df[["StudentName", "Batch", "Present_Count_Calc"]].sort_values(["Present_Count_Calc", "StudentName"], ascending=[True, True]).reset_index(drop=True)

    total_att = int(daily_totals["Attendance"].sum())
    kpis = pd.DataFrame({
        "Metric": ["Registered Members", "Operating Days", "Total Center Visits", "Avg Visits per Member", "Busiest Day"],
        "Value": [len(df), len(date_cols), total_att, round(total_att/max(1, len(df)), 1), dow_totals.loc[dow_totals['Attendance'].idxmax(), 'DayOfWeek']]
    })
    
    insights = [
        f"Habit Building: On average, each member visited {kpis.iloc[3]['Value']} times this month.",
        f"Batch Popularity: The '{batch_analysis.iloc[0]['Batch'] if not batch_analysis.empty else 'N/A'}' batch had the highest footfall.",
        f"Peak Activity: {kpis.iloc[4]['Value']} saw the most activity across the month."
    ]

    return {"presence": presence, "daily_totals": daily_totals, "dow_totals": dow_totals, "batch_analysis": batch_analysis, "top_attendees": top_attendees, "least_active": least_active, "segment_counts": segment_counts, "kpis": kpis, "insights": insights}

# =============================================================================
# Plotting & Reporting Builders
# =============================================================================
def plot_charts(analytics, tmpdir):
    dow = analytics["dow_totals"]
    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.bar(dow["DayOfWeek"], dow["Attendance"], color=CHART_RED)
    ax.set_title("Attendance by Day of Week", fontweight="bold")
    for bar in bars: ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + dow["Attendance"].max()*0.02, str(int(bar.get_height())), ha="center", va="bottom")
    plt.xticks(rotation=20)
    plt.tight_layout()
    dow_path = os.path.join(tmpdir, "dow.png")
    plt.savefig(dow_path, dpi=150)
    plt.close(fig)

    daily = analytics["daily_totals"]
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(daily["ShortDate"], daily["Attendance"], marker="o", color=CHART_RED)
    ax.set_title("Daily Attendance Trend", fontweight="bold")
    for x, y in zip(daily["ShortDate"], daily["Attendance"]): ax.text(x, y + daily["Attendance"].max()*0.02, str(int(y)), ha="center")
    plt.xticks(rotation=45)
    plt.tight_layout()
    trend_path = os.path.join(tmpdir, "trend.png")
    plt.savefig(trend_path, dpi=150)
    plt.close(fig)
    return {"dow": dow_path, "trend": trend_path}

def make_pdf_table(df, headers, widths):
    tbl = Table([headers] + df.astype(str).values.tolist(), colWidths=widths, repeatRows=1)
    tbl.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), BRAND_RED), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.4, colors.gray), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, BRAND_LIGHT]), ("FONTSIZE", (0,0), (-1,-1), 8.5), ("ALIGN", (1,0), (-1,-1), "CENTER")]))
    return tbl

def build_pdf(out_pdf, meta, analytics, charts, logo=None):
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name="T", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, textColor=BRAND_DARK)
    sec_style = ParagraphStyle(name="S", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, textColor=BRAND_DARK, spaceBefore=6)
    
    doc = SimpleDocTemplate(out_pdf, pagesize=PAGE_SIZE, rightMargin=RIGHT_MARGIN, leftMargin=LEFT_MARGIN, topMargin=TOP_MARGIN, bottomMargin=BOTTOM_MARGIN)
    c = [Paragraph("ATTENDANCE ANALYSIS REPORT", title_style), Paragraph(meta.center_name, title_style), Spacer(1, 4*mm)]
    
    kpi_t = Table([["Metric", "Value"]] + analytics["kpis"].astype(str).values.tolist(), colWidths=[90*mm, 50*mm])
    kpi_t.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), BRAND_RED), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("GRID", (0,0), (-1,-1), 0.4, colors.gray)]))
    c.extend([kpi_t, Spacer(1, 4*mm), Paragraph("Key Insights", sec_style)])
    for ins in analytics["insights"]: c.append(Paragraph(f"• {ins}", styles["BodyText"]))

    c.extend([Paragraph("Attendance by Day of Week", sec_style), Image(charts["dow"], width=160*mm, height=75*mm), PageBreak()])
    c.extend([Paragraph("Daily Attendance Trend", sec_style), Image(charts["trend"], width=160*mm, height=75*mm), Spacer(1, 4*mm)])

    w1 = Table([[make_pdf_table(analytics["top_attendees"].head(8), ["Top Attendees", "Batch", "Visits"], [45*mm, 30*mm, 17*mm]), make_pdf_table(analytics["least_active"].head(8), ["Least Active", "Batch", "Visits"], [45*mm, 30*mm, 17*mm])]])
    w1.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    w2 = Table([[make_pdf_table(analytics["segment_counts"], ["Segment", "Count"], [65*mm, 27*mm]), make_pdf_table(analytics["batch_analysis"].head(5)[["Batch", "Members", "Total_Present"]], ["Batch", "Users", "Visits"], [55*mm, 18*mm, 19*mm])]])
    w2.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    c.extend([w1, Spacer(1, 4*mm), w2])
    
    def footer(canvas, doc):
        if logo and os.path.exists(logo):
            try: canvas.drawImage(logo, A4[0]-RIGHT_MARGIN-28*mm, A4[1]-TOP_MARGIN+2*mm, width=28*mm, height=18*mm, preserveAspectRatio=True, mask='auto')
            except: pass
    doc.build(c, onFirstPage=footer, onLaterPages=footer)

def write_excel_report(out_xlsx, meta, df, analytics, charts):
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        analytics["kpis"].to_excel(writer, sheet_name="Summary", index=False, startrow=4)
        pd.DataFrame({"Insight": analytics["insights"]}).to_excel(writer, sheet_name="Summary", index=False, startrow=4 + len(analytics["kpis"]) + 3)
        analytics["segment_counts"].to_excel(writer, sheet_name="Member_Segments", index=False)
        analytics["batch_analysis"].to_excel(writer, sheet_name="Batch_Analysis", index=False)
        analytics["daily_totals"].to_excel(writer, sheet_name="Daily_Attendance", index=False)
        analytics["dow_totals"].to_excel(writer, sheet_name="Day_of_Week", index=False)
        analytics["top_attendees"].head(50).to_excel(writer, sheet_name="Top_Attendees", index=False)
        analytics["least_active"].head(50).to_excel(writer, sheet_name="Least_Active", index=False)
        df.to_excel(writer, sheet_name="Clean_Data", index=False)
        
    wb = load_workbook(out_xlsx)
    for ws in wb.worksheets:
        for c in ws[1]: c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="8A1E1E"), Alignment(horizontal="center")
        ws.freeze_panes = "A2"
    ws = wb["Summary"]
    ws["A1"], ws["B1"], ws["A2"], ws["B2"], ws["A3"], ws["B3"] = "Center", meta.center_name, "Month", meta.report_month, "Period", meta.report_period
    for cell in ["A1", "A2", "A3"]: ws[cell].font = Font(bold=True)
    
    try:
        ws.add_image(XLImage(charts["dow"]), "E2")
        ws.add_image(XLImage(charts["trend"]), "E22")
    except: pass
    wb.save(out_xlsx)

# =============================================================================
# Streamlit Web UI
# =============================================================================
st.set_page_config(page_title="Yoga Center Reports", page_icon="🧘", layout="centered")

st.title("🧘 Yoga Center Attendance Dashboard")
st.markdown("Upload your monthly attendance Excel file to instantly generate your Manager PDF Dashboard and Excel Audit file.")

uploaded_file = st.file_uploader("1. Upload Attendance Excel (.xlsx)", type=["xlsx"])
uploaded_logo = st.file_uploader("2. (Optional) Upload Center Logo (.png, .jpg)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    if st.button("Generate Reports ⚙️", type="primary", use_container_width=True):
        with st.spinner("Crunching numbers and building reports..."):
            try:
                # Create a secure temporary workspace for the current user
                with tempfile.TemporaryDirectory() as tmpdir:
                    in_path = os.path.join(tmpdir, "input.xlsx")
                    with open(in_path, "wb") as f: f.write(uploaded_file.getvalue())

                    logo_path = None
                    if uploaded_logo:
                        logo_path = os.path.join(tmpdir, "logo.png")
                        with open(logo_path, "wb") as f: f.write(uploaded_logo.getvalue())

                    pdf_path = os.path.join(tmpdir, "Report.pdf")
                    xlsx_path = os.path.join(tmpdir, "Audit.xlsx")

                    # Execute pipeline
                    df, date_cols, date_idx, meta, presence = read_and_clean(in_path)
                    analytics = build_analytics(df, date_cols, date_idx, presence)
                    charts = plot_charts(analytics, tmpdir)
                    build_pdf(pdf_path, meta, analytics, charts, logo_path)
                    write_excel_report(xlsx_path, meta, df, analytics, charts)

                    # Read generated files into memory for user download
                    with open(pdf_path, "rb") as f: pdf_bytes = f.read()
                    with open(xlsx_path, "rb") as f: xlsx_bytes = f.read()

                st.success(f"✅ Success! Generated reports for {meta.center_name} ({meta.report_month})")
                
                # Render side-by-side download buttons
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📄 Download PDF Dashboard",
                        data=pdf_bytes,
                        file_name=f"{meta.center_name}_Report_{meta.month_key}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                with col2:
                    st.download_button(
                        label="📊 Download Excel Audit",
                        data=xlsx_bytes,
                        file_name=f"{meta.center_name}_Audit_{meta.month_key}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"❌ An error occurred while processing your file.")
                st.exception(e)