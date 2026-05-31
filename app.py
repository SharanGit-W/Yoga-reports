"""
Yoga Center Attendance Generator - Web Dashboard
"""
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple, List
from datetime import datetime

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
TOP_MARGIN, BOTTOM_MARGIN = 20 * mm, 20 * mm  # Increased for header/footer space

BRAND_RED = colors.HexColor("#8A1E1E")
BRAND_DARK = colors.HexColor("#222222")
BRAND_GREY = colors.HexColor("#666666")
BRAND_LIGHT = colors.HexColor("#F9F9F9")
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
# Core Logic (Updated with Strict Logic & Duplicate Handling)
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
    # STRICT LOGIC: Force string conversion and strictly check for explicit present markers.
    s = series.astype(str).str.strip().str.lower()
    return s.isin({"present", "p", "x", "1", "yes", "y", "true"})

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
    
    # Remove entirely empty rows
    raw = raw.dropna(how="all")
    
    # DUPLICATE HANDLING: Ensure StudentId exists, sort by data richness, and drop duplicates cleanly
    if "StudentId" in raw.columns:
        raw = raw.dropna(subset=["StudentId"])
        raw['data_points'] = raw.notna().sum(axis=1)
        raw = raw.sort_values('data_points', ascending=False).drop_duplicates(subset=['StudentId'], keep='first')
        raw = raw.drop(columns=['data_points'])

    raw = raw.reset_index(drop=True)

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
    avg_visits = round(total_att/max(1, len(df)), 1)
    kpis = pd.DataFrame({
        "Metric": ["Registered Members", "Operating Days", "Total Center Visits", "Avg Visits per Member", "Busiest Day"],
        "Value": [len(df), len(date_cols), total_att, avg_visits, dow_totals.loc[dow_totals['Attendance'].idxmax(), 'DayOfWeek']]
    })
    
    insights = [
        f"Habit Building: On average, each member visited {avg_visits} times this month.",
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
    
    # Hide top and right spines for a cleaner look
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    for bar in bars: ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + dow["Attendance"].max()*0.02, str(int(bar.get_height())), ha="center", va="bottom", fontweight="bold")
    plt.xticks(rotation=0)  # Straight text if possible
    plt.tight_layout()
    dow_path = os.path.join(tmpdir, "dow.png")
    plt.savefig(dow_path, dpi=150)
    plt.close(fig)

    daily = analytics["daily_totals"]
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(daily["ShortDate"], daily["Attendance"], marker="o", color=CHART_RED, linewidth=2)
    ax.set_title("Daily Attendance Trend", fontweight="bold")
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.5)
    
    for x, y in zip(daily["ShortDate"], daily["Attendance"]): ax.text(x, y + daily["Attendance"].max()*0.02, str(int(y)), ha="center", fontsize=8)
    plt.xticks(rotation=45)
    plt.tight_layout()
    trend_path = os.path.join(tmpdir, "trend.png")
    plt.savefig(trend_path, dpi=150)
    plt.close(fig)
    return {"dow": dow_path, "trend": trend_path}

def make_pdf_table(df, headers, widths):
    tbl = Table([headers] + df.astype(str).values.tolist(), colWidths=widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BRAND_RED), 
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), 
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("TOPPADDING", (0,0), (-1,0), 8),
        ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey), 
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, BRAND_LIGHT]), 
        ("FONTSIZE", (0,0), (-1,-1), 8.5), 
        ("ALIGN", (1,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE")
    ]))
    return tbl

def build_pdf(out_pdf, meta, analytics, charts, logo=None):
    styles = getSampleStyleSheet()
    
    # Custom PDF Styles for Professional Look
    title_style = ParagraphStyle(name="ReportTitle", fontName="Helvetica-Bold", fontSize=20, textColor=BRAND_DARK, spaceAfter=2)
    subtitle_style = ParagraphStyle(name="ReportSubtitle", fontName="Helvetica", fontSize=12, textColor=BRAND_GREY, spaceAfter=15)
    
    section_style = ParagraphStyle(
        name="SectionHeader",
        fontName="Helvetica-Bold",
        fontSize=12,
        textColor=colors.white,
        backColor=BRAND_RED,
        spaceBefore=15,
        spaceAfter=10,
        borderPadding=(4, 6, 4, 6)
    )
    
    bullet_style = ParagraphStyle(name="Bullet", parent=styles["BodyText"], leftIndent=15, spaceAfter=5, bulletIndent=5)
    
    doc = SimpleDocTemplate(out_pdf, pagesize=PAGE_SIZE, rightMargin=RIGHT_MARGIN, leftMargin=LEFT_MARGIN, topMargin=TOP_MARGIN, bottomMargin=BOTTOM_MARGIN)
    c = []
    
    # Header Information
    c.append(Paragraph(f"{meta.center_name} - Attendance Report", title_style))
    c.append(Paragraph(f"Reporting Period: {meta.report_period}", subtitle_style))
    c.append(Spacer(1, 2*mm))

    # SECTION 1: Executive Summary
    c.append(Paragraph("1. Executive Summary", section_style))
    
    # Better KPI Table
    kpi_data = [["Metric", "Value"]] + analytics["kpis"].astype(str).values.tolist()
    kpi_t = Table(kpi_data, colWidths=[100*mm, 60*mm])
    kpi_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BRAND_DARK), 
        ("TEXTCOLOR", (0,0), (-1,0), colors.white), 
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, BRAND_LIGHT]),
        ("FONTNAME", (0,1), (0,-1), "Helvetica-Bold"), # Bold metric names
        ("ALIGN", (1,0), (1,-1), "CENTER"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 6)
    ]))
    c.append(kpi_t)
    c.append(Spacer(1, 6*mm))
    
    # Insights
    c.append(Paragraph("Key Monthly Insights:", ParagraphStyle(name="Sub", fontName="Helvetica-Bold", spaceAfter=6)))
    for ins in analytics["insights"]: 
        c.append(Paragraph(f"• {ins}", bullet_style))

    c.append(Spacer(1, 8*mm))

    # SECTION 2: Attendance Trends
    c.append(Paragraph("2. Attendance Trends", section_style))
    c.append(Image(charts["dow"], width=165*mm, height=75*mm))
    c.append(Spacer(1, 4*mm))
    c.append(Image(charts["trend"], width=165*mm, height=75*mm))
    
    c.append(PageBreak())

    # SECTION 3: Member Insights
    c.append(Paragraph("3. Member Insights & Demographics", section_style))
    c.append(Spacer(1, 2*mm))
    
    w1 = Table([[
        make_pdf_table(analytics["top_attendees"].head(10), ["Top Attendees", "Batch", "Visits"], [48*mm, 30*mm, 17*mm]), 
        make_pdf_table(analytics["least_active"].head(10), ["Least Active", "Batch", "Visits"], [48*mm, 30*mm, 17*mm])
    ]])
    w1.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    w2 = Table([[
        make_pdf_table(analytics["segment_counts"], ["Engagement Segment", "Count"], [68*mm, 27*mm]), 
        make_pdf_table(analytics["batch_analysis"].head(5)[["Batch", "Members", "Total_Present"]], ["Batch", "Users", "Visits"], [58*mm, 18*mm, 19*mm])
    ]])
    w2.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    c.extend([w1, Spacer(1, 8*mm), w2])
    
    # Custom Page Setup for Header/Footer
    def draw_page_setup(canvas, doc):
        canvas.saveState()
        
        # Logo Logic
        if logo and os.path.exists(logo):
            try: 
                canvas.drawImage(logo, A4[0]-RIGHT_MARGIN-28*mm, A4[1]-18*mm, width=28*mm, height=14*mm, preserveAspectRatio=True, mask='auto')
            except: 
                pass
                
        # Footer Divider Line
        canvas.setStrokeColor(BRAND_RED)
        canvas.setLineWidth(1)
        canvas.line(LEFT_MARGIN, BOTTOM_MARGIN - 5*mm, A4[0] - RIGHT_MARGIN, BOTTOM_MARGIN - 5*mm)
        
        # Footer Text
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(BRAND_GREY)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        canvas.drawString(LEFT_MARGIN, BOTTOM_MARGIN - 10*mm, f"Generated automatically on {timestamp}")
        canvas.drawRightString(A4[0] - RIGHT_MARGIN, BOTTOM_MARGIN - 10*mm, f"Page {doc.page}")
        
        canvas.restoreState()

    doc.build(c, onFirstPage=draw_page_setup, onLaterPages=draw_page_setup)

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

uploaded_file = st.file_uploader("1. Upload Attendance Excel (.xlsx, .csv)", type=["xlsx", "csv"])
uploaded_logo = st.file_uploader("2. (Optional) Upload Center Logo (.png, .jpg)", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    if st.button("Generate Reports ⚙️", type="primary", use_container_width=True):
        with st.spinner("Crunching numbers and building reports..."):
            try:
                # Create a secure temporary workspace
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Detect if CSV or Excel based on filename
                    if uploaded_file.name.endswith('.csv'):
                        in_path = os.path.join(tmpdir, "input.csv")
                        with open(in_path, "wb") as f: f.write(uploaded_file.getvalue())
                        # If the user uploads a CSV directly, use pandas to read and save it as an Excel immediately so find_header_row doesn't break
                        df_temp = pd.read_csv(in_path, header=None)
                        in_path = os.path.join(tmpdir, "input.xlsx")
                        df_temp.to_excel(in_path, index=False, header=False)
                    else:
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

                    # Read generated files into memory
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
