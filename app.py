"""
Yoga Center Attendance & Fee Status Dashboard
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
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak

PAGE_SIZE = A4
LEFT_MARGIN, RIGHT_MARGIN = 14 * mm, 14 * mm
TOP_MARGIN, BOTTOM_MARGIN = 20 * mm, 20 * mm

BRAND_RED = colors.HexColor("#8A1E1E")
BRAND_DARK = colors.HexColor("#222222")
BRAND_GREY = colors.HexColor("#666666")
BRAND_LIGHT = colors.HexColor("#F9F9F9")
CHART_RED = "#8A1E1E"
DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

class ReportValidationError(Exception):
    pass

@dataclass
class ReportMeta:
    center_name: str
    report_month: str
    report_period: str
    month_key: str
    date_start: datetime

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
    replacements = {"vijayanagar(blore)": "Vijayanagara", "vijayanagar": "Vijayanagara", "sadashivanagar": "Sadashivanagara", "gayathri nagar": "Gayathri Nagar", "gyn": "Gayathri Nagar"}
    for k, v in replacements.items():
        if lower.replace(" ", "") == k.replace(" ", ""): return v.upper()
    return txt.upper()

def clean_batch_name(raw_name: str) -> str:
    txt = safe_str(raw_name)
    txt = re.sub(r'(?i)(general)\s+general', r'\1', txt)
    txt = re.sub(r'(?i)(therapy)\s+therapy', r'\1', txt)
    txt = re.sub(r'(?i)(junior)\s+junior', r'\1', txt)
    txt = re.sub(r'(?i)(senior)\s+senior', r'\1', txt)
    txt = re.sub(r'(?i)children\s*yoga.*children\s*yoga.*', 'Children Yoga', txt)
    return txt.strip()

def extract_numeric_id(val) -> str:
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    match = re.search(r'(\d+)$', s)
    v = match.group(1) if match else s
    return v.lstrip('0') or '0'

def find_header_row(path: str, target_cols: list) -> int:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in range(1, min(30, ws.max_row) + 1):
        values = [safe_str(ws.cell(r, c).value).lower() for c in range(1, min(ws.max_column, 25) + 1)]
        if any(t.lower() in values for t in target_cols): return r - 1
    return -1

def attendance_cell_to_bool(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.lower()
    return s.isin({"present", "p", "x", "1", "yes", "y", "true"})

def extract_metadata(raw: pd.DataFrame) -> str:
    for _, row in raw.head(5).iterrows():
        vals = [safe_str(v) for v in row.tolist()[:6]]
        joined = " | ".join(v for v in vals if v)
        if any(k in joined.lower() for k in ["vijayanagar", "sadashivanagar", "jayanagar", "rysri", "gayathri"]):
            for v in vals:
                if v and not re.search(r"^\d{1,2}[-/][A-Za-z]{3}[-/]\d{2,4}$", v):
                    if any(k in v.lower() for k in ["vijayanagar", "sadashivanagar", "jayanagar", "rysri", "gayathri"]): return v
    return ""

def read_fee_report(input_file: str) -> pd.DataFrame:
    header_row = find_header_row(input_file, ["stud id", "stud name", "mobile no"])
    if header_row == -1: raise ReportValidationError("Could not locate the header row in the Fee Report. Please ensure it contains 'Stud ID'.")
        
    raw = pd.read_excel(input_file, header=header_row)
    raw.columns = [str(c).strip() for c in raw.columns]
    
    id_col = next((c for c in raw.columns if "stud id" in c.lower() or "student id" in c.lower()), None)
    status_col = next((c for c in raw.columns if "status" in c.lower()), None)
    mobile_col = next((c for c in raw.columns if "mobile" in c.lower()), None)
    particulars_col = next((c for c in raw.columns if "particulars" in c.lower()), None)
    pymt_dt_col = next((c for c in raw.columns if "pymt dt" in c.lower() or "payment date" in c.lower()), None)
    
    if not id_col or not status_col: raise ReportValidationError("Fee report missing required columns ('Stud ID' or 'Status').")
        
    fee_records = []
    for idx, row in raw.iterrows():
        val_id = row[id_col]
        if pd.notna(val_id):
            num_id = extract_numeric_id(val_id)
            status = safe_str(row[status_col]).lower()
            
            mobile = "N/A"
            if mobile_col and pd.notna(row[mobile_col]):
                m = row[mobile_col]
                if isinstance(m, float): 
                    try: mobile = str(int(m))
                    except: mobile = "N/A"
                else: 
                    mobile = str(m).strip()
                    
            particulars = safe_str(row[particulars_col]) if particulars_col and pd.notna(row[particulars_col]) else ""
            
            pymt_dt = None
            if pymt_dt_col and pd.notna(row[pymt_dt_col]):
                try:
                    pymt_dt = pd.to_datetime(row[pymt_dt_col])
                except:
                    pymt_dt = None
                    
            fee_records.append({
                "Norm_ID": num_id,
                "Status": status,
                "Mobile": mobile,
                "Particulars": particulars,
                "Pymt_Dt": pymt_dt
            })
    return pd.DataFrame(fee_records)

def read_and_clean(input_file: str):
    wb = load_workbook(input_file, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    sample_text = ""
    for r in range(1, min(10, ws.max_row) + 1): sample_text += " ".join([safe_str(ws.cell(r, c).value).lower() for c in range(1, min(ws.max_column, 20) + 1)])
    if "admission fee" in sample_text or "subscription fees" in sample_text:
        raise ReportValidationError("Oops! It looks like you uploaded the Fee Report in the Attendance slot. Please swap them.")
    
    header_row = find_header_row(input_file, ["studentname", "studentid", "slno"])
    if header_row == -1: raise ReportValidationError("Invalid Attendance File format. Could not find 'StudentId' or 'StudentName' headers.")
    
    raw = pd.read_excel(input_file, header=header_row)
    raw.columns = [normalize_col(c) for c in raw.columns]
    raw = raw.dropna(how="all")
    
    if "StudentId" in raw.columns:
        raw = raw.dropna(subset=["StudentId"])
        raw['data_points'] = raw.notna().sum(axis=1)
        raw = raw.sort_values('data_points', ascending=False).drop_duplicates(subset=['StudentId'], keep='first')
        raw = raw.drop(columns=['data_points'])

    raw = raw.reset_index(drop=True)
    if len(raw) == 0: raise ReportValidationError("No valid student records found after cleaning the attendance data.")
        
    status_idx = raw.columns.get_loc("Status") if "Status" in raw.columns else -1
    date_cols, date_index = [], []
    
    start_col = status_idx + 1 if status_idx != -1 else 0
    for c in raw.columns[start_col:]:
        if safe_str(c).strip().lower() in {"present", "absent"}: continue
        dt = pd.to_datetime(c, format="%d-%b-%Y", errors="coerce")
        if pd.notna(dt):
            date_cols.append(c)
            date_index.append(dt)
        else:
            dt2 = pd.to_datetime(c, errors="coerce")
            if pd.notna(dt2):
                date_cols.append(c)
                date_index.append(dt2)

    if not date_cols: raise ReportValidationError("No valid attendance dates were found in this file. Please check the date formatting.")

    sort_idx = np.argsort(pd.to_datetime(date_index))
    date_cols = [date_cols[i] for i in sort_idx]
    date_index = pd.DatetimeIndex([date_index[i] for i in sort_idx])

    df = raw.copy()
    for col in ["Center", "Status", "StudentName", "StudentId"]:
        if col in df.columns: df[col] = df[col].map(safe_str)
    if "Batch" in df.columns: df["Batch"] = df["Batch"].apply(clean_batch_name)

    presence = pd.DataFrame({c: attendance_cell_to_bool(df[c]) for c in date_cols}, index=df.index)
    df["Present_Count_Calc"] = presence.sum(axis=1).astype(int)
    df["System_ID"] = df["StudentId"].apply(extract_numeric_id)
    
    center_raw = df["Center"].mode().iloc[0] if "Center" in df.columns and df["Center"].str.strip().ne("").any() else extract_metadata(raw)
    date_start, date_end = date_index.min(), date_index.max()
    meta = ReportMeta(
        center_name=clean_center_name(center_raw), 
        report_month=date_start.strftime("%B %Y"), 
        report_period=f"{date_start.strftime('%d-%b-%Y')} to {date_end.strftime('%d-%b-%Y')}", 
        month_key=date_start.strftime("%b_%Y"),
        date_start=date_start
    )
    return df, date_cols, date_index, meta, presence

def reconcile_fees(df: pd.DataFrame, fee_df: pd.DataFrame, target_month_tag_regex: str):
    if fee_df.empty:
        return set(), df, pd.DataFrame()
        
    active_fees = fee_df[fee_df['Status'] == 'active'].copy()
    
    active_fees['Covers_Target_Month'] = active_fees['Particulars'].str.contains(target_month_tag_regex, case=False, na=False, regex=True)
    valid_fees = active_fees[active_fees['Covers_Target_Month']].copy()
    
    valid_fees = valid_fees.sort_values(by='Pymt_Dt', ascending=False, na_position='last')
    latest_valid_fees = valid_fees.drop_duplicates(subset=['Norm_ID'], keep='first')
    
    paid_ids = set(latest_valid_fees['Norm_ID'].tolist())
    
    all_active_fees_sorted = active_fees.sort_values(by='Pymt_Dt', ascending=False, na_position='last')
    latest_any_fees = all_active_fees_sorted.drop_duplicates(subset=['Norm_ID'], keep='first')
    
    mobile_map = latest_any_fees.set_index('Norm_ID')['Mobile'].to_dict()
    pymt_dt_map = latest_any_fees.set_index('Norm_ID')['Pymt_Dt'].to_dict()
    coverage_map = latest_any_fees.set_index('Norm_ID')['Particulars'].to_dict()
    
    df['Mobile No'] = df['System_ID'].map(mobile_map).fillna("No Record")
    
    def format_date(d):
        if pd.isna(d) or d is None: return "Never"
        return d.strftime("%d-%b-%Y")
        
    df['Last Payment Date'] = df['System_ID'].map(pymt_dt_map).apply(format_date)
    
    def extract_coverage(s):
        if not s: return "None"
        months = re.findall(r'\[.*?\]', s)
        return ", ".join(months) if months else "None"
        
    df['Subscription Coverage'] = df['System_ID'].map(coverage_map).fillna("").apply(extract_coverage)
    
    attended_mask = df['Present_Count_Calc'] > 0
    not_paid_mask = ~df['System_ID'].isin(paid_ids)
    pending_students = df[attended_mask & not_paid_mask].copy()
    
    cols_to_keep = ['StudentName', 'System_ID', 'Present_Count_Calc', 'Mobile No', 'Last Payment Date', 'Subscription Coverage']
    if 'Status' in pending_students.columns:
        cols_to_keep.insert(2, 'Status')
        
    pending_students = pending_students[cols_to_keep].sort_values(by='Present_Count_Calc', ascending=False).reset_index(drop=True)
    pending_students.rename(columns={'Present_Count_Calc': 'Days Attended'}, inplace=True)
    
    return paid_ids, df, pending_students

def build_analytics(df, date_cols, date_index, presence, fee_df=None, meta=None):
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

    daily_totals["WeekOfMonth"] = ((daily_totals["Date"].dt.day - 1) // 7) + 1
    weekly_totals = daily_totals.groupby("WeekOfMonth", as_index=False)["Attendance"].sum()
    weekly_totals["Week"] = "Week " + weekly_totals["WeekOfMonth"].astype(str)
    weekly_totals = weekly_totals[["Week", "Attendance"]]
    
    busiest_week = weekly_totals.loc[weekly_totals['Attendance'].idxmax()]
    peak_day_row = daily_totals.loc[daily_totals['Attendance'].idxmax()]

    total_att = int(daily_totals["Attendance"].sum())
    avg_visits = round(total_att/max(1, len(df)), 1)
    kpis = pd.DataFrame({"Metric": ["Registered Members", "Operating Days", "Total Center Visits", "Avg Visits per Member", "Busiest Day"], "Value": [len(df), len(date_cols), total_att, avg_visits, dow_totals.loc[dow_totals['Attendance'].idxmax(), 'DayOfWeek']]})
    
    insights = [f"Habit Building: On average, each member visited {avg_visits} times this month.", f"Batch Popularity: The '{batch_analysis.iloc[0]['Batch'] if not batch_analysis.empty else 'N/A'}' batch had the highest footfall.", f"Peak Activity: {kpis.iloc[4]['Value']} saw the most activity across the month."]
    op_insights = [f"Highest Single Day: The center saw its peak daily footfall on {peak_day_row['ShortDate']} with {peak_day_row['Attendance']} recorded visits.", f"Busiest Period: {busiest_week['Week']} was the most active operational period, capturing {busiest_week['Attendance']} total visits."]
    
    pending_students = pd.DataFrame()
    fee_summary_text = ""
    
    if fee_df is not None and not fee_df.empty and meta is not None:
        month_str = MONTH_ABBR[meta.date_start.month-1]
        year_str = meta.date_start.strftime('%y')
        target_month_tag_regex = rf"{month_str}[\s\-/]*{year_str}"
        
        paid_ids, df, pending_students = reconcile_fees(df, fee_df, target_month_tag_regex)
        
        total_attending = len(df[df['Present_Count_Calc'] > 0])
        total_unpaid = len(pending_students)
        
        if total_attending > 0:
            unpaid_ratio = total_unpaid / total_attending
            if unpaid_ratio > 0.70:
                raise ReportValidationError(f"Data Anomaly Detected: {int(unpaid_ratio*100)}% of attending students are flagged as unpaid. This exceeds the 70% safety threshold and indicates a potential ID mismatch or missing fee data. Please verify the input files.")
        
        total_pending_visits = int(pending_students['Days Attended'].sum()) if not pending_students.empty else 0
        total_pending_individuals = len(pending_students)
        
        if total_pending_individuals > 0:
            fee_summary_text = f"A total of {total_pending_individuals} students attended classes without an active fee subscription for {month_str} {year_str}, accumulating {total_pending_visits} attendance days."
        else:
            fee_summary_text = f"Fee Status: Excellent. All attending students currently have active subscriptions covering {month_str} {year_str}."

    return {"presence": presence, "daily_totals": daily_totals, "dow_totals": dow_totals, "weekly_totals": weekly_totals, "batch_analysis": batch_analysis, "top_attendees": top_attendees, "least_active": least_active, "segment_counts": segment_counts, "kpis": kpis, "insights": insights, "op_insights": op_insights, "pending_students": pending_students, "fee_summary_text": fee_summary_text, "df": df}

def plot_charts(analytics, tmpdir):
    dow = analytics["dow_totals"]
    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.bar(dow["DayOfWeek"], dow["Attendance"], color=CHART_RED)
    ax.set_title("Attendance by Day of Week", fontweight="bold")
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for bar in bars: ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + dow["Attendance"].max()*0.02, str(int(bar.get_height())), ha="center", va="bottom", fontweight="bold")
    plt.xticks(rotation=0)
    plt.tight_layout()
    dow_path = os.path.join(tmpdir, "dow.png")
    plt.savefig(dow_path, dpi=150)
    plt.close(fig)

    daily = analytics["daily_totals"]
    fig, ax = plt.subplots(figsize=(8, 4))
    day_numbers = daily["Date"].dt.strftime("%d")
    ax.plot(day_numbers, daily["Attendance"], marker="o", color=CHART_RED, linewidth=2)
    ax.set_title("Daily Attendance Trend", fontweight="bold")
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.5)
    for x, y in zip(day_numbers, daily["Attendance"]): ax.text(x, y + daily["Attendance"].max()*0.02, str(int(y)), ha="center", fontsize=8)
    ax.set_xticks(day_numbers[::2])
    plt.xticks(rotation=0)
    plt.tight_layout()
    trend_path = os.path.join(tmpdir, "trend.png")
    plt.savefig(trend_path, dpi=150)
    plt.close(fig)
    return {"dow": dow_path, "trend": trend_path}

def make_pdf_table(df, headers, widths):
    tbl = Table([headers] + df.astype(str).values.tolist(), colWidths=widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BRAND_RED), ("TEXTCOLOR", (0,0), (-1,0), colors.white), 
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("BOTTOMPADDING", (0,0), (-1,0), 8), ("TOPPADDING", (0,0), (-1,0), 8),
        ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, BRAND_LIGHT]), 
        ("FONTSIZE", (0,0), (-1,-1), 8.5), ("ALIGN", (1,0), (-1,-1), "CENTER"), ("VALIGN", (0,0), (-1,-1), "MIDDLE")
    ]))
    return tbl

def build_pdf(out_pdf, meta, analytics, charts):
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name="ReportTitle", fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=BRAND_DARK, spaceAfter=6)
    subtitle_style = ParagraphStyle(name="ReportSubtitle", fontName="Helvetica", fontSize=12, leading=14, textColor=BRAND_GREY, spaceAfter=18)
    section_style = ParagraphStyle(name="SectionHeader", fontName="Helvetica-Bold", fontSize=12, textColor=colors.white, backColor=BRAND_RED, spaceBefore=15, spaceAfter=10, borderPadding=(4, 6, 4, 6))
    bullet_style = ParagraphStyle(name="Bullet", parent=styles["BodyText"], leftIndent=15, spaceAfter=5, bulletIndent=5)
    
    doc = SimpleDocTemplate(out_pdf, pagesize=PAGE_SIZE, rightMargin=RIGHT_MARGIN, leftMargin=LEFT_MARGIN, topMargin=TOP_MARGIN, bottomMargin=BOTTOM_MARGIN, title=f"{meta.center_name} - Attendance Report")
    c = []
    
    c.append(Paragraph(f"{meta.center_name} - Attendance Report", title_style))
    c.append(Paragraph(f"Reporting Period: {meta.report_period}", subtitle_style))
    c.append(Spacer(1, 2*mm))

    c.append(Paragraph("1. Executive Summary", section_style))
    kpi_t = Table([["Metric", "Value"]] + analytics["kpis"].astype(str).values.tolist(), colWidths=[100*mm, 60*mm])
    kpi_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BRAND_DARK), ("TEXTCOLOR", (0,0), (-1,0), colors.white), 
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, BRAND_LIGHT]), ("FONTNAME", (0,1), (0,-1), "Helvetica-Bold"),
        ("ALIGN", (1,0), (1,-1), "CENTER"), ("BOTTOMPADDING", (0,0), (-1,-1), 6), ("TOPPADDING", (0,0), (-1,-1), 6)
    ]))
    c.append(kpi_t)
    c.append(Spacer(1, 6*mm))
    
    c.append(Paragraph("Key Monthly Insights:", ParagraphStyle(name="Sub", fontName="Helvetica-Bold", spaceAfter=6)))
    for ins in analytics["insights"]: c.append(Paragraph(f"• {ins}", bullet_style))
    c.append(Spacer(1, 8*mm))

    c.append(Paragraph("2. Attendance Trends & Highlights", section_style))
    for op_ins in analytics["op_insights"]: c.append(Paragraph(f"• {op_ins}", bullet_style))
    c.append(Spacer(1, 4*mm))
    
    week_t = Table([["Weekly Operational Period", "Total Center Visits"]] + analytics["weekly_totals"].astype(str).values.tolist(), colWidths=[80*mm, 60*mm])
    week_t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BRAND_DARK), ("TEXTCOLOR", (0,0), (-1,0), colors.white), 
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, BRAND_LIGHT]), ("ALIGN", (1,0), (1,-1), "CENTER"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6), ("TOPPADDING", (0,0), (-1,-1), 6)
    ]))
    c.append(week_t)
    c.append(Spacer(1, 6*mm))
    
    c.append(Image(charts["dow"], width=160*mm, height=70*mm))
    c.append(Spacer(1, 4*mm))
    c.append(Image(charts["trend"], width=160*mm, height=70*mm))
    c.append(PageBreak())

    c.append(Paragraph("3. Member Insights & Demographics", section_style))
    c.append(Spacer(1, 2*mm))
    
    top_disp = analytics["top_attendees"].head(10).copy()
    top_disp["StudentName"] = top_disp["StudentName"].apply(lambda x: (x[:14] + '..') if len(str(x)) > 16 else x)
    least_disp = analytics["least_active"].head(10).copy()
    least_disp["StudentName"] = least_disp["StudentName"].apply(lambda x: (x[:14] + '..') if len(str(x)) > 16 else x)

    w1 = Table([[make_pdf_table(top_disp, ["Top Attendees", "Batch", "Visits"], [45*mm, 32*mm, 17*mm]), make_pdf_table(least_disp, ["Least Active", "Batch", "Visits"], [45*mm, 32*mm, 17*mm])]])
    w1.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    w2 = Table([[make_pdf_table(analytics["segment_counts"], ["Engagement Segment", "Count"], [67*mm, 27*mm]), make_pdf_table(analytics["batch_analysis"].head(5)[["Batch", "Members", "Total_Present"]], ["Batch", "Users", "Visits"], [57*mm, 18*mm, 19*mm])]])
    w2.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    c.extend([w1, Spacer(1, 8*mm), w2])
    
    if analytics.get("fee_summary_text"):
        c.append(Spacer(1, 8*mm))
        c.append(Paragraph("4. Fee Reconciliation: Unpaid Attendees", section_style))
        c.append(Spacer(1, 4*mm))
        
        summary_style = ParagraphStyle(name="FeeSummary", fontName="Helvetica-Bold", fontSize=11, textColor=BRAND_DARK, spaceAfter=12)
        c.append(Paragraph(analytics["fee_summary_text"], summary_style))
        
        if not analytics["pending_students"].empty:
            def_disp = analytics["pending_students"].copy()
            def_disp["StudentName"] = def_disp["StudentName"].apply(lambda x: (x[:16] + '..') if len(str(x)) > 18 else x)
            
            cols = ['StudentName', 'System_ID', 'Days Attended', 'Mobile No', 'Last Payment Date', 'Subscription Coverage']
            headers = ["Student Name", "Student ID", "Days Attended", "Mobile No", "Last Payment", "Coverage"]
            widths = [35*mm, 25*mm, 20*mm, 30*mm, 25*mm, 37*mm]
            
            if 'Status' in def_disp.columns:
                cols.insert(2, 'Status')
                headers.insert(2, "Status")
                widths = [32*mm, 22*mm, 18*mm, 18*mm, 28*mm, 22*mm, 32*mm]

            c.append(make_pdf_table(def_disp[cols], headers, widths))

    def draw_page_setup(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(BRAND_RED)
        canvas.setLineWidth(1)
        canvas.line(LEFT_MARGIN, BOTTOM_MARGIN - 5*mm, A4[0] - RIGHT_MARGIN, BOTTOM_MARGIN - 5*mm)
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(BRAND_GREY)
        canvas.drawString(LEFT_MARGIN, BOTTOM_MARGIN - 10*mm, "Quality and Systems, Rashtrotthana Parishat")
        canvas.drawRightString(A4[0] - RIGHT_MARGIN, BOTTOM_MARGIN - 10*mm, f"Page {doc.page}")
        canvas.restoreState()

    doc.build(c, onFirstPage=draw_page_setup, onLaterPages=draw_page_setup)

def write_excel_report(out_xlsx, meta, df, analytics, charts):
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        analytics["kpis"].to_excel(writer, sheet_name="Summary", index=False, startrow=4)
        
        all_insights = analytics["insights"] + analytics["op_insights"]
        if analytics.get("fee_summary_text"): all_insights.append(analytics["fee_summary_text"])
        pd.DataFrame({"Insight": all_insights}).to_excel(writer, sheet_name="Summary", index=False, startrow=4 + len(analytics["kpis"]) + 3)
        
        if not analytics["pending_students"].empty:
            analytics["pending_students"].to_excel(writer, sheet_name="Pending_Renewals_List", index=False)
            
        analytics["segment_counts"].to_excel(writer, sheet_name="Member_Segments", index=False)
        analytics["batch_analysis"].to_excel(writer, sheet_name="Batch_Analysis", index=False)
        analytics["weekly_totals"].to_excel(writer, sheet_name="Weekly_Trends", index=False)
        analytics["daily_totals"].to_excel(writer, sheet_name="Daily_Attendance", index=False)
        analytics["dow_totals"].to_excel(writer, sheet_name="Day_of_Week", index=False)
        analytics["top_attendees"].head(50).to_excel(writer, sheet_name="Top_Attendees", index=False)
        analytics["least_active"].head(50).to_excel(writer, sheet_name="Least_Active", index=False)
        df.drop(columns=['System_ID'], errors='ignore').to_excel(writer, sheet_name="Clean_Data", index=False)
        
    wb = load_workbook(out_xlsx)
    for ws in wb.worksheets:
        for c_cell in ws[1]: c_cell.font, c_cell.fill, c_cell.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="8A1E1E"), Alignment(horizontal="center")
        ws.freeze_panes = "A2"
    ws = wb["Summary"]
    ws["A1"], ws["B1"], ws["A2"], ws["B2"], ws["A3"], ws["B3"] = "Center", meta.center_name, "Month", meta.report_month, "Period", meta.report_period
    for cell in ["A1", "A2", "A3"]: ws[cell].font = Font(bold=True)
    try: ws.add_image(XLImage(charts["dow"]), "E2"); ws.add_image(XLImage(charts["trend"]), "E22")
    except: pass
    wb.save(out_xlsx)

st.set_page_config(page_title="Yoga Center Operations Manager", page_icon="🧘", layout="centered")
st.title("🧘 Yoga Center Operations Manager")
st.markdown("Upload your monthly attendance to generate the standard report. **Upload the Fee Report alongside it to activate the Fee Reconciliation checks.**")

uploaded_att = st.file_uploader("1. Required: Upload Attendance Report (.xlsx, .csv)", type=["xlsx", "csv"])
uploaded_fee = st.file_uploader("2. Optional: Upload Fee Report (.xlsx, .csv)", type=["xlsx", "csv"])

if uploaded_att is not None:
    if st.button("Generate Reports ⚙️", type="primary", use_container_width=True):
        with st.spinner("Crunching numbers and building reports..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    if uploaded_att.name.endswith('.csv'):
                        att_path = os.path.join(tmpdir, "input.csv")
                        with open(att_path, "wb") as f: f.write(uploaded_att.getvalue())
                        pd.read_csv(att_path, header=None).to_excel(os.path.join(tmpdir, "input.xlsx"), index=False, header=False)
                        att_path = os.path.join(tmpdir, "input.xlsx")
                    else:
                        att_path = os.path.join(tmpdir, "input.xlsx")
                        with open(att_path, "wb") as f: f.write(uploaded_att.getvalue())

                    fee_df = None
                    if uploaded_fee:
                        if uploaded_fee.name.endswith('.csv'):
                            fee_path = os.path.join(tmpdir, "fee.csv")
                            with open(fee_path, "wb") as f: f.write(uploaded_fee.getvalue())
                            pd.read_csv(fee_path, header=None).to_excel(os.path.join(tmpdir, "fee.xlsx"), index=False, header=False)
                            fee_path = os.path.join(tmpdir, "fee.xlsx")
                        else:
                            fee_path = os.path.join(tmpdir, "fee.xlsx")
                            with open(fee_path, "wb") as f: f.write(uploaded_fee.getvalue())
                        fee_df = read_fee_report(fee_path)

                    pdf_path = os.path.join(tmpdir, "Report.pdf"); xlsx_path = os.path.join(tmpdir, "Audit.xlsx")
                    df, date_cols, date_idx, meta, presence = read_and_clean(att_path)
                    analytics = build_analytics(df, date_cols, date_idx, presence, fee_df, meta)
                    
                    df = analytics["df"]
                    
                    charts = plot_charts(analytics, tmpdir)
                    build_pdf(pdf_path, meta, analytics, charts)
                    write_excel_report(xlsx_path, meta, df, analytics, charts)

                    with open(pdf_path, "rb") as f: pdf_bytes = f.read()
                    with open(xlsx_path, "rb") as f: xlsx_bytes = f.read()

                success_msg = f"✅ Success! Standard reports generated for {meta.center_name}." if fee_df is None else f"✅ Success! Advanced Attendance & Fee Reconciliation reports generated for {meta.center_name}."
                st.success(success_msg)
                
                col1, col2 = st.columns(2)
                with col1: st.download_button(label="📄 Download PDF Dashboard", data=pdf_bytes, file_name=f"{meta.center_name}_Report_{meta.month_key}.pdf", mime="application/pdf", use_container_width=True)
                with col2: st.download_button(label="📊 Download Excel Audit", data=xlsx_bytes, file_name=f"{meta.center_name}_Audit_{meta.month_key}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            except ReportValidationError as ve: st.warning(f"⚠️ {str(ve)}")
            except Exception as e:
                st.error("❌ An error occurred. Please ensure the files are in the correct format.")
                st.exception(e)
