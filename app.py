import re
import io
import pandas as pd
import streamlit as st
import matplotlib
import matplotlib.pyplot as plt

# Force Matplotlib to not use any Xwindows backend (Crucial for cloud deployment)
matplotlib.use('Agg')

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm

# ==========================================
# 1. HELPER & PARSING FUNCTIONS
# ==========================================

def load_excel_smart(file_bytes, target_col):
    """
    Dynamically finds the header row containing the target column.
    Robust against title text above the actual table headers.
    """
    # Use engine='openpyxl' for xlsx files, fallback gracefully for csv/xls if needed
    try:
        df = pd.read_excel(file_bytes, header=None, engine="openpyxl")
    except Exception:
        # Fallback to default if it's an older format
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, header=None)
        
    header_idx = 0
    target_clean = "".join(target_col.lower().split())
    
    for i, row in df.iterrows():
        row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
        # Must find the target column AND the word "batch" to confirm it's the real header row
        if (target_col.lower() in row_str or target_clean in "".join(row_str.split())) and "batch" in row_str:
            header_idx = i
            break
            
    new_columns = [str(x).strip() if pd.notna(x) else f"Unnamed_{j}" for j, x in enumerate(df.iloc[header_idx])]
    df.columns = new_columns
    df = df.iloc[header_idx + 1:].reset_index(drop=True)
    df = df.dropna(how='all')
    return df

def normalize_name(name):
    if pd.isna(name):
        return ""
    return " ".join(str(name).strip().lower().split())

def normalize_batch(batch):
    if pd.isna(batch):
        return ""
    batch = str(batch).lower().strip()
    # CRITICAL FIX: Handle batches like "General,General"
    batch = batch.split(',')[0] 
    batch = batch.replace("batch", "")
    batch = re.sub(r'\s+y[gc]\b', '', batch)
    return " ".join(batch.split())

MONTH_REGEX = r"\[\s*([A-Za-z]{3}\s*[-/\s]\s*\d{2,4})\s*\]"

def parse_months(particulars: str):
    if pd.isna(particulars):
        return []
    matches = re.findall(MONTH_REGEX, str(particulars), flags=re.IGNORECASE)
    cleaned = []
    for month_str in matches:
        m_clean = month_str.replace(" ", "").replace("/", "-")
        try:
            # CRITICAL FIX: Explicitly enforce format so "May-26" isn't parsed as May 26th of the current year!
            dt = pd.to_datetime(m_clean, format="%b-%y")
            cleaned.append(dt.strftime("%b-%y"))
        except ValueError:
            try:
                # Fallback for "Jan-2026"
                dt = pd.to_datetime(m_clean)
                cleaned.append(dt.strftime("%b-%y"))
            except Exception:
                cleaned.append(m_clean.title())
    return cleaned

@st.cache_data
def load_fee_data(file_bytes):
    fee = load_excel_smart(io.BytesIO(file_bytes), "Stud Name")
    
    # Safely check for columns ignoring trailing spaces
    actual_cols = [str(c).strip() for c in fee.columns]
    fee.columns = actual_cols
    
    required_columns = ["Stud Name", "Batch", "Status", "Particulars"]
    for col in required_columns:
        if col not in fee.columns:
            raise ValueError(f"Missing column in Fee Report: {col}. Found: {fee.columns.tolist()}")

    fee = fee[
        ~fee["Status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .isin(["cancelled", "canceled"])
    ]

    fee["Name_Key"] = fee["Stud Name"].apply(normalize_name)
    fee["Batch_Key"] = fee["Batch"].apply(normalize_batch)

    expanded_rows = []
    for _, row in fee.iterrows():
        months = parse_months(row["Particulars"])
        for month in months:
            expanded_rows.append({
                "Student Name": row["Stud Name"],
                "Batch": row["Batch"],
                "Month": month,
                "Name_Key": row["Name_Key"],
                "Batch_Key": row["Batch_Key"],
            })

    paid_df = pd.DataFrame(
        expanded_rows,
        columns=["Student Name", "Batch", "Month", "Name_Key", "Batch_Key"]
    )

    if not paid_df.empty:
        paid_df = paid_df.drop_duplicates(subset=["Name_Key", "Batch_Key", "Month"])

    return paid_df

@st.cache_data
def load_attendance_data(file_bytes):
    attendance = load_excel_smart(io.BytesIO(file_bytes), "StudentName")
    
    actual_cols = [str(c).strip() for c in attendance.columns]
    attendance.columns = actual_cols
    
    required_columns = ["StudentName", "Batch"]
    for col in required_columns:
        if col not in attendance.columns:
            raise ValueError(f"Missing column in Attendance Report: {col}. Found: {attendance.columns.tolist()}")

    attendance["Name_Key"] = attendance["StudentName"].apply(normalize_name)
    attendance["Batch_Key"] = attendance["Batch"].apply(normalize_batch)

    date_columns = []
    col_to_month_map = {}
    
    for col in attendance.columns:
        try:
            col_str = str(col).strip()
            if col_str.isdigit() or len(col_str) < 4:
                continue
                
            # Safely parse Excel columns ensuring "May-26" isn't misread
            if isinstance(col, str) and re.match(r"^[A-Za-z]{3}-\d{2}$", col_str):
                dt = pd.to_datetime(col_str, format="%b-%y")
            else:
                dt = pd.to_datetime(col)
                
            date_columns.append(col)
            col_to_month_map[col] = dt.strftime("%b-%y")
        except Exception:
            continue

    attendance_rows = []
    for _, row in attendance.iterrows():
        monthly_days = {}
        for col in date_columns:
            value = str(row[col]).strip().lower()
            if value in ["present", "p"]:
                month = col_to_month_map[col]
                monthly_days[month] = monthly_days.get(month, 0) + 1

        for month, days in monthly_days.items():
            attendance_rows.append({
                "Student Name": row["StudentName"],
                "Batch": row["Batch"],
                "Month": month,
                "Days Attended": days,
                "Name_Key": row["Name_Key"],
                "Batch_Key": row["Batch_Key"],
            })

    return pd.DataFrame(
        attendance_rows,
        columns=["Student Name", "Batch", "Month", "Days Attended", "Name_Key", "Batch_Key"]
    )

def generate_defaulters(attendance_df, paid_df):
    if attendance_df.empty:
        return pd.DataFrame(columns=["Student Name", "Batch", "Month", "Days Attended"])
        
    if paid_df.empty:
        return attendance_df[
            ["Student Name", "Batch", "Month", "Days Attended"]
        ].sort_values("Days Attended", ascending=False)

    merged = attendance_df.merge(
        paid_df[["Name_Key", "Batch_Key", "Month"]].drop_duplicates(),
        on=["Name_Key", "Batch_Key", "Month"],
        how="left",
        indicator=True
    )
    
    defaulters = merged[merged["_merge"] == "left_only"].copy()
    
    return defaulters[
        ["Student Name", "Batch", "Month", "Days Attended"]
    ].sort_values("Days Attended", ascending=False)

# ==========================================
# 2. PDF GENERATION FUNCTIONS
# ==========================================

def create_batch_chart(defaulters):
    buf = io.BytesIO()
    if len(defaulters) == 0:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No Defaulters", ha="center", va="center")
        plt.savefig(buf, format='png')
        plt.close()
        buf.seek(0)
        return buf

    summary = defaulters.groupby("Batch")["Days Attended"].sum().sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(8, 4))
    summary.plot(kind="bar", ax=ax)
    ax.set_title("Unpaid Attendance Days By Batch")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf

def build_pdf_report(defaulters, metrics):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Yoga Fee Compliance Report", styles["Title"]))
    story.append(Spacer(1, 10))

    kpi_data = [
        ["Metric", "Value"],
        ["Students Attended", metrics["students"]],
        ["Paid Students", metrics["paid"]],
        ["Defaulters", metrics["defaulters"]],
        ["Unpaid Days", metrics["unpaid_days"]],
    ]
    table = Table(kpi_data, colWidths=[200, 100])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ]))
    story.append(table)
    story.append(PageBreak())

    chart_buf = create_batch_chart(defaulters)
    story.append(Paragraph("Batch Summary", styles["Heading1"]))
    
    # CRITICAL FIX: Removed ImageReader wrapper to prevent Type Errors in newer ReportLab versions
    story.append(Image(chart_buf, width=170 * mm, height=90 * mm))
    story.append(PageBreak())

    story.append(Paragraph("Detailed Defaulters", styles["Heading1"]))
    data = [["Student", "Batch", "Month", "Days"]]
    for _, row in defaulters.iterrows():
        data.append([
            row["Student Name"],
            row["Batch"],
            row["Month"],
            int(row["Days Attended"]),
        ])

    details_table = Table(data, repeatRows=1, splitRows=1)
    details_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkred),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.append(details_table)

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# ==========================================
# 3. STREAMLIT UI
# ==========================================

st.set_page_config(page_title="Yoga Fee Dashboard", layout="wide")
st.title("🧘 Yoga Fee Compliance Dashboard")

with st.sidebar:
    st.header("Upload Files")
    fee_file = st.file_uploader("Fee Report (Excel/CSV)", type=["xlsx", "xls", "csv"])
    attendance_file = st.file_uploader("Attendance Report (Excel)", type=["xlsx", "xls"])

if fee_file and attendance_file:
    try:
        with st.spinner("Processing files..."):
            paid_df = load_fee_data(fee_file.getvalue())
            attendance_df = load_attendance_data(attendance_file.getvalue())
            defaulters = generate_defaulters(attendance_df, paid_df)

        students = attendance_df["Student Name"].nunique() if not attendance_df.empty else 0
        paid_students = paid_df["Student Name"].nunique() if not paid_df.empty else 0
        total_defaulters = defaulters["Student Name"].nunique() if not defaulters.empty else 0
        unpaid_days = int(defaulters["Days Attended"].sum()) if not defaulters.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Students", students)
        c2.metric("Paid", paid_students)
        c3.metric("Defaulters", total_defaulters)
        c4.metric("Unpaid Days", unpaid_days)

        st.divider()

        search = st.text_input("Search Student")
        
        if not defaulters.empty:
            month_options = sorted(defaulters["Month"].unique().tolist())
            batch_options = sorted(defaulters["Batch"].unique().tolist())
        else:
            month_options = []
            batch_options = []

        month_filter = st.multiselect("Month", month_options)
        batch_filter = st.multiselect("Batch", batch_options)

        filtered = defaulters.copy()

        if search and not filtered.empty:
            filtered = filtered[filtered["Student Name"].str.contains(search, case=False, na=False)]
        if month_filter and not filtered.empty:
            filtered = filtered[filtered["Month"].isin(month_filter)]
        if batch_filter and not filtered.empty:
            filtered = filtered[filtered["Batch"].isin(batch_filter)]

        st.subheader("Students Attending Without Paying")
        st.dataframe(filtered, use_container_width=True, height=600)

        if not filtered.empty:
            csv_data = filtered.to_csv(index=False).encode("utf-8")
            st.download_button("⬇ Download CSV", csv_data, "defaulters.csv", "text/csv")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                filtered.to_excel(writer, index=False)
            st.download_button(
                "⬇ Download Excel",
                output.getvalue(),
                "defaulters.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            pdf = build_pdf_report(
                filtered,
                {
                    "students": students,
                    "paid": paid_students,
                    "defaulters": total_defaulters,
                    "unpaid_days": unpaid_days,
                },
            )
            st.download_button("⬇ Download PDF", pdf, "Yoga_Fee_Report.pdf", "application/pdf")
        else:
            st.info("No defaulters found matching the criteria! All students have paid.")
            
    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.info("Please ensure the Excel files match the expected format.")

else:
    st.info("Please upload both the Fee Report and Attendance Report to begin.")
