import re
import io
import pandas as pd
import streamlit as st
import matplotlib
import matplotlib.pyplot as plt

# Force Matplotlib to use Agg backend for thread-safe cloud deployment
matplotlib.use('Agg')

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm

# ==========================================
# 1. DATA INGESTION & PARSING LAYER
# ==========================================

def load_file_smart(file_bytes, target_col, filename):
    """
    Dynamically finds the header row containing the target column.
    Robust against title text above actual headers and fully supports CSV/Excel.
    """
    file_io = io.BytesIO(file_bytes)
    
    try:
        if filename.lower().endswith('.csv'):
            try:
                df = pd.read_csv(file_io, header=None, encoding='utf-8')
            except UnicodeDecodeError:
                file_io.seek(0)
                df = pd.read_csv(file_io, header=None, encoding='ISO-8859-1')
        else:
            try:
                df = pd.read_excel(file_io, header=None, engine="openpyxl")
            except Exception:
                file_io.seek(0)
                df = pd.read_excel(file_io, header=None)
    except pd.errors.EmptyDataError:
        raise ValueError(f"The file '{filename}' appears to be empty.")
    except Exception as e:
        raise ValueError(f"Could not read '{filename}'. Ensure it is a valid format. Error: {str(e)}")
        
    header_idx = -1
    target_clean = "".join(target_col.lower().split())
    
    for i, row in df.iterrows():
        row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
        if (target_col.lower() in row_str or target_clean in "".join(row_str.split())) and "batch" in row_str:
            header_idx = i
            break
            
    if header_idx == -1:
        raise ValueError(f"Mandatory column '{target_col}' not found in '{filename}'. Please check file structure.")
        
    new_columns = [str(x).strip() if pd.notna(x) else f"Unnamed_{j}" for j, x in enumerate(df.iloc[header_idx])]
    df.columns = new_columns
    df = df.iloc[header_idx + 1:].reset_index(drop=True)
    df = df.dropna(how='all')
    return df

def normalize_name(name):
    if pd.isna(name): return ""
    return " ".join(str(name).strip().lower().split())

def normalize_batch(batch):
    if pd.isna(batch): return ""
    batch = str(batch).lower().strip().split(',')[0] 
    batch = batch.replace("batch", "")
    batch = re.sub(r'\s+y[gc]\b', '', batch)
    return " ".join(batch.split())

MONTH_REGEX = r"\[\s*([A-Za-z]{3}\s*[-/\s]\s*\d{2,4})\s*\]"

def parse_months(particulars: str):
    """Extracts target fee service months out of the particulars string."""
    if pd.isna(particulars): return []
    matches = re.findall(MONTH_REGEX, str(particulars), flags=re.IGNORECASE)
    cleaned = []
    for month_str in matches:
        m_clean = month_str.replace(" ", "").replace("/", "-")
        try:
            dt = pd.to_datetime(m_clean, format="%b-%y")
            cleaned.append(dt.strftime("%b-%y"))
        except ValueError:
            try:
                dt = pd.to_datetime(m_clean)
                cleaned.append(dt.strftime("%b-%y"))
            except Exception:
                cleaned.append(m_clean.title())
    return cleaned

@st.cache_data
def load_fee_data(file_bytes, filename):
    fee = load_file_smart(file_bytes, "Stud Name", filename)
    fee.columns = [str(c).strip() for c in fee.columns]
    
    required_columns = ["Stud Name", "Batch", "Status", "Particulars"]
    for col in required_columns:
        if col not in fee.columns:
            raise ValueError(f"Missing mandatory column in Fee Report: '{col}'. Found: {fee.columns.tolist()}")

    fee = fee[~fee["Status"].fillna("").astype(str).str.strip().str.lower().isin(["cancelled", "canceled"])]

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

    paid_df = pd.DataFrame(expanded_rows)
    if not paid_df.empty:
        paid_df = paid_df.drop_duplicates(subset=["Name_Key", "Batch_Key", "Month"])

    return paid_df

@st.cache_data
def load_attendance_data(file_bytes, filename):
    attendance = load_file_smart(file_bytes, "StudentName", filename)
    attendance.columns = [str(c).strip() for c in attendance.columns]
    
    required_columns = ["StudentName", "Batch"]
    for col in required_columns:
        if col not in attendance.columns:
            raise ValueError(f"Missing mandatory column in Attendance Report: '{col}'. Found: {attendance.columns.tolist()}")

    attendance["Name_Key"] = attendance["StudentName"].apply(normalize_name)
    attendance["Batch_Key"] = attendance["Batch"].apply(normalize_batch)

    date_columns = []
    col_to_month_map = {}
    
    # Robust Date Detection (Point 2)
    for col in attendance.columns:
        col_str = str(col).strip()
        if col_str.lower() in ["studentname", "batch", "name_key", "batch_key"] or col_str.startswith("Unnamed"):
            continue
            
        try:
            # Let pandas natively deduce the datetime structure
            dt = pd.to_datetime(col_str, errors='raise')
            date_columns.append(col)
            col_to_month_map[col] = dt.strftime("%b-%y")
        except Exception:
            # Fallback for strict "Jan-25" formats where pandas might object
            if re.match(r"^[A-Za-z]{3}-\d{2,4}$", col_str):
                try:
                    dt = pd.to_datetime(col_str, format="%b-%y")
                    date_columns.append(col)
                    col_to_month_map[col] = dt.strftime("%b-%y")
                except Exception:
                    pass

    attendance_rows = []
    for _, row in attendance.iterrows():
        monthly_days = {}
        for col in date_columns:
            if str(row[col]).strip().lower() in ["present", "p"]:
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

    return pd.DataFrame(attendance_rows)

# ==========================================
# 2. BUSINESS LOGIC LAYER
# ==========================================

@st.cache_data
def generate_business_logic(attendance_df, paid_df):
    """Returns defaulters and unmatched exception records."""
    if attendance_df.empty:
        return pd.DataFrame(columns=["Student Name", "Batch", "Month", "Days Attended"]), pd.DataFrame()
        
    if paid_df.empty:
        return attendance_df.sort_values("Days Attended", ascending=False), pd.DataFrame()

    # 1. Generate Defaulters (Left Join)
    merged = attendance_df.merge(
        paid_df[["Name_Key", "Batch_Key", "Month"]].drop_duplicates(),
        on=["Name_Key", "Batch_Key", "Month"],
        how="left",
        indicator=True
    )
    defaulters = merged[merged["_merge"] == "left_only"].copy()
    defaulters = defaulters[["Student Name", "Batch", "Month", "Days Attended"]].sort_values("Days Attended", ascending=False)
    
    # 2. Identify Exceptions (Payments recorded but no matching attendance)
    paid_merged = paid_df.merge(
        attendance_df[["Name_Key", "Batch_Key", "Month"]].assign(Attended=1).drop_duplicates(),
        on=["Name_Key", "Batch_Key", "Month"],
        how="left"
    )
    unmatched_payments = paid_merged[paid_merged["Attended"].isna()].copy()
    unmatched_payments = unmatched_payments[["Student Name", "Batch", "Month"]]

    return defaulters, unmatched_payments

# ==========================================
# 3. REPORTING & VISUALIZATION LAYER
# ==========================================

@st.cache_data
def create_batch_chart(defaulters):
    buf = io.BytesIO()
    if defaulters.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, "No Defaulters Data", ha="center", va="center")
        plt.savefig(buf, format='png')
        plt.close()
        buf.seek(0)
        return buf.getvalue()

    summary = defaulters.groupby("Batch")["Days Attended"].sum().sort_values(ascending=False).head(15)
    fig, ax = plt.subplots(figsize=(8, 4))
    summary.plot(kind="bar", ax=ax, color='tomato')
    ax.set_title("Unpaid Attendance Days By Batch (Top 15)")
    ax.set_ylabel("Days Attended")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf.getvalue()

@st.cache_data
def build_pdf_report(defaulters, metrics):
    """Generates PDF using Table Chunking to prevent Flowable Scaling/Memory issues."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Yoga Fee Compliance Report", styles["Title"]))
    story.append(Spacer(1, 10))

    kpi_data = [
        ["Metric", "Value"],
        ["Total Batch Enrollments", metrics["students"]],
        ["Paid Enrollments", metrics["paid"]],
        ["Default Enrollments", metrics["defaulters"]],
        ["Total Unpaid Days", metrics["unpaid_days"]],
    ]
    kpi_table = Table(kpi_data, colWidths=[200, 100])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ]))
    story.append(kpi_table)
    story.append(PageBreak())

    # Add Chart
    chart_bytes = create_batch_chart(defaulters)
    story.append(Paragraph("Batch Default Summary", styles["Heading1"]))
    story.append(Image(io.BytesIO(chart_bytes), width=170 * mm, height=90 * mm))
    story.append(PageBreak())

    # Scalable Detailed Reporting - Chunking Logic
    story.append(Paragraph("Detailed Defaulters", styles["Heading1"]))
    
    if defaulters.empty:
        story.append(Paragraph("No defaulters found.", styles["Normal"]))
    else:
        MAX_ROWS = 40 # Pagination size limit
        for i in range(0, len(defaulters), MAX_ROWS):
            chunk = defaulters.iloc[i:i+MAX_ROWS]
            data = [["Student Name", "Batch", "Month", "Days"]]
            for _, row in chunk.iterrows():
                data.append([str(row["Student Name"]), str(row["Batch"]), str(row["Month"]), int(row["Days Attended"])])
            
            chunk_table = Table(data, repeatRows=1)
            chunk_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.darkred),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ]))
            story.append(chunk_table)
            
            if i + MAX_ROWS < len(defaulters):
                story.append(PageBreak())

    doc.build(story)
    return buffer.getvalue()

# ==========================================
# 4. MAIN APPLICATION UI
# ==========================================

st.set_page_config(page_title="Yoga Fee Dashboard", layout="wide")
st.title("🧘 Yoga Fee Compliance Dashboard")

with st.sidebar:
    st.header("Upload Data")
    st.info("ℹ️ Note: The system assumes the month referenced in the fee 'Particulars' is the official Service Month.")
    fee_file = st.file_uploader("1. Fee Report (Excel/CSV)", type=["xlsx", "xls", "csv"])
    attendance_file = st.file_uploader("2. Attendance Report (Excel/CSV)", type=["xlsx", "xls", "csv"])

if fee_file and attendance_file:
    try:
        with st.spinner("Parsing files & evaluating business rules..."):
            paid_df = load_fee_data(fee_file.getvalue(), fee_file.name)
            attendance_df = load_attendance_data(attendance_file.getvalue(), attendance_file.name)
            defaulters, exceptions = generate_business_logic(attendance_df, paid_df)

        # KPIs at Student-Batch Combination Level (Point 5)
        def create_combo(df):
            if df.empty: return []
            return df["Student Name"].astype(str) + "_" + df["Batch"].astype(str)

        total_enrolls = len(set(create_combo(attendance_df))) if not attendance_df.empty else 0
        paid_enrolls = len(set(create_combo(paid_df))) if not paid_df.empty else 0
        default_enrolls = len(set(create_combo(defaulters))) if not defaulters.empty else 0
        unpaid_days = int(defaulters["Days Attended"].sum()) if not defaulters.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Batch Enrollments", total_enrolls)
        c2.metric("Paid Enrollments", paid_enrolls)
        c3.metric("Defaulter Enrollments", default_enrolls, delta_color="inverse")
        c4.metric("Unpaid Attendance Days", unpaid_days, delta_color="inverse")

        st.divider()

        # Implementation of Point 9 & 8 (Summary Dashboards & Exceptions)
        tab1, tab2, tab3 = st.tabs(["📋 Detailed Defaulters", "📊 Summary Dashboards", "⚠️ Data Exceptions"])

        with tab1:
            st.subheader("Students Attending Without Matching Payment")
            search = st.text_input("Search by Student Name (Defaulters):")
            
            filtered = defaulters.copy()
            if not filtered.empty:
                m_filter = st.multiselect("Filter by Month", sorted(filtered["Month"].unique()))
                b_filter = st.multiselect("Filter by Batch", sorted(filtered["Batch"].unique()))
                
                if search:
                    filtered = filtered[filtered["Student Name"].str.contains(search, case=False, na=False)]
                if m_filter:
                    filtered = filtered[filtered["Month"].isin(m_filter)]
                if b_filter:
                    filtered = filtered[filtered["Batch"].isin(b_filter)]

            st.dataframe(filtered, use_container_width=True, height=400)

            if not filtered.empty:
                col1, col2, col3 = st.columns([1,1,2])
                
                # CSV Download
                csv_data = filtered.to_csv(index=False).encode("utf-8")
                col1.download_button("⬇ Download CSV", csv_data, "defaulters.csv", "text/csv")

                # Scalable PDF Download (Point 3 & 4)
                pdf_bytes = build_pdf_report(filtered, {
                    "students": total_enrolls, "paid": paid_enrolls,
                    "defaulters": default_enrolls, "unpaid_days": unpaid_days
                })
                col2.download_button("⬇ Download PDF Report", pdf_bytes, "Compliance_Report.pdf", "application/pdf")
            else:
                st.success("No defaulters found! 100% Fee Compliance.")

        with tab2:
            st.subheader("Operational Overview")
            if not defaulters.empty:
                dash_col1, dash_col2 = st.columns(2)
                
                with dash_col1:
                    st.write("**Batch-Wise Compliance & Defaulters**")
                    att_batch = attendance_df.groupby("Batch")["Student Name"].nunique().rename("Attended Students")
                    def_batch = defaulters.groupby("Batch")["Student Name"].nunique().rename("Defaulters")
                    batch_summary = pd.concat([att_batch, def_batch], axis=1).fillna(0).astype(int)
                    batch_summary["Compliance %"] = ((batch_summary["Attended Students"] - batch_summary["Defaulters"]) / batch_summary["Attended Students"] * 100).round(1).astype(str) + "%"
                    st.dataframe(batch_summary, use_container_width=True)

                with dash_col2:
                    st.write("**Top 10 Defaulters (By Days Attended)**")
                    top_defs = defaulters.groupby(["Student Name", "Batch"])["Days Attended"].sum().reset_index()
                    st.dataframe(top_defs.sort_values("Days Attended", ascending=False).head(10), use_container_width=True)
                    
                st.write("**Month-Wise Unpaid Summary**")
                month_sum = defaulters.groupby("Month").agg(Defaulters=("Student Name", "nunique"), Total_Unpaid_Days=("Days Attended", "sum")).reset_index()
                st.dataframe(month_sum, use_container_width=True)
            else:
                st.info("Insufficient data to generate dashboards.")

        with tab3:
            st.subheader("Unlinked / Unmatched Payments")
            st.write("The following records represent payments logged in the Fee system where the student/batch did not appear in the Attendance system for that specific month. Please check for spelling variations or missing attendance registers.")
            if not exceptions.empty:
                st.dataframe(exceptions, use_container_width=True)
                ex_csv = exceptions.to_csv(index=False).encode('utf-8')
                st.download_button("⬇ Download Exceptions Report", ex_csv, "unmatched_payments.csv", "text/csv")
            else:
                st.success("Perfect Matching! No exceptions identified.")

    # Specific Exception Handling (Point 7)
    except ValueError as ve:
        st.error(f"Validation Error: {str(ve)}")
    except Exception as e:
        st.error(f"An unexpected critical error occurred: {str(e)}")
        st.info("Please verify the structure of your Excel/CSV files matches standard reporting outputs.")

else:
    st.info("Please upload both the Fee Report and Attendance Report to begin reconciliation.")
