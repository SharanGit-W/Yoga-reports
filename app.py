import streamlit as st
import pandas as pd
import datetime
import re
from fpdf import FPDF

# --- CONFIGURATION & SOP SETTINGS ---
st.set_page_config(page_title="Automated Fee & Attendance Reconciliation", layout="wide")

st.title("Attendance vs. Fee Reconciliation Portal")
st.markdown("""
This tool automates the cross-referencing of monthly attendance sheets against fee collection reports. 
It establishes a strict data-matching procedure to identify users who have marked attendance but lack an active, overlapping fee payment.
""")

# --- HELPER FUNCTIONS ---
def load_data(uploaded_file, skip_rows):
    """Loads CSV or Excel data dynamically based on file extension."""
    filename = uploaded_file.name.lower()
    if filename.endswith('.csv'):
        return pd.read_csv(uploaded_file, skiprows=skip_rows)
    elif filename.endswith('.xlsx') or filename.endswith('.xls'):
        return pd.read_excel(uploaded_file, skiprows=skip_rows)
    else:
        st.error(f"Unsupported file format for {uploaded_file.name}. Please upload CSV or Excel.")
        return None

def extract_core_id(raw_id):
    """
    Extracts the core numeric ID, ignoring center/batch prefixes and float conversions. 
    Matches '7/KMP/MD/BR/47257', '8-NGB/PDC//TP10229', '47324', or '47324.0' -> returns string digits.
    """
    if pd.isna(raw_id):
        return ""
        
    raw_str = str(raw_id).strip()
    if raw_str.lower() == 'nan' or raw_str == '':
        return ""
        
    # Handle pandas converting integer columns with empty rows to float64 (e.g., '41243.0')
    if re.search(r'\.0+$', raw_str):
        raw_str = re.sub(r'\.0+$', '', raw_str)
        
    match = re.search(r'(\d+)$', raw_str)
    return match.group(1) if match else raw_str

def get_dynamic_column(df, keywords):
    """Finds a column name dynamically based on a list of keywords."""
    for col in df.columns:
        col_clean = re.sub(r'[^a-zA-Z]', '', str(col)).lower()
        if any(kw in col_clean for kw in keywords):
            return col
    return None

def generate_pdf_report(df_unpaid):
    """Generates a structured PDF document from the unpaid DataFrame."""
    pdf = FPDF()
    pdf.add_page()
    
    # Header
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="Unpaid Attendance Exception Report", ln=True, align='C')
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(200, 10, txt=f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True, align='C')
    pdf.ln(10)

    # Table Header setup
    pdf.set_font("Arial", 'B', 10)
    cols = ['Core ID', 'Name', 'Center', 'Months Unpaid', 'Days Attended']
    col_widths = [20, 50, 45, 45, 30]
    
    for i, col in enumerate(cols):
        pdf.cell(col_widths[i], 10, col, border=1, align='C')
    pdf.ln()

    # Table Body
    pdf.set_font("Arial", '', 9)
    for _, row in df_unpaid.iterrows():
        # Sanitize text to prevent FPDF unicode encoding errors
        stud_id = str(row['Core ID'])[:15].encode('latin-1', 'replace').decode('latin-1')
        name = str(row['Name'])[:25].encode('latin-1', 'replace').decode('latin-1')
        center = str(row['Center'])[:22].encode('latin-1', 'replace').decode('latin-1')
        months = str(row['Unpaid Months']).encode('latin-1', 'replace').decode('latin-1')
        days = str(row['Unpaid Days Attended'])

        pdf.cell(col_widths[0], 10, stud_id, border=1, align='C')
        pdf.cell(col_widths[1], 10, name, border=1)
        pdf.cell(col_widths[2], 10, center, border=1)
        pdf.cell(col_widths[3], 10, months, border=1)
        pdf.cell(col_widths[4], 10, days, border=1, align='C')
        pdf.ln()

    return pdf.output(dest='S').encode('latin-1')

# --- SIDEBAR PARAMETERS ---
with st.sidebar:
    st.header("Extraction Parameters")
    skip_rows = st.number_input("Header Rows to Skip", min_value=0, max_value=10, value=3, 
                                help="Number of title/blank rows before the actual column names start.")
    active_status_only = st.checkbox("Filter Fee Report by 'Active' status only", value=True)

# --- MAIN UI ---
col1, col2 = st.columns(2)
with col1:
    att_file = st.file_uploader("1. Upload Attendance Report", type=['csv', 'xlsx'])
with col2:
    fee_file = st.file_uploader("2. Upload Fee Report", type=['csv', 'xlsx'])

if att_file and fee_file:
    if st.button("Execute Reconciliation SOP", type="primary"):
        try:
            with st.spinner("Processing datasets and cross-referencing records..."):
                # 1. Ingest Data
                df_att = load_data(att_file, skip_rows)
                df_fee = load_data(fee_file, skip_rows)
                
                # 2. Dynamically locate key columns ONCE to optimize processing speed
                att_id_col = get_dynamic_column(df_att, ['studentid', 'studid'])
                fee_id_col = get_dynamic_column(df_fee, ['studentid', 'studid'])
                part_col = get_dynamic_column(df_fee, ['particular']) or 'Particulars'
                fee_status_col = get_dynamic_column(df_fee, ['status']) or 'Status'
                
                att_name_col = get_dynamic_column(df_att, ['name']) or 'StudentName'
                att_center_col = get_dynamic_column(df_att, ['center']) or 'Center'

                if not att_id_col or not fee_id_col:
                    st.error("Critical Error: Could not locate Student ID columns in files. Check skipped rows.")
                    st.stop()

                # 3. Standardize Primary Keys
                df_att['Core_ID'] = df_att[att_id_col].apply(extract_core_id)
                df_fee['Core_ID'] = df_fee[fee_id_col].apply(extract_core_id)

                # 4. Filter active payments
                if active_status_only and fee_status_col in df_fee.columns:
                    df_fee = df_fee[df_fee[fee_status_col].astype(str).str.strip().str.lower() == 'active']

                # Identify strictly date-formatted columns in attendance 
                date_cols = [col for col in df_att.columns if re.search(r'\d{4}-\d{2}-\d{2}', str(col))]

                unpaid_records = []

                # 5. Core matching logic
                for index, row in df_att.iterrows():
                    core_id = row['Core_ID']
                    raw_id = row[att_id_col]
                    
                    if not core_id: 
                        continue

                    # Extract dates marked "Present" or "P"
                    present_dates = [col for col in date_cols if str(row[col]).strip().lower() in ['present', 'p']]
                    
                    if not present_dates:
                        continue 

                    # Map dates to specific month footprints
                    months_attended = {}
                    for d in present_dates:
                        date_str = re.search(r'\d{4}-\d{2}-\d{2}', str(d)).group(0)
                        dt = datetime.datetime.strptime(date_str, '%Y-%m-%d')
                        month_str = dt.strftime('%b-%y').lower() # Yields 'may-26', 'jun-26'
                        months_attended[month_str] = months_attended.get(month_str, 0) + 1

                    # Fetch user's active fee records
                    user_fees = df_fee[df_fee['Core_ID'] == core_id]
                    
                    # Consolidate 'Particulars' string for this user and force lowercase
                    paid_particulars = " ".join(user_fees[part_col].fillna('').astype(str).tolist()).lower()

                    unpaid_days = 0
                    unpaid_months = []

                    # Reconcile attendance footprint against payment footprint
                    for month, days_present in months_attended.items():
                        if month not in paid_particulars:
                            unpaid_days += days_present
                            unpaid_months.append(month.capitalize())

                    if unpaid_days > 0:
                        unpaid_records.append({
                            'Core ID': core_id,
                            'Original ID': raw_id,
                            'Name': row.get(att_name_col, 'Unknown'),
                            'Center': row.get(att_center_col, 'Unknown'),
                            'Unpaid Months': ", ".join(unpaid_months),
                            'Unpaid Days Attended': unpaid_days
                        })

                # 6. Output Generation
                if not unpaid_records:
                    st.success("Reconciliation Complete: All attending students have valid fee records.")
                else:
                    df_unpaid = pd.DataFrame(unpaid_records)
                    st.error(f"Reconciliation Complete: Found {len(df_unpaid)} student(s) attending without payment.")
                    
                    st.dataframe(df_unpaid, use_container_width=True)

                    pdf_bytes = generate_pdf_report(df_unpaid)
                    csv_bytes = df_unpaid.to_csv(index=False).encode('utf-8')

                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label="📄 Download PDF Report",
                            data=pdf_bytes,
                            file_name=f"Unpaid_Attendance_Report_{datetime.date.today()}.pdf",
                            mime="application/pdf"
                        )
                    with col_dl2:
                        st.download_button(
                            label="📊 Download Raw CSV (For Excel)",
                            data=csv_bytes,
                            file_name=f"Unpaid_Attendance_Data_{datetime.date.today()}.csv",
                            mime="text/csv"
                        )

        except Exception as e:
            st.error(f"Execution failed. Error trace: {e}")
