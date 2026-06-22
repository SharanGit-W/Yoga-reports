import streamlit as st
import pandas as pd
import datetime
import re
from fpdf import FPDF

# --- SYSTEM CONFIGURATION ---
st.set_page_config(page_title="Automated Fee & Attendance Reconciliation", layout="wide")

st.title("Enterprise Attendance & Fee Reconciliation Portal")
st.markdown("""
This production portal automates the cross-referencing of monthly attendance sheets against centralized fee collection reports. 
It executes a strict data-matching routine to isolate records where attendance has been recorded without corresponding financial clearance.
""")

# --- ROBUST UTILITY OPERATIONS ---
def load_data(uploaded_file, skip_rows):
    """Safely ingests CSV or Excel file formats based on file extension."""
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
    Extracts the core numeric user ID, stripping away center prefixes, batch codes, 
    nested directory slashes, and floating point conversions.
    """
    if pd.isna(raw_id):
        return ""
        
    raw_str = str(raw_id).strip()
    if raw_str.lower() == 'nan' or raw_str == '':
        return ""
        
    # Rectify float truncation bug caused by trailing null rows in Excel sheets
    if re.search(r'\.0+$', raw_str):
        raw_str = re.sub(r'\.0+$', '', raw_str)
        
    match = re.search(r'(\d+)$', raw_str)
    return match.group(1) if match else raw_str

def get_dynamic_column(df, keywords):
    """Dynamically resolves column names to mitigate human error in naming schema across branches."""
    for col in df.columns:
        col_clean = re.sub(r'[^a-zA-Z]', '', str(col)).lower()
        if any(kw in col_clean for kw in keywords):
            return col
    return None

def generate_pdf_report(df_unpaid, center_name):
    """Compiles a professionally customized, branded PDF document for branch visibility."""
    pdf = FPDF()
    pdf.add_page()
    
    # Clean up center name for display safety
    display_center = str(center_name).encode('latin-1', 'replace').decode('latin-1')
    
    # Professional Header Banner
    pdf.set_fill_color(31, 78, 121) # Corporate Navy
    pdf.rect(0, 0, 210, 40, 'F')
    
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(190, 15, txt="RASHTROTTHANA YOGIC SCIENCES & RESEARCH INSTITUTE", ln=True, align='L')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 5, txt=f"EXCEPTION REPORT: {display_center.upper()}", ln=True, align='L')
    pdf.ln(15)
    
    # Reset text color for body
    pdf.set_text_color(0, 0, 0)
    
    # Summary Dashboard Block
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(190, 8, txt="Executive Summary", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(2)
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(95, 6, txt=f"Generated On: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=False)
    pdf.cell(95, 6, txt=f"Total Policy Violations Detected: {len(df_unpaid)} accounts", ln=True)
    pdf.cell(95, 6, txt=f"Total Unpaid Service Days: {df_unpaid['Unpaid Days Attended'].sum()} days", ln=True)
    pdf.ln(8)

    # Table Grid Layout Definition
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(240, 240, 240)
    cols = ['Core ID', 'Name', 'Months Unpaid', 'Days Attended']
    col_widths = [25, 65, 60, 40]
    
    for i, col in enumerate(cols):
        pdf.cell(col_widths[i], 10, col, border=1, align='C', fill=True)
    pdf.ln()

    # Data Row Populating with Latin-1 Encoding Sanitation
    pdf.set_font("Arial", '', 9)
    for _, row in df_unpaid.iterrows():
        stud_id = str(row['Core ID'])[:15].encode('latin-1', 'replace').decode('latin-1')
        name = str(row['Name'])[:30].encode('latin-1', 'replace').decode('latin-1')
        months = str(row['Unpaid Months']).encode('latin-1', 'replace').decode('latin-1')
        days = str(row['Unpaid Days Attended'])

        pdf.cell(col_widths[0], 10, stud_id, border=1, align='C')
        pdf.cell(col_widths[1], 10, name, border=1)
        pdf.cell(col_widths[2], 10, months, border=1)
        pdf.cell(col_widths[3], 10, days, border=1, align='C')
        pdf.ln()

    return pdf.output(dest='S').encode('latin-1')

# --- USER CONTROL INTERFACE ---
with st.sidebar:
    st.header("SOP Extraction Controls")
    skip_rows = st.number_input("Metadata Header Rows to Skip", min_value=0, max_value=10, value=3, 
                                help="Number of title/blank metadata rows above the main column headers.")
    active_status_only = st.checkbox("Filter Fee Ledger by 'Active' status only", value=True)

# --- WORKFLOW FILES INGESTION ---
col1, col2 = st.columns(2)
with col1:
    att_file = st.file_uploader("1. Select Branch Attendance File", type=['csv', 'xlsx'])
with col2:
    fee_file = st.file_uploader("2. Select Branch Fee Ledger File", type=['csv', 'xlsx'])

if att_file and fee_file:
    if st.button("Run Reconciliation Engine", type="primary"):
        try:
            with st.spinner("Executing cross-reconciliation mapping..."):
                # 1. Load Datasets
                df_att = load_data(att_file, skip_rows)
                df_fee = load_data(fee_file, skip_rows)
                
                # 2. Hoist Dynamic Target Columns to Optimize Loop Performance
                att_id_col = get_dynamic_column(df_att, ['studentid', 'studid'])
                fee_id_col = get_dynamic_column(df_fee, ['studentid', 'studid'])
                part_col = get_dynamic_column(df_fee, ['particular']) or 'Particulars'
                fee_status_col = get_dynamic_column(df_fee, ['status']) or 'Status'
                
                att_name_col = get_dynamic_column(df_att, ['name']) or 'StudentName'
                att_center_col = get_dynamic_column(df_att, ['center']) or 'Center'

                if not att_id_col or not fee_id_col:
                    st.error("Data Alignment Error: Failed to identify primary Student ID attributes. Please verify 'Rows to Skip' config.")
                    st.stop()

                # 3. Clean and Unify Joins Mapping
                df_att['Core_ID'] = df_att[att_id_col].apply(extract_core_id)
                df_fee['Core_ID'] = df_fee[fee_id_col].apply(extract_core_id)

                # 4. Enforce Financial System Filters
                if active_status_only and fee_status_col in df_fee.columns:
                    df_fee = df_fee[df_fee[fee_status_col].astype(str).str.strip().str.lower() == 'active']

                # Isolate target timeline arrays (YYYY-MM-DD columns)
                date_cols = [col for col in df_att.columns if re.search(r'\d{4}-\d{2}-\d{2}', str(col))]

                unpaid_records = []

                # 5. Core Matrix Matching Loop
                for index, row in df_att.iterrows():
                    core_id = row['Core_ID']
                    raw_id = row[att_id_col]
                    
                    if not core_id: 
                        continue

                    # Capture any variations of 'Present' marker notations
                    present_dates = [col for col in date_cols if str(row[col]).strip().lower() in ['present', 'p']]
                    
                    if not present_dates:
                        continue 

                    # Build the student's actual attendance footprint
                    months_attended = {}
                    for d in present_dates:
                        date_str = re.search(r'\d{4}-\d{2}-\d{2}', str(d)).group(0)
                        dt = datetime.datetime.strptime(date_str, '%Y-%m-%d')
                        month_str = dt.strftime('%b-%y').lower()
                        months_attended[month_str] = months_attended.get(month_str, 0) + 1

                    # Query the active financial footprint for this student
                    user_fees = df_fee[df_fee['Core_ID'] == core_id]
                    paid_particulars = " ".join(user_fees[part_col].fillna('').astype(str).tolist()).lower()

                    unpaid_days = 0
                    unpaid_months = []

                    # Run cross-ledger validation checks
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

                # Determine the specific center name dynamically from the dataset
                detected_center = "Unknown_Center"
                if att_center_col in df_att.columns and not df_att.empty:
                    valid_centers = df_att[att_center_col].dropna()
                    if not valid_centers.empty:
                        detected_center = str(valid_centers.iloc[0]).strip()
                
                # Create a clean, stylized string for file naming (e.g., "RYSRI_RBI_Layout")
                clean_center_filename = re.sub(r'[^a-zA-Z0-9]', '_', detected_center)
                clean_center_filename = re.sub(r'_+', '_', clean_center_filename).strip('_')

                # 6. User Interface Output Delivery
                if not unpaid_records:
                    st.success(f"Reconciliation Complete for {detected_center}: Zero exceptions found. All attending users possess active financial clearance.")
                else:
                    df_unpaid = pd.DataFrame(unpaid_records)
                    st.error(f"Reconciliation Complete: Isolated {len(df_unpaid)} accounts in {detected_center} violating policy.")
                    
                    st.dataframe(df_unpaid, use_container_width=True)

                    # Prepare customized high-fidelity assets for distribution
                    pdf_bytes = generate_pdf_report(df_unpaid, detected_center)
                    csv_bytes = df_unpaid.to_csv(index=False).encode('utf-8')

                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label=f"📄 Download {detected_center} PDF Report",
                            data=pdf_bytes,
                            file_name=f"Unpaid_Report_{clean_center_filename}_{datetime.date.today()}.pdf",
                            mime="application/pdf"
                        )
                    with col_dl2:
                        st.download_button(
                            label="📊 Download Raw Exception Data (CSV)",
                            data=csv_bytes,
                            file_name=f"Unpaid_Data_{clean_center_filename}_{datetime.date.today()}.csv",
                            mime="text/csv"
                        )

        except Exception as e:
            st.error(f"System execution interdicted. Technical stack trace: {e}")
