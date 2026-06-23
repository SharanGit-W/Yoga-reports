import streamlit as st
import pandas as pd
import re
import tempfile
from fpdf import FPDF

# --- App Configuration ---
st.set_page_config(page_title="Unpaid Fee Tracker - Enterprise Edition", page_icon="🛡️", layout="wide")

st.title("🛡️ Unpaid Fee Tracker")
st.markdown("Upload the Attendance and Fee reports to instantly generate a comprehensive PDF of unpaid students (Yoga Batches Only, 3+ Days Attended).")

# --- UI: Month and Year Selectors ---
col1, col2 = st.columns(2)
with col1:
    selected_month = st.selectbox(
        "Select Month", 
        ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    )
with col2:
    selected_year = st.selectbox("Select Year", list(range(2024, 2031)), index=2) 

# --- Reference Maps ---
month_map = {
    "January": ("01", "Jan"), "February": ("02", "Feb"), "March": ("03", "Mar"),
    "April": ("04", "Apr"), "May": ("05", "May"), "June": ("06", "Jun"),
    "July": ("07", "Jul"), "August": ("08", "Aug"), "September": ("09", "Sep"),
    "October": ("10", "Oct"), "November": ("11", "Nov"), "December": ("12", "Dec")
}

# Allowed Yoga Batches (Strict Filter)
ALLOWED_YOGA_BATCHES = [
    "general",
    "weekend yoga adults",
    "children yoga"
]

# --- Resilient Helper Functions ---
def safe_str(val, max_len=None):
    if pd.isna(val): return ""
    s = str(val).encode('latin-1', 'replace').decode('latin-1').strip()
    if s.endswith('.0'): s = s[:-2] 
    if max_len and len(s) > max_len:
        return s[:max_len-2] + ".."
    return s

def extract_numeric_id(val):
    if pd.isna(val): return None
    val_str = str(val).strip()
    if '/' in val_str:
        val_str = val_str.split('/')[-1]
    match = re.search(r'(\d+)', val_str)
    return float(match.group(1)) if match else None

def find_header_row(df, keywords):
    for idx, row in df.iterrows():
        row_str = "".join(str(val).lower().replace(" ", "") for val in row.dropna())
        if any(kw in row_str for kw in keywords):
            return idx
    return -1

def get_col(df, possible_names):
    clean_possible = [str(p).strip().lower().replace(" ", "").replace("_", "") for p in possible_names]
    for col in df.columns:
        col_clean = str(col).strip().lower().replace(" ", "").replace("_", "")
        if col_clean in clean_possible:
            return col
    return None

def is_allowed_yoga_batch(batch_val):
    if pd.isna(batch_val): return False
    b_clean = re.sub(r'\s+', ' ', str(batch_val).strip().lower())
    return any(kw in b_clean for kw in ALLOWED_YOGA_BATCHES)

# --- BEAUTIFIED PDF GENERATION ---
def create_pdf(dataframe, month, year, org_name, center_name):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    # Theme Colors
    primary_color = (41, 128, 185) # Professional Blue
    text_color = (50, 50, 50)      # Soft Black
    line_color = (200, 200, 200)   # Light Gray
    
    # 1. Organization Name (Brand Color)
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(*primary_color)
    pdf.cell(0, 10, safe_str(org_name), ln=True, align='C')
    
    # 2. Center Name
    pdf.set_font("Arial", 'B', 12)
    pdf.set_text_color(*text_color)
    pdf.cell(0, 6, safe_str(center_name), ln=True, align='C')
    pdf.ln(4)
    
    # Decorative Divider Line
    pdf.set_draw_color(*primary_color)
    pdf.set_line_width(0.5)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(6)
    
    # 3. Report Title
    pdf.set_font("Arial", 'B', 13)
    pdf.set_text_color(*text_color)
    pdf.cell(0, 8, f"Actionable Unpaid Students Report (Yoga Batches) - {month} {year}", ln=True, align='C')
    
    # 4. Summary & Criteria
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(0, 6, "Criteria: Attended 3 or more classes without fee payment.", ln=True, align='C')
    
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, f"Total Actionable Students: {len(dataframe)}", ln=True, align='C')
    pdf.ln(6)
    
    # 5. Table Layout Setup
    w_id = 30
    w_name = 50
    w_batch = 40
    w_time = 50
    w_days = 20
    
    # Table Header (Solid Blue Background, White Text)
    pdf.set_fill_color(*primary_color)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 9)
    pdf.set_draw_color(*line_color)
    pdf.set_line_width(0.2)
    
    pdf.cell(w_id, 10, "Student ID", border=1, align='C', fill=True)
    pdf.cell(w_name, 10, "Student Name", border=1, align='C', fill=True)
    pdf.cell(w_batch, 10, "Batch", border=1, align='C', fill=True)
    pdf.cell(w_time, 10, "Timing", border=1, align='C', fill=True)
    pdf.cell(w_days, 10, "Days Att.", border=1, ln=True, align='C', fill=True)
    
    # 6. Table Data (Zebra Striping for readability)
    pdf.set_font("Arial", '', 9)
    pdf.set_text_color(*text_color)
    
    fill = False # Toggle for alternating row colors
    for _, row in dataframe.iterrows():
        # Set zebra stripe color (light gray vs white)
        if fill:
            pdf.set_fill_color(245, 245, 245)
        else:
            pdf.set_fill_color(255, 255, 255)
            
        pdf.cell(w_id, 8, safe_str(row.get('StudentId', ''), 15), border=1, align='C', fill=True)
        pdf.cell(w_name, 8, safe_str(row.get('StudentName', ''), 25), border=1, align='L', fill=True)
        pdf.cell(w_batch, 8, safe_str(row.get('Batch', ''), 18), border=1, align='C', fill=True)
        pdf.cell(w_time, 8, safe_str(row.get('Timing', ''), 30), border=1, align='C', fill=True)
        pdf.cell(w_days, 8, safe_str(row.get('Days_Attended', '')), border=1, ln=True, align='C', fill=True)
        
        fill = not fill # Switch color for next row
        
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf.output(tmp.name)
        return tmp.name

# --- UI: File Uploaders ---
st.info("💡 You can upload either `.xlsx` or `.csv` files.")
attendance_file = st.file_uploader("1. Upload Attendance Report", type=["xlsx", "xls", "csv"])
fee_file = st.file_uploader("2. Upload Fee Report", type=["xlsx", "xls", "csv"])

# --- Core Logic Execution ---
if st.button("Generate Professional Report"):
    if attendance_file and fee_file:
        try:
            with st.spinner("Executing strict data cross-referencing to ensure 100% accuracy..."):
                target_month_num, target_month_short = month_map[selected_month]
                fee_target = f"{target_month_short}-{str(selected_year)[-2:]}"

                # ==========================================
                # 1. PROCESS FEE REPORT
                # ==========================================
                if fee_file.name.endswith('.csv'):
                    fee_raw = pd.read_csv(fee_file, header=None)
                else:
                    fee_raw = pd.read_excel(fee_file, header=None)
                
                fee_header_idx = find_header_row(fee_raw, ["studid", "particulars"])
                if fee_header_idx == -1:
                    st.error("🚨 Error: Could not locate 'Stud ID' or 'Particulars' in the Fee Report.")
                    st.stop()
                
                fee_df = fee_raw.copy()
                fee_df.columns = fee_df.iloc[fee_header_idx]
                fee_df = fee_df.iloc[fee_header_idx+1:].reset_index(drop=True)
                
                fee_id_col = get_col(fee_df, ["studid", "studentid"])
                fee_part_col = get_col(fee_df, ["particulars", "particular"])
                fee_timing_col = get_col(fee_df, ["timing", "timings"])
                
                if not fee_id_col or not fee_part_col:
                    st.error("🚨 Error: Missing required columns in the Fee Report.")
                    st.stop()
                
                fee_df = fee_df.dropna(subset=[fee_part_col, fee_id_col])
                
                timing_map = {}
                if fee_timing_col:
                    valid_timings = fee_df.dropna(subset=[fee_id_col, fee_timing_col]).copy()
                    valid_timings['Clean_ID'] = valid_timings[fee_id_col].apply(extract_numeric_id)
                    valid_timings = valid_timings.dropna(subset=['Clean_ID']).drop_duplicates(subset=['Clean_ID'], keep='last')
                    timing_map = valid_timings.set_index('Clean_ID')[fee_timing_col].to_dict()

                paid_df = fee_df[fee_df[fee_part_col].astype(str).str.contains(fee_target, case=False, na=False)]
                paid_ids = set(paid_df[fee_id_col].apply(extract_numeric_id).dropna().unique())

                # ==========================================
                # 2. PROCESS ATTENDANCE
                # ==========================================
                if attendance_file.name.endswith('.csv'):
                    att_raw = pd.read_csv(attendance_file, header=None)
                else:
                    att_raw = pd.read_excel(attendance_file, header=None)
                
                raw_org = att_raw.iloc[0, 0] if len(att_raw) > 0 else None
                raw_center = att_raw.iloc[1, 0] if len(att_raw) > 1 else None
                org_name_val = str(raw_org).strip() if pd.notna(raw_org) else "Organization Name Not Found"
                center_name_val = str(raw_center).strip() if pd.notna(raw_center) else "Center Name Not Found"
                
                att_header_idx = find_header_row(att_raw, ["studentid", "studentname"])
                if att_header_idx == -1:
                    st.error("🚨 Error: Could not locate 'StudentId' or 'StudentName' in the Attendance Report.")
                    st.stop()
                
                att_df = att_raw.copy()
                att_df.columns = att_df.iloc[att_header_idx]
                att_df = att_df.iloc[att_header_idx+1:].reset_index(drop=True)
                
                att_id_col = get_col(att_df, ["studentid", "studid"])
                att_name_col = get_col(att_df, ["studentname", "name"])
                att_batch_col = get_col(att_df, ["batch", "batches"])
                att_present_col = get_col(att_df, ["present", "totalpresent", "dayspresent"])
                
                if not att_id_col or not att_name_col or not att_present_col:
                    st.error("🚨 Error: Missing required columns in the Attendance Report. Ensure the 'Present' column exists.")
                    st.stop()
                
                # Filter for only the allowed Yoga Batches
                att_df = att_df[att_df[att_batch_col].apply(is_allowed_yoga_batch)].copy()
                
                # Extract days strictly from the Present column
                att_df['Days_Attended'] = pd.to_numeric(att_df[att_present_col], errors='coerce').fillna(0)
                
                # STRICT FILTER: Only keep students with 3 or more days attended
                attended_df = att_df[att_df['Days_Attended'] >= 3].copy()
                
                attended_df['Clean_ID'] = attended_df[att_id_col].apply(extract_numeric_id)
                attended_df = attended_df.dropna(subset=['Clean_ID'])
                
                # ==========================================
                # 3. CROSS-REFERENCE & OUTPUT
                # ==========================================
                unpaid_df = attended_df[~attended_df['Clean_ID'].isin(paid_ids)].copy()
                
                unpaid_df['StudentId'] = unpaid_df[att_id_col]
                unpaid_df['StudentName'] = unpaid_df[att_name_col]
                unpaid_df['Batch'] = unpaid_df[att_batch_col] if att_batch_col else "N/A"
                unpaid_df['Timing'] = unpaid_df['Clean_ID'].map(timing_map).fillna("Timing Not Found")
                
                unpaid_df = unpaid_df[['StudentId', 'StudentName', 'Batch', 'Timing', 'Days_Attended']].reset_index(drop=True)
                
                st.subheader("Report Output")
                st.write(f"**Organization:** {org_name_val}")
                st.write(f"**Center:** {center_name_val}")
                
                if unpaid_df.empty:
                    st.success(f"🎉 100% Match! All Yoga students (3+ days) in {selected_month} {selected_year} have paid.")
                else:
                    st.warning(f"Found {len(unpaid_df)} actionable Yoga students who have attended 3+ days without paying.")
                    st.dataframe(unpaid_df, use_container_width=True)
                    
                    pdf_path = create_pdf(unpaid_df, selected_month, selected_year, org_name_val, center_name_val)
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                    
                    safe_filename_center = re.sub(r'[^A-Za-z0-9_-]', '_', center_name_val)
                    st.download_button(
                        label="📥 Download Certified PDF Report",
                        data=pdf_bytes,
                        file_name=f"Yoga_Unpaid_Report_{safe_filename_center}_{selected_month}_{selected_year}.pdf",
                        mime="application/pdf",
                        type="primary"
                    )
        except Exception as e:
            st.error(f"🚨 A system error occurred while processing: {e}. Please check the files and try again.")
    else:
        st.info("⚠️ Please upload both the Attendance and Fee reports to begin.")
