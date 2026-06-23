import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io
from fpdf import FPDF

# --- STREAMLIT PAGE CONFIG ---
st.set_page_config(page_title="Yoga Fee Defaulters Report", page_icon="🧘‍♂️", layout="wide")

# Customizing Streamlit UI a bit to make it look modern
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stButton>button {background-color: #004d40; color: white; border-radius: 5px; width: 100%;}
    h1, h2, h3 {color: #004d40;}
    .report-card {background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);}
    </style>
""", unsafe_allow_html=True)

# --- PDF GENERATOR CLASS ---
class PDFReport(FPDF):
    def __init__(self, org_name, kendra_name, report_month):
        super().__init__()
        self.org_name = org_name
        self.kendra_name = kendra_name
        self.report_month = report_month

    def header(self):
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 8, str(self.org_name), border=0, align='C', new_x="LMARGIN", new_y="NEXT")
        self.set_font('Helvetica', 'B', 12)
        self.cell(0, 8, str(self.kendra_name), border=0, align='C', new_x="LMARGIN", new_y="NEXT")
        self.set_font('Helvetica', 'I', 11)
        self.cell(0, 8, f'Defaulters Report - {self.report_month}', border=0, align='C', new_x="LMARGIN", new_y="NEXT")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', border=0, align='C')

# --- HELPER FUNCTIONS ---
def load_data(file):
    """Loads CSV or Excel safely, skipping 3 header rows for tabular data"""
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, skiprows=3)
        # Read the top two lines for metadata
        df_meta = pd.read_csv(file, nrows=2, header=None)
    else:
        df = pd.read_excel(file, skiprows=3)
        df_meta = pd.read_excel(file, nrows=2, header=None)
    
    org_name = str(df_meta.iloc[0, 0]).strip()
    kendra_name = str(df_meta.iloc[1, 0]).strip()
    return df, org_name, kendra_name

def extract_months(part_string):
    """Extracts strings like [Jan-26], [Feb-26] into a list"""
    return re.findall(r'\[(.*?)\]', str(part_string))

def process_defaulters(fee_file, att_file):
    # 1. Load Data
    df_fee, org_name, kendra_name = load_data(fee_file)
    df_att, _, _ = load_data(att_file)

    # 2. Process Fee Data
    # Filter only Active and Ashtanga Yoga records
    df_fee = df_fee[df_fee['Status'].astype(str).str.strip().str.lower() == 'active']
    df_fee = df_fee[df_fee['Particulars'].astype(str).str.contains('Ashtanga Yoga', case=False, na=False)]
    
    # Extract numerical Stud ID robustly
    df_fee['Clean_ID'] = df_fee['Stud ID'].astype(str).str.extract(r'(\d+)$')
    
    # Aggregate paid months for duplicates/monthly payers
    df_fee['Paid_Months'] = df_fee['Particulars'].apply(extract_months)
    fee_agg = df_fee.groupby('Clean_ID').agg({
        'Paid_Months': lambda x: set([item for sublist in x for item in sublist]),
        'Timing': 'first' # Take the timing they are registered for
    }).reset_index()

    # 3. Process Attendance Data
    # Filter Active students
    df_att = df_att[df_att['Status'].astype(str).str.strip().str.lower() == 'active']
    
    # Extract numerical ID from composite strings like 7/KMP/Y/1182
    df_att['Clean_ID'] = df_att['StudentId'].astype(str).str.extract(r'(\d+)$')
    
    # Identify the date columns and determine the report month (e.g., "May-26")
    date_cols = [col for col in df_att.columns if re.match(r'\d{4}-\d{2}-\d{2}', str(col))]
    if not date_cols:
        st.error("Could not find standard date columns in the attendance file.")
        return None, None, None, None
        
    first_date = datetime.strptime(date_cols[0], "%Y-%m-%d")
    report_month_str = first_date.strftime("%b-%y") # Maps 2026-05-01 -> "May-26"
    
    # Calculate Total Present Days
    df_att['Days_Attended'] = df_att[date_cols].apply(
        lambda row: row.astype(str).str.contains('Present', case=False, na=False).sum(), axis=1
    )
    
    # Filter for > 2 classes
    df_att = df_att[df_att['Days_Attended'] > 2]

    # 4. Merge Data and Find Defaulters
    merged = pd.merge(df_att, fee_agg, on='Clean_ID', how='left')
    
    def check_defaulter(row):
        paid_set = row['Paid_Months']
        if isinstance(paid_set, set):
            return report_month_str not in paid_set
        return True # If NaN (never paid anything for Ashtanga), they are a defaulter

    merged['Is_Defaulter'] = merged.apply(check_defaulter, axis=1)
    defaulters = merged[merged['Is_Defaulter'] == True]
    
    # Format the final output table
    final_cols = ['Clean_ID', 'StudentName', 'Batch', 'Timing', 'Days_Attended']
    defaulters_display = defaulters[final_cols].copy()
    defaulters_display.columns = ['Student ID', 'Student Name', 'Batch', 'Timing', 'Days Attended (Unpaid)']
    defaulters_display['Timing'] = defaulters_display['Timing'].fillna('Not mapped in Fee Report')
    
    return defaulters_display, org_name, kendra_name, report_month_str

def generate_pdf(df, org_name, kendra_name, report_month_str):
    pdf = PDFReport(org_name, kendra_name, report_month_str)
    pdf.add_page()
    
    # Table Header
    pdf.set_font("Helvetica", "B", 10)
    col_widths = [25, 55, 45, 40, 25]
    headers = ["Stud ID", "Name", "Batch", "Timing", "Days Unpaid"]
    
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, border=1, align='C')
    pdf.ln()

    # Table Body
    pdf.set_font("Helvetica", "", 9)
    for _, row in df.iterrows():
        # Truncate strings to prevent PDF wrapping overflow
        s_id = str(row['Student ID'])
        s_name = str(row['Student Name'])[:25] 
        s_batch = str(row['Batch'])[:22]
        s_timing = str(row['Timing'])[:18]
        s_days = str(row['Days Attended (Unpaid)'])
        
        pdf.cell(col_widths[0], 10, s_id, border=1, align='C')
        pdf.cell(col_widths[1], 10, s_name, border=1, align='L')
        pdf.cell(col_widths[2], 10, s_batch, border=1, align='C')
        pdf.cell(col_widths[3], 10, s_timing, border=1, align='C')
        pdf.cell(col_widths[4], 10, s_days, border=1, align='C')
        pdf.ln()
        
    return pdf.output()

# --- MAIN APP UI ---
st.title("🧘‍♂️ RYSRI Fee Defaulters Tracker")
st.markdown("Upload the Fee Report and Attendance Report to cross-reference Ashtanga Yoga students who have attended **more than 2 classes** without paying for the current month.")

col1, col2 = st.columns(2)
with col1:
    fee_upload = st.file_uploader("Upload Fee Report (.csv / .xlsx)", type=['csv', 'xlsx'])
with col2:
    att_upload = st.file_uploader("Upload Attendance Report (.csv / .xlsx)", type=['csv', 'xlsx'])

if st.button("Generate Defaulters Report", type="primary"):
    if fee_upload and att_upload:
        with st.spinner("Processing Data and Cross-referencing IDs..."):
            try:
                result_df, org_name, kendra_name, report_month = process_defaulters(fee_upload, att_upload)
                
                if result_df is not None:
                    st.success(f"Report Generated Successfully for {kendra_name} ({report_month})!")
                    
                    st.markdown("<div class='report-card'>", unsafe_allow_html=True)
                    st.subheader(f"📊 Defaulters Found: {len(result_df)}")
                    st.dataframe(result_df, use_container_width=True, hide_index=True)
                    
                    # Generate and allow PDF download
                    pdf_bytes = generate_pdf(result_df, org_name, kendra_name, report_month)
                    
                    colA, colB, colC = st.columns([1, 2, 1])
                    with colB:
                        st.download_button(
                            label="📄 Download Professional PDF Report",
                            data=pdf_bytes,
                            file_name=f"Defaulters_Report_{kendra_name.replace(' ', '_')}_{report_month}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    st.markdown("</div>", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"An error occurred while processing the files. Ensure the file formats match the expected layout. Error Details: {e}")
    else:
        st.warning("⚠️ Please upload both the Fee Report and the Attendance Report before processing.")
