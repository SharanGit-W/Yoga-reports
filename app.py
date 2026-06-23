import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io
from fpdf import FPDF

# ==========================================
# PAGE CONFIG & CUSTOM CSS (From your version)
# ==========================================
st.set_page_config(page_title="Yoga Kendra Fee Defaulter Tracker", layout="wide", page_icon="🧘‍♂️")

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #2C3E50;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #7F8C8D;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        border-left: 5px solid #3498DB;
    }
    .stButton>button {
        background-color: #2C3E50;
        color: white;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: 600;
        border: none;
    }
    .stButton>button:hover {
        background-color: #34495E;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# HELPER FUNCTIONS
# ==========================================
class PDFReport(FPDF):
    def __init__(self, org_name, center_name, report_month):
        super().__init__()
        self.org_name = org_name
        self.center_name = center_name
        self.report_month = report_month

    def header(self):
        self.set_font('Helvetica', 'B', 16)
        self.cell(0, 10, str(self.org_name), 0, 1, 'C')
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 10, str(self.center_name), 0, 1, 'C')
        self.set_font('Helvetica', '', 12)
        self.cell(0, 10, f'Fee Defaulters Report ({self.report_month})', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def load_data(file):
    """Loads CSV or Excel safely, skipping 3 header rows"""
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, skiprows=3)
        df_meta = pd.read_csv(file, nrows=2, header=None)
    else:
        df = pd.read_excel(file, skiprows=3)
        df_meta = pd.read_excel(file, nrows=2, header=None)
    
    org_name = str(df_meta.iloc[0, 0]).strip() if not df_meta.empty else "Organization Name Not Found"
    kendra_name = str(df_meta.iloc[1, 0]).strip() if not df_meta.empty else "Kendra Name Not Found"
    return df, org_name, kendra_name

def extract_months(part_string):
    """Extracts strings like [Jan-26], [Feb-26] into a list"""
    return set(re.findall(r'\[(.*?)\]', str(part_string)))

def generate_pdf(df_def, org, center, report_month):
    pdf = PDFReport(org, center, report_month)
    pdf.add_page()
    
    # Summary Section
    pdf.set_font('Helvetica', '', 10)
    total_defaulters = len(df_def)
    total_unpaid_days = df_def['Unpaid Attended Days'].sum()
    pdf.cell(0, 8, f'Total Defaulters: {total_defaulters}   |   Total Unpaid Attended Days: {total_unpaid_days}', 0, 1, 'L')
    pdf.ln(5)
    
    # Table Headers
    pdf.set_font('Helvetica', 'B', 9)
    headers = ['Student ID', 'Name', 'Batch', 'Timing', 'Days Unpaid']
    col_widths = [25, 60, 45, 40, 20]
    
    pdf.set_fill_color(44, 62, 80)
    pdf.set_text_color(255, 255, 255)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, 1, 0, 'C', True)
    pdf.ln()
    
    # Table Data (Your awesome Zebra-striping design)
    pdf.set_font('Helvetica', '', 8)
    pdf.set_text_color(0, 0, 0)
    
    for idx, row in enumerate(df_def.itertuples()):
        if idx % 2 == 0:
            pdf.set_fill_color(245, 245, 245)
        else:
            pdf.set_fill_color(255, 255, 255)
            
        name = str(row.Name)[:30]
        batch = str(row.Batch)[:25]
        timing = str(row.Timing)[:20]
        
        pdf.cell(col_widths[0], 7, str(row.Student_ID), 1, 0, 'C', True)
        pdf.cell(col_widths[1], 7, name, 1, 0, 'L', True)
        pdf.cell(col_widths[2], 7, batch, 1, 0, 'C', True)
        pdf.cell(col_widths[3], 7, timing, 1, 0, 'C', True)
        pdf.cell(col_widths[4], 7, str(row.Unpaid_Attended_Days), 1, 1, 'C', True)
        
    return pdf.output()

# ==========================================
# STREAMLIT UI
# ==========================================
st.markdown('<p class="main-header">Yoga Kendra Fee Defaulter Tracker</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload your Attendance and Fee reports to identify Ashtanga Yoga students attending without payment.</p>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    att_file = st.file_uploader("Upload Attendance Report (.csv / .xlsx)", type=['csv', 'xlsx'])
with col2:
    fee_file = st.file_uploader("Upload Fee Report (.csv / .xlsx)", type=['csv', 'xlsx'])

if st.button("Generate Defaulter Report", use_container_width=True):
    if att_file and fee_file:
        with st.spinner("Processing data and calculating discrepancies..."):
            try:
                # 1. Load Data
                df_fee, org_name, center_name = load_data(fee_file)
                df_att, _, _ = load_data(att_file)

                # 2. Process Fee Report
                df_fee = df_fee[df_fee['Status'].astype(str).str.strip().str.lower() == 'active']
                df_fee = df_fee[df_fee['Particulars'].astype(str).str.contains('Ashtanga Yoga', case=False, na=False)]
                
                df_fee['Clean_ID'] = df_fee['Stud ID'].astype(str).str.extract(r'(\d+)$')
                df_fee = df_fee.dropna(subset=['Clean_ID'])
                
                df_fee['Paid_Months'] = df_fee['Particulars'].apply(extract_months)
                fee_agg = df_fee.groupby('Clean_ID').agg({
                    'Paid_Months': lambda x: set().union(*x),
                    'Timing': 'first'
                }).reset_index()

                # 3. Process Attendance Report
                df_att = df_att[df_att['Status'].astype(str).str.strip().str.lower() == 'active']
                
                # Filter out Bharatanatyam and other non-yoga batches
                yoga_mask = df_att['Batch'].astype(str).str.contains(r'General|Yoga|Junior|Y G', case=False, na=False) | \
                            df_att['StudentId'].astype(str).str.contains(r'/Y/', na=False)
                df_att = df_att[yoga_mask]
                
                df_att['Clean_ID'] = df_att['StudentId'].astype(str).str.extract(r'(\d+)$')
                df_att = df_att.dropna(subset=['Clean_ID'])
                
                # Robust date column detection (Handles 2026-05-01 formatting safely)
                date_cols = [col for col in df_att.columns if re.match(r'\d{4}-\d{2}-\d{2}', str(col)) or re.match(r'\d{1,2}/\d{1,2}/\d{2,4}', str(col))]
                
                if not date_cols:
                    st.error("Could not detect date columns. Ensure the file contains dates like YYYY-MM-DD or MM/DD/YY.")
                    st.stop()
                
                # Get the report month dynamically
                first_col_str = str(date_cols[0]).split(' ')[0]
                try:
                    first_date = pd.to_datetime(first_col_str)
                    report_month_str = first_date.strftime("%b-%y")
                except:
                    report_month_str = "Current Month"

                # Calculate attended days (Column-wise safely)
                present_mask = df_att[date_cols].apply(lambda col: col.astype(str).str.contains('Present', case=False, na=False))
                df_att['Unpaid Attended Days'] = present_mask.sum(axis=1)
                
                # Apply the rule: More than 2 classes
                df_att = df_att[df_att['Unpaid Attended Days'] > 2]

                # 4. Merge and Find Defaulters
                merged = pd.merge(df_att, fee_agg, on='Clean_ID', how='left')
                
                def check_defaulter(row):
                    paid_set = row['Paid_Months']
                    if isinstance(paid_set, set):
                        return report_month_str not in paid_set
                    return True # True if no payment record at all

                merged['Is_Defaulter'] = merged.apply(check_defaulter, axis=1)
                df_defaulters = merged[merged['Is_Defaulter'] == True]
                
                # Format final dataframe
                if not df_defaulters.empty:
                    df_defaulters = df_defaulters[['Clean_ID', 'StudentName', 'Batch', 'Timing', 'Unpaid Attended Days']].copy()
                    df_defaulters.columns = ['Student_ID', 'Name', 'Batch', 'Timing', 'Unpaid_Attended_Days']
                    df_defaulters['Timing'] = df_defaulters['Timing'].fillna('Not mapped in Fee Report')
                    df_defaulters = df_defaulters.sort_values(by='Unpaid_Attended_Days', ascending=False)
                else:
                    df_defaulters = pd.DataFrame(columns=['Student_ID', 'Name', 'Batch', 'Timing', 'Unpaid_Attended_Days'])

                # 5. Display Results (Using your styled UI)
                st.markdown("---")
                
                if not df_defaulters.empty:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total Defaulters", len(df_defaulters))
                    m2.metric("Total Unpaid Days", df_defaulters['Unpaid_Attended_Days'].sum())
                    m3.metric("Center", center_name)
                    
                    st.dataframe(
                        df_defaulters, 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "Student_ID": st.column_config.TextColumn("Student ID", width="small"),
                            "Name": st.column_config.TextColumn("Name", width="medium"),
                            "Batch": st.column_config.TextColumn("Batch", width="small"),
                            "Timing": st.column_config.TextColumn("Timing", width="medium"),
                            "Unpaid_Attended_Days": st.column_config.NumberColumn("Unpaid Days", width="small")
                        }
                    )
                    
                    # 6. Output Buttons (CSV + PDF)
                    pdf_bytes = generate_pdf(df_defaulters, org_name, center_name, report_month_str)
                    csv_bytes = df_defaulters.to_csv(index=False).encode('utf-8')
                    
                    colA, colB = st.columns(2)
                    with colA:
                        st.download_button(
                            label="📄 Download Professional PDF",
                            data=pdf_bytes,
                            file_name=f"Defaulter_Report_{center_name.replace(' ', '_')}_{report_month_str}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    with colB:
                        st.download_button(
                            label="📊 Download CSV Data",
                            data=csv_bytes,
                            file_name=f"Defaulter_Report_{center_name.replace(' ', '_')}_{report_month_str}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                else:
                    st.success("✅ No defaulters found! All attending students (>2 classes) have paid their fees.")
                    
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
                st.info("Please ensure the files are in the correct format.")
    else:
        st.warning("Please upload both the Attendance and Fee report files to proceed.")
