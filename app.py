import streamlit as st
import pandas as pd
import re
import tempfile
from fpdf import FPDF

# --- App Configuration ---
st.set_page_config(page_title="Unpaid Fee Tracker", page_icon="📊", layout="centered")

st.title("📊 Unpaid Fee Tracker")
st.markdown("Upload the Attendance and Fee reports to instantly generate a PDF of unpaid students.")

# --- UI: Month and Year Selectors ---
col1, col2 = st.columns(2)
with col1:
    selected_month = st.selectbox(
        "Select Month", 
        ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    )
with col2:
    # Defaults to 2026 (index 2 of the list)
    selected_year = st.selectbox("Select Year", list(range(2024, 2031)), index=2) 

# --- Logic: Format Dates for Searching ---
month_map = {
    "January": ("01", "Jan"), "February": ("02", "Feb"), "March": ("03", "Mar"),
    "April": ("04", "Apr"), "May": ("05", "May"), "June": ("06", "Jun"),
    "July": ("07", "Jul"), "August": ("08", "Aug"), "September": ("09", "Sep"),
    "October": ("10", "Oct"), "November": ("11", "Nov"), "December": ("12", "Dec")
}

# e.g., '2026-06' and 'Jun-26'
attendance_target = f"{selected_year}-{month_map[selected_month][0]}"
fee_target = f"{month_map[selected_month][1]}-{str(selected_year)[-2:]}"

# --- UI: File Uploaders ---
st.info("💡 You can upload either `.xlsx` or `.csv` files.")
attendance_file = st.file_uploader("1. Upload Attendance Report", type=["xlsx", "xls", "csv"])
fee_file = st.file_uploader("2. Upload Fee Report", type=["xlsx", "xls", "csv"])

# --- PDF Generation Function ---
def create_pdf(dataframe, month, year, org_name, center_name):
    pdf = FPDF()
    pdf.add_page()
    
    # 1. Organization Name
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 8, org_name, ln=True, align='C')
    
    # 2. Center Name
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, center_name, ln=True, align='C')
    pdf.ln(4)
    
    # 3. Report Title
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, f"Unpaid Students Report - {month} {year}", ln=True, align='C')
    
    # 4. Summary
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, f"Total Students Attending but Unpaid: {len(dataframe)}", ln=True, align='C')
    pdf.ln(6)
    
    # 5. Table Header
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(20) # Left margin
    pdf.cell(50, 10, "Student ID", border=1, align='C')
    pdf.cell(100, 10, "Student Name", border=1, ln=True, align='C')
    
    # 6. Table Data
    pdf.set_font("Arial", '', 11)
    for _, row in dataframe.iterrows():
        # Encode strings to prevent FPDF from crashing on special characters
        safe_name = str(row['StudentName']).encode('latin-1', 'replace').decode('latin-1')
        safe_id = str(row['StudentId']).encode('latin-1', 'replace').decode('latin-1')
        
        pdf.cell(20) 
        pdf.cell(50, 10, safe_id, border=1, align='C')
        pdf.cell(100, 10, safe_name, border=1, ln=True, align='L')
        
    # Save to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf.output(tmp.name)
        return tmp.name

# --- Execution Logic ---
if st.button("Generate Report"):
    if attendance_file and fee_file:
        try:
            with st.spinner("Analyzing records..."):
                
                # --- 1. Process Attendance ---
                if attendance_file.name.endswith('.csv'):
                    att_raw = pd.read_csv(attendance_file, header=None)
                else:
                    att_raw = pd.read_excel(attendance_file, header=None)
                
                # Extract Top Info
                org_name_val = str(att_raw.iloc[0, 0]).strip()
                center_name_val = str(att_raw.iloc[1, 0]).strip()
                
                # Reformat the dataframe to use Row 3 as columns
                att_df = att_raw.copy()
                att_df.columns = att_df.iloc[3]
                att_df = att_df.iloc[4:].reset_index(drop=True)
                
                # Validation check
                if 'StudentId' not in att_df.columns or 'StudentName' not in att_df.columns:
                    st.error("Invalid Attendance file. Could not find 'StudentId' or 'StudentName' columns. Did you accidentally upload the Fee Report here?")
                    st.stop()
                
                # Find current month columns
                month_cols = [col for col in att_df.columns if pd.notnull(col) and attendance_target in str(col)]
                
                if not month_cols:
                    st.error(f"Could not find any attendance data for {selected_month} {selected_year}. Please ensure you selected the correct month/year matching the file.")
                    st.stop()
                
                # Identify attendees
                att_df['Attended'] = att_df[month_cols].apply(lambda row: 'Present' in row.values, axis=1)
                attended_df = att_df[att_df['Attended'] == True].copy()
                
                # Ensure the extraction works even if some IDs are null/empty
                attended_df['Numeric_ID'] = attended_df['StudentId'].astype(str).str.extract(r'(\d+)$').astype(float)
                
                # --- 2. Process Fee Report ---
                if fee_file.name.endswith('.csv'):
                    fee_df = pd.read_csv(fee_file, skiprows=3)
                else:
                    fee_df = pd.read_excel(fee_file, skiprows=3)
                
                # Validation check
                if 'Particulars' not in fee_df.columns or 'Stud ID' not in fee_df.columns:
                    st.error("Invalid Fee file. Could not find 'Particulars' or 'Stud ID' columns. Did you accidentally upload the Attendance Report here?")
                    st.stop()
                    
                fee_df = fee_df.dropna(subset=['Particulars', 'Stud ID'])
                
                # Find paid students
                paid_df = fee_df[fee_df['Particulars'].str.contains(fee_target, case=False, na=False)]
                paid_ids = paid_df['Stud ID'].unique() 
                
                # --- 3. Find Unpaid ---
                unpaid_df = attended_df[~attended_df['Numeric_ID'].isin(paid_ids)][['StudentId', 'StudentName']].dropna()
                unpaid_df = unpaid_df.reset_index(drop=True)
                
                # --- 4. Render Output ---
                st.subheader("Report Output")
                st.write(f"**Organization:** {org_name_val}")
                st.write(f"**Center:** {center_name_val}")
                
                if unpaid_df.empty:
                    st.success(f"🎉 Great news! All students attending in {selected_month} {selected_year} have paid!")
                else:
                    st.warning(f"Found {len(unpaid_df)} students who attended without paying.")
                    st.dataframe(unpaid_df, use_container_width=True)
                    
                    # Generate PDF
                    pdf_path = create_pdf(unpaid_df, selected_month, selected_year, org_name_val, center_name_val)
                    
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                    
                    st.download_button(
                        label="📥 Download Professional PDF Report",
                        data=pdf_bytes,
                        file_name=f"Unpaid_Report_{center_name_val.replace(' ', '_')}_{selected_month}_{selected_year}.pdf",
                        mime="application/pdf",
                        type="primary"
                    )
        except Exception as e:
            st.error(f"An error occurred while processing the files: {e}")
    else:
        st.info("⚠️ Please upload both the Attendance and Fee reports to continue.")
