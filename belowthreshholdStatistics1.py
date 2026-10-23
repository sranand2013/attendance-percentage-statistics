import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from collections import defaultdict

# -----------------------------------------------------------------------------
# Configuration & UI Setup
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Attendance Shortage Analyzer", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header { font-size: 2.5rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0rem;}
    .sub-header { font-size: 1.1rem; color: #6B7280; margin-bottom: 2rem;}
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Core Processing Logic
# -----------------------------------------------------------------------------
def process_attendance_files(files, threshold):
    subject_records = []
    student_records = []
    
    for file in files:
        spec_name = os.path.splitext(file.name)[0]
        xls = pd.ExcelFile(file)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            
            header_row_idx = None
            for idx, row in df.iterrows():
                if "Reg/USN" in row.values or "USN" in row.values:
                    header_row_idx = idx
                    break
            
            if header_row_idx is None:
                continue 
                
            subject_cols = {}
            overall_col = None
            
            for col_idx in range(4, df.shape[1]):
                val = str(df.iloc[header_row_idx, col_idx]).strip()
                if val and val.lower() != 'nan':
                    if 'overall' in val.lower():
                        overall_col = col_idx + 2  
                        break
                    else:
                        if ' - ' in val:
                            code, name = val.split(' - ', 1)
                        else:
                            code, name = val, val
                            
                        subject_cols[val] = {
                            'col_idx': col_idx + 2,
                            'code': code.strip(),
                            'name': name.strip()
                        }
            
            section = sheet_name.replace('Section - ', '').strip()
            
            for row_idx in range(header_row_idx + 2, df.shape[0]):
                sno = df.iloc[row_idx, 0]
                
                if pd.isna(sno): continue
                try:
                    int(sno)
                except ValueError:
                    continue
                
                # --- 1. Student-Level Processing (For Section/Row Strengths) ---
                overall_pct = None
                if overall_col is not None:
                    overall_pct = pd.to_numeric(df.iloc[row_idx, overall_col], errors='coerce')
                    
                is_above_threshold = False
                if pd.notna(overall_pct) and overall_pct >= threshold:
                    is_above_threshold = True
                    
                student_records.append({
                    'Specialization': spec_name,
                    'Section': section,
                    'SNo': sno,
                    'Is_Above_Threshold': is_above_threshold
                })

                # --- 2. Subject-Level Processing ---
                if overall_col is not None:
                    overall_conducted = pd.to_numeric(df.iloc[row_idx, overall_col - 2], errors='coerce')
                    if pd.notna(overall_conducted) and overall_conducted > 0:
                        is_shortage = False
                        if pd.notna(overall_pct) and overall_pct < threshold:
                            is_shortage = True
                            
                        subject_records.append({
                            'Specialization': spec_name,
                            'Section': section,
                            'Subject_Full': 'OVERALL %',
                            'Subject_Code': 'OVERALL',
                            'Subject_Name': 'OVERALL PERCENTAGE',
                            'Is_Shortage': is_shortage
                        })
                
                for full_name, info in subject_cols.items():
                    pct_col = info['col_idx']
                    conducted = pd.to_numeric(df.iloc[row_idx, pct_col - 2], errors='coerce')
                    pct = pd.to_numeric(df.iloc[row_idx, pct_col], errors='coerce')
                    
                    if pd.notna(conducted) and conducted > 0:
                        is_shortage = False
                        if pd.notna(pct) and pct < threshold:
                            is_shortage = True
                            
                        subject_records.append({
                            'Specialization': spec_name,
                            'Section': section,
                            'Subject_Full': full_name,
                            'Subject_Code': info['code'],
                            'Subject_Name': info['name'],
                            'Is_Shortage': is_shortage
                        })
                            
    return pd.DataFrame(subject_records), pd.DataFrame(student_records)

def format_applicability(tuples_series):
    unique_tuples = sorted(list(set(tuples_series)))
    spec_dict = defaultdict(list)
    for spec, sec in unique_tuples:
        spec_dict[spec].append(sec)
        
    out = []
    for spec, secs in spec_dict.items():
        out.append(f"{spec} ({', '.join(secs)})")
    return " | ".join(out)

def build_smart_matrix(df_subj, df_stud, group_cols, threshold):
    """Dynamically builds a side-by-side matrix (Below/Above) with Totals and NA logic."""
    thr_int = int(threshold)
    
    # 1. Row Metrics
    stud_stats = df_stud.groupby(group_cols).agg(
        Total_Students=('SNo', 'count'),
        Above_Threshold=('Is_Above_Threshold', 'sum')
    ).reset_index()
    
    # 2. Subject Metrics
    subj_agg = df_subj.groupby(group_cols + ['Subject_Full']).agg(
        Applicable=('Is_Shortage', 'count'),
        Below=('Is_Shortage', 'sum')
    ).reset_index()
    subj_agg['Above'] = subj_agg['Applicable'] - subj_agg['Below']
    
    # 3. Create Pivot Tables for fast lookups
    pivot_below = subj_agg.pivot_table(index=group_cols, columns='Subject_Full', values='Below', aggfunc='sum', fill_value=0)
    pivot_above = subj_agg.pivot_table(index=group_cols, columns='Subject_Full', values='Above', aggfunc='sum', fill_value=0)
    pivot_app = subj_agg.pivot_table(index=group_cols, columns='Subject_Full', values='Applicable', aggfunc='sum', fill_value=0)
    
    # Base dataframe to construct
    final_df = stud_stats.copy()
    col_total_strength = 'Total Row Strength'
    col_above_thr = f'Overall >= {thr_int}%'
    final_df.rename(columns={'Total_Students': col_total_strength, 'Above_Threshold': col_above_thr}, inplace=True)
    
    # Iterate over unique subjects and build side-by-side columns
    subjects = sorted(pivot_below.columns.tolist())
    subj_cols_ordered = []
    
    for subj in subjects:
        col_below = f"{subj} (< {thr_int}%)"
        col_above = f"{subj} (>= {thr_int}%)"
        subj_cols_ordered.extend([col_below, col_above])
        
        below_vals = []
        above_vals = []
        
        for idx, row in final_df.iterrows():
            lookup_idx = tuple(row[c] for c in group_cols) if len(group_cols) > 1 else row[group_cols[0]]
            
            # Check Applicability
            if lookup_idx in pivot_app.index and pivot_app.loc[lookup_idx, subj] > 0:
                below_vals.append(pivot_below.loc[lookup_idx, subj])
                above_vals.append(pivot_above.loc[lookup_idx, subj])
            else:
                below_vals.append('NA')
                above_vals.append('NA')
                
        final_df[col_below] = below_vals
        final_df[col_above] = above_vals
        
    # Calculate Total Shortages across all subjects (ignoring the 'OVERALL %' tracker)
    tot_shortages = []
    for idx, row in final_df.iterrows():
        shortage_sum = 0
        for subj in subjects:
            if subj != 'OVERALL %':
                val = row[f"{subj} (< {thr_int}%)"]
                if val != 'NA':
                    shortage_sum += val
        tot_shortages.append(shortage_sum)
        
    final_df['Total Shortages (All Subjects)'] = tot_shortages
    
    # Reorder the final DataFrame
    front_cols = group_cols
    back_cols = [col_total_strength, col_above_thr, 'Total Shortages (All Subjects)']
    final_df = final_df[front_cols + subj_cols_ordered + back_cols]
    
    # Create Bottom TOTAL row
    total_row = {}
    for col in front_cols:
        total_row[col] = 'TOTAL' if col == front_cols[0] else ''
        
    for subj in subjects:
        col_below = f"{subj} (< {thr_int}%)"
        col_above = f"{subj} (>= {thr_int}%)"
        tot_app = pivot_app[subj].sum()
        
        if tot_app > 0:
            total_row[col_below] = pivot_below[subj].sum()
            total_row[col_above] = pivot_above[subj].sum()
        else:
            total_row[col_below] = 'NA'
            total_row[col_above] = 'NA'
            
    total_row[col_total_strength] = stud_stats['Total_Students'].sum()
    total_row[col_above_thr] = stud_stats['Above_Threshold'].sum()
    total_row['Total Shortages (All Subjects)'] = sum(tot_shortages)
    
    final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)
    return final_df

# -----------------------------------------------------------------------------
# Web Interface Main Application
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙ Configuration")
    threshold = st.number_input(
        "Attendance Threshold (%)", 
        min_value=0.0, max_value=100.0, value=75.0, step=1.0,
        help="Calculates 'Above' and 'Below' logic dynamically."
    )
    st.markdown("---")
    st.info("Upload files where each sheet is a section and rows contain student subject-wise percentages.")

uploaded_files = st.file_uploader("Upload Specialization Excel Files", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("Generate Reports", type="primary"):
        with st.spinner('Parsing files and building reports...'):
            df_subj, df_stud = process_attendance_files(uploaded_files, threshold)
            
            if df_subj.empty:
                st.warning("No data found in the uploaded files.")
            else:
                st.success("✅ Analysis Complete!")
                thr_int = int(threshold)
                
                # ==========================================
                # 1. BUILD SMART MATRIX REPORTS
                # ==========================================
                section_matrix = build_smart_matrix(df_subj, df_stud, ['Specialization', 'Section'], threshold)
                spec_matrix = build_smart_matrix(df_subj, df_stud, ['Specialization'], threshold)
                
                matrix_output = io.BytesIO()
                with pd.ExcelWriter(matrix_output, engine='openpyxl') as writer:
                    section_matrix.to_excel(writer, sheet_name='Section Matrix', index=False)
                    spec_matrix.to_excel(writer, sheet_name='Specialization Matrix', index=False)
                
                # ==========================================
                # 2. BUILD SUBJECT-WISE SUMMARY REPORT
                # ==========================================
                subjects_df = df_subj[df_subj['Subject_Code'] != 'OVERALL'].copy()
                subjects_df['Spec_Tuple'] = list(zip(subjects_df['Specialization'], subjects_df['Section']))
                
                summary_df = subjects_df.groupby(['Subject_Code', 'Subject_Name']).agg(
                    Applicability=('Spec_Tuple', format_applicability),
                    Total_Students=('Is_Shortage', 'count'),
                    Below_Threshold=('Is_Shortage', 'sum')
                ).reset_index()
                
                summary_df['Above_Threshold'] = summary_df['Total_Students'] - summary_df['Below_Threshold']
                
                summary_df = summary_df.rename(columns={
                    'Subject_Code': 'Subject Code',
                    'Subject_Name': 'Subject',
                    'Applicability': 'Specialization & Sections',
                    'Total_Students': 'Subject wise Actual Students',
                    'Below_Threshold': f'No. of Students Below {thr_int}%',
                    'Above_Threshold': f'No of Above {thr_int}%'
                })
                
                columns_order = [
                    'Subject Code', 'Subject', 'Specialization & Sections', 
                    'Subject wise Actual Students', f'No. of Students Below {thr_int}%', 
                    f'No of Above {thr_int}%'
                ]
                summary_df = summary_df[columns_order]
                
                summary_output = io.BytesIO()
                with pd.ExcelWriter(summary_output, engine='openpyxl') as writer:
                    summary_df.to_excel(writer, sheet_name='All Data', startrow=1, index=False)
                    worksheet = writer.sheets['All Data']
                    current_date = datetime.now().strftime("%d-%m-%Y")
                    worksheet.cell(row=1, column=1, value=f"SUBJECT WISE DETAILS - AS ON {current_date}")
                    
                    for column_cells in worksheet.columns:
                        length = max(len(str(cell.value)) for cell in column_cells if cell.value is not None)
                        worksheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)

                # ==========================================
                # UI Previews & Downloads
                # ==========================================
                st.subheader(f"Section-Wise Smart Matrix Preview")
                st.dataframe(section_matrix, use_container_width=True, hide_index=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="💾 Download Smart Matrix Report",
                        data=matrix_output.getvalue(),
                        file_name=f"Attendance_Matrix_{thr_int}pct.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                
                with col2:
                    st.download_button(
                        label="💾 Download Subject-Wise Summary",
                        data=summary_output.getvalue(),
                        file_name=f"Subject_Wise_Summary_{thr_int}pct.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary"
                    )
else:
    st.info("Awaiting file uploads to begin.")