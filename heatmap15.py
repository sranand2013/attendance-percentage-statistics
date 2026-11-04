import streamlit as st
import pandas as pd
import io
import os
import tempfile
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import base64
import streamlit.components.v1 as components

st.set_page_config(page_title="Attendance Heatmap Generator", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0rem;}
    .sub-header { font-size: 1.1rem; color: #6B7280; margin-bottom: 1.5rem;}
    </style>
""", unsafe_allow_html=True)
st.markdown("<h1 class='main-header'>Attendance Shortage Heatmap Generator</h1>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Interactive heatmap where clicking any cell instantly downloads the specific student details (with embedded logos) for that shortage group.</div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙ Configuration")
    threshold = st.number_input("Attendance Threshold (%)", min_value=0.0, max_value=100.0, value=75.0, step=1.0)
    st.markdown("---")
    logo_file = st.file_uploader("Upload Institute Logo (PNG/JPG)", type=['png', 'jpg', 'jpeg'])

uploaded_files = st.file_uploader("Upload Specialization Excel Files", type=['xlsx'], accept_multiple_files=True)

def process_attendance(files, threshold):
    subject_records = []
    student_records = []
    
    for file in files:
        fname = file.name if hasattr(file, 'name') else os.path.basename(file)
        spec_name = fname.split('Attendance_Report_')[0].strip(' _-')
        if not spec_name:
            spec_name = os.path.splitext(fname)[0]
            
        xls = pd.ExcelFile(file)
        
        for sheet_name in xls.sheet_names:
            if sheet_name.strip().lower() == 'consolidated':
                continue
                
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            section = sheet_name.replace('Section - ', '').strip()
            
            header_row_idx = None
            usn_col = 1
            name_col = 2
            
            for idx, row in df.iterrows():
                row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if "reg/usn" in row_vals or "usn" in row_vals or "registration number" in row_vals:
                    header_row_idx = idx
                    for c_idx, val in enumerate(row.values):
                        if pd.notna(val):
                            v_str = str(val).strip().lower()
                            if v_str in ['reg/usn', 'usn', 'registration number']:
                                usn_col = c_idx
                            elif v_str in ['name', 'student name', 'name of student', 'name of the student']:
                                name_col = c_idx
                    break
            
            if header_row_idx is None:
                continue 
                
            subject_cols = {}
            overall_col = None
            
            for col_idx in range(4, df.shape[1]):
                val = str(df.iloc[header_row_idx, col_idx]).strip()
                if val and val.lower() != 'nan':
                    if 'overall' in val.lower():
                        if col_idx + 2 < df.shape[1]:
                            overall_col = col_idx + 2
                        break
                    elif 'signature' in val.lower():
                        continue
                    else:
                        if col_idx + 2 < df.shape[1]:
                            if ' - ' in val:
                                code, name = val.split(' - ', 1)
                            else:
                                code, name = val, val
                            subj_format = f"{name.strip()} ({code.strip()})"
                            subject_cols[subj_format] = col_idx + 2 
            
            for row_idx in range(header_row_idx + 2, df.shape[0]):
                sno_val = df.iloc[row_idx, 0]
                if pd.isna(sno_val):
                    continue
                try:
                    int(sno_val)
                except ValueError:
                    continue
                
                usn = df.iloc[row_idx, usn_col] if usn_col < df.shape[1] else sno_val
                student_name = df.iloc[row_idx, name_col] if name_col < df.shape[1] else "Unknown"
                
                overall_pct = None
                if overall_col is not None:
                    overall_pct = pd.to_numeric(df.iloc[row_idx, overall_col], errors='coerce')
                
                is_above_threshold = False
                if pd.notna(overall_pct) and overall_pct >= threshold:
                    is_above_threshold = True
                    
                student_records.append({
                    'Specialization': spec_name,
                    'Section': section,
                    'SNo': usn,
                    'Name': student_name,
                    'Is_Above_Threshold': is_above_threshold
                })
                
                if overall_col is not None:
                    overall_conducted = pd.to_numeric(df.iloc[row_idx, overall_col - 2], errors='coerce')
                    overall_attended = pd.to_numeric(df.iloc[row_idx, overall_col - 1], errors='coerce')
                    if pd.notna(overall_conducted) and overall_conducted > 0:
                        is_shortage = False
                        if pd.notna(overall_pct) and overall_pct < threshold:
                            is_shortage = True
                        subject_records.append({
                            'Specialization': spec_name,
                            'Section': section,
                            'SNo': usn,
                            'Name': student_name,
                            'Subject_Full': 'OVERALL %',
                            'Classes Conducted': overall_conducted,
                            'Classes Attended': overall_attended,
                            'Percentage': overall_pct,
                            'Is_Shortage': is_shortage
                        })
                
                for full_name, pct_col in subject_cols.items():
                    conducted = pd.to_numeric(df.iloc[row_idx, pct_col - 2], errors='coerce')
                    attended = pd.to_numeric(df.iloc[row_idx, pct_col - 1], errors='coerce')
                    pct = pd.to_numeric(df.iloc[row_idx, pct_col], errors='coerce')
                    if pd.notna(conducted) and conducted > 0:
                        is_shortage = False
                        if pd.notna(pct) and pct < threshold:
                            is_shortage = True
                        subject_records.append({
                            'Specialization': spec_name,
                            'Section': section,
                            'SNo': usn,
                            'Name': student_name,
                            'Subject_Full': full_name,
                            'Classes Conducted': conducted,
                            'Classes Attended': attended,
                            'Percentage': pct,
                            'Is_Shortage': is_shortage
                        })
                        
    return pd.DataFrame(subject_records), pd.DataFrame(student_records)

def build_smart_matrix(df_subj, df_stud, threshold):
    thr_int = int(threshold)
    
    stud_stats = df_stud.groupby(['Specialization', 'Section']).agg(
        Total_Students=('SNo', 'count'),
        Above_Threshold=('Is_Above_Threshold', 'sum')
    ).reset_index()
    
    subj_agg = df_subj.groupby(['Specialization', 'Section', 'Subject_Full']).agg(
        Applicable=('Is_Shortage', 'count'),
        Below=('Is_Shortage', 'sum')
    ).reset_index()
    subj_agg['Above'] = subj_agg['Applicable'] - subj_agg['Below']
    
    pivot_below = subj_agg.pivot_table(index=['Specialization', 'Section'], columns='Subject_Full', values='Below', aggfunc='sum', fill_value=0)
    pivot_above = subj_agg.pivot_table(index=['Specialization', 'Section'], columns='Subject_Full', values='Above', aggfunc='sum', fill_value=0)
    pivot_app = subj_agg.pivot_table(index=['Specialization', 'Section'], columns='Subject_Full', values='Applicable', aggfunc='sum', fill_value=0)
    
    final_df = stud_stats.copy()
    col_total_strength = 'Total Strength'
    col_above_thr = f'Overall >= {thr_int}%'
    final_df.rename(columns={'Total_Students': col_total_strength, 'Above_Threshold': col_above_thr}, inplace=True)
    
    all_subjects = sorted(pivot_below.columns.tolist())
    
    section_counts = {}
    for subj in all_subjects:
        if subj != 'OVERALL %':
            valid_sections = pivot_app[pivot_app[subj] > 0].shape[0]
            section_counts[subj] = valid_sections
            
    common_subjects = [s for s, count in section_counts.items() if count > 2]
    multi_other_subjects = [s for s in all_subjects if s not in common_subjects and s != 'OVERALL %' and section_counts.get(s, 0) > 1]
    single_section_subjects = [s for s in all_subjects if section_counts.get(s, 0) == 1]
    
    ordered_subjects = common_subjects + multi_other_subjects
    if 'OVERALL %' in all_subjects:
        ordered_subjects.append('OVERALL %')
        
    subj_cols_ordered = []
    for subj in ordered_subjects:
        col_below = f"{subj} (< {thr_int}%)"
        col_above = f"{subj} (>= {thr_int}%)"
        subj_cols_ordered.extend([col_below, col_above])
        
        below_vals = []
        above_vals = []
        
        for idx, row in final_df.iterrows():
            lookup_idx = (row['Specialization'], row['Section'])
            if lookup_idx in pivot_app.index and pivot_app.loc[lookup_idx, subj] > 0:
                below_vals.append(int(pivot_below.loc[lookup_idx, subj]))
                above_vals.append(int(pivot_above.loc[lookup_idx, subj]))
            else:
                below_vals.append('NA')
                above_vals.append('NA')
                
        final_df[col_below] = below_vals
        final_df[col_above] = above_vals
        
    tot_shortages = []
    for idx, row in final_df.iterrows():
        shortage_sum = 0
        for subj in common_subjects:
            val = row.get(f"{subj} (< {thr_int}%)", 'NA')
            if val != 'NA':
                shortage_sum += val
        tot_shortages.append(shortage_sum)
        
    final_df['Total Shortages (Common)'] = tot_shortages
    
    front_cols = ['Specialization', 'Section', col_total_strength]
    back_cols = [col_above_thr, 'Total Shortages (Common)']
    
    single_subj_cols = []
    for subj in single_section_subjects:
        col_below = f"{subj} (< {thr_int}%)"
        col_above = f"{subj} (>= {thr_int}%)"
        single_subj_cols.extend([col_below, col_above])
        
        below_vals = []
        above_vals = []
        
        for idx, row in final_df.iterrows():
            lookup_idx = (row['Specialization'], row['Section'])
            if lookup_idx in pivot_app.index and pivot_app.loc[lookup_idx, subj] > 0:
                below_vals.append(int(pivot_below.loc[lookup_idx, subj]))
                above_vals.append(int(pivot_above.loc[lookup_idx, subj]))
            else:
                below_vals.append('NA')
                above_vals.append('NA')
                
        final_df[col_below] = below_vals
        final_df[col_above] = above_vals
        
    col_layout = front_cols + subj_cols_ordered + back_cols + single_subj_cols
    final_df = final_df[col_layout]
    
    total_row = {}
    for col in front_cols:
        total_row[col] = 'TOTAL' if col == 'Specialization' else ''
        
    for subj in ordered_subjects + single_section_subjects:
        col_below = f"{subj} (< {thr_int}%)"
        col_above = f"{subj} (>= {thr_int}%)"
        tot_app = pivot_app[subj].sum()
        
        if tot_app > 0:
            total_row[col_below] = int(pivot_below[subj].sum())
            total_row[col_above] = int(pivot_above[subj].sum())
        else:
            total_row[col_below] = 'NA'
            total_row[col_above] = 'NA'
            
    total_row[col_total_strength] = int(stud_stats['Total_Students'].sum())
    total_row[col_above_thr] = int(stud_stats['Above_Threshold'].sum())
    total_row['Total Shortages (Common)'] = int(sum(tot_shortages))
    
    avg_row = {}
    for col in front_cols:
        avg_row[col] = 'AVERAGE' if col == 'Specialization' else ''
    for col in col_layout:
        if col not in front_cols:
            if col == col_above_thr:
                avg_row[col] = round(stud_stats['Above_Threshold'].mean(), 1)
            elif col == 'Total Shortages (Common)':
                avg_row[col] = round(pd.Series(tot_shortages).mean(), 1)
            elif col == col_total_strength:
                avg_row[col] = round(stud_stats['Total_Students'].mean(), 1)
            else:
                vals = pd.to_numeric(final_df[col], errors='coerce').dropna()
                avg_row[col] = round(vals.mean(), 1) if not vals.empty else 'NA'
                
    pct_row = {}
    for col in front_cols:
        pct_row[col] = 'PERCENTAGE' if col == 'Specialization' else ''
    for col in col_layout:
        if col not in front_cols:
            if col == col_total_strength:
                pct_row[col] = '100%'
            elif col in [col_above_thr, 'Total Shortages (Common)']:
                pct_row[col] = ''
            elif '(<' in col:
                above_col = col.replace('(<', '(>=')
                valid_mask = final_df[col] != 'NA'
                num_sum = pd.to_numeric(final_df.loc[valid_mask, col], errors='coerce').sum()
                above_sum = pd.to_numeric(final_df.loc[valid_mask, above_col], errors='coerce').sum()
                den_sum = num_sum + above_sum
                if den_sum > 0:
                    pct_val = (num_sum / den_sum) * 100
                    pct_row[col] = f"{round(pct_val, 1)}%"
                else:
                    pct_row[col] = '0%'
            elif '(>=' in col:
                below_col = col.replace('(>=', '(<')
                valid_mask = final_df[col] != 'NA'
                num_sum = pd.to_numeric(final_df.loc[valid_mask, col], errors='coerce').sum()
                below_sum = pd.to_numeric(final_df.loc[valid_mask, below_col], errors='coerce').sum()
                den_sum = num_sum + below_sum
                if den_sum > 0:
                    pct_val = (num_sum / den_sum) * 100
                    pct_row[col] = f"{round(pct_val, 1)}%"
                else:
                    pct_row[col] = '0%'

    ratio_sum_row = {}
    for col in front_cols:
        ratio_sum_row[col] = 'RATIO SUM (NUM/DEN)' if col == 'Specialization' else ''
    for col in col_layout:
        if col not in front_cols:
            if col == col_total_strength:
                ratio_sum_row[col] = int(stud_stats['Total_Students'].sum())
            elif col in [col_above_thr, 'Total Shortages (Common)']:
                ratio_sum_row[col] = ''
            elif '(<' in col:
                above_col = col.replace('(<', '(>=')
                valid_mask = final_df[col] != 'NA'
                num_sum = pd.to_numeric(final_df.loc[valid_mask, col], errors='coerce').sum()
                above_sum = pd.to_numeric(final_df.loc[valid_mask, above_col], errors='coerce').sum()
                den_sum = num_sum + above_sum
                if den_sum > 0:
                    ratio_sum_row[col] = f"{int(num_sum)}/{int(den_sum)}"
                else:
                    ratio_sum_row[col] = '0/0'
            elif '(>=' in col:
                below_col = col.replace('(>=', '(<')
                valid_mask = final_df[col] != 'NA'
                num_sum = pd.to_numeric(final_df.loc[valid_mask, col], errors='coerce').sum()
                below_sum = pd.to_numeric(final_df.loc[valid_mask, below_col], errors='coerce').sum()
                den_sum = num_sum + below_sum
                if den_sum > 0:
                    ratio_sum_row[col] = f"{int(num_sum)}/{int(den_sum)}"
                else:
                    ratio_sum_row[col] = '0/0'

    summary_df = pd.DataFrame([total_row, avg_row, pct_row, ratio_sum_row])
    final_df = pd.concat([final_df, summary_df], ignore_index=True)
    return final_df

def get_interpolated_color(pct):
    """Calculates a smooth gradient from Green (0%) to Yellow (50%) to Red (100%)."""
    if pct <= 0: return "#63BE7B"
    if pct >= 100: return "#F8696B"
    
    if pct <= 50:
        ratio = pct / 50.0
        r = int(99 + ratio * (255 - 99))
        g = int(190 + ratio * (235 - 190))
        b = int(123 + ratio * (132 - 123))
    else:
        ratio = (pct - 50) / 50.0
        r = int(255 + ratio * (248 - 255))
        g = int(235 + ratio * (105 - 235))
        b = int(132 + ratio * (107 - 132))
        
    return f"#{r:02x}{g:02x}{b:02x}"

def generate_interactive_html_heatmap(matrix_df, df_subj, threshold, logo_bytes):
    shortage_df = df_subj[df_subj['Is_Shortage'] == True].copy()
    
    shortage_dict = {}
    for (spec, sec, subj), group in shortage_df.groupby(['Specialization', 'Section', 'Subject_Full']):
        export_group = group[['SNo', 'Name', 'Specialization', 'Section', 'Subject_Full', 'Classes Attended', 'Classes Conducted', 'Percentage']].rename(
            columns={
                'SNo': 'Reg/USN',
                'Subject_Full': 'Subject (Name & Code)',
                'Percentage': 'Percentage (%)'
            }
        )
        shortage_dict[(spec, sec, subj)] = export_group.to_dict('records')
    
    html = ["<style>"]
    html.append(".hm-table { border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; min-width: 100%;}")
    html.append(".hm-th { border: 1px solid #d1d5db; background-color: #f3f4f6; padding: 12px 6px; text-align: center; position: sticky; top: 0; z-index: 10; font-weight: bold; }")
    html.append(".hm-th-left { border: 1px solid #d1d5db; background-color: #f3f4f6; padding: 12px 10px; text-align: left; position: sticky; left: 0; z-index: 11; font-weight: bold;}")
    html.append(".hm-td { border: 1px solid #d1d5db; text-align: center; padding: 0; vertical-align: middle; }")
    html.append(".hm-link { display: flex; flex-direction: column; justify-content: center; align-items: center; padding: 12px 4px; text-decoration: none; color: #111827; height: 100%; width: 100%; box-sizing: border-box; }")
    html.append(".hm-link:hover { background-color: rgba(0,0,0,0.08); font-weight: 600; text-decoration: underline; }")
    html.append(".hm-nolink { display: flex; flex-direction: column; justify-content: center; align-items: center; padding: 12px 4px; color: #374151; height: 100%; }")
    html.append("</style>")
    
    html.append("<div style='overflow: auto; max-height: 800px; width: 100%; border: 1px solid #d1d5db;'>")
    html.append("<table class='hm-table'>")
    
    data_df = matrix_df.iloc[:-4].copy()
    subjects_to_plot = [c for c in matrix_df.columns if '(<' in c and 'OVERALL' not in c]
    
    html.append("<tr>")
    html.append("<th class='hm-th-left'>Specialization & Section</th>")
    for subj in subjects_to_plot:
        clean_subj = subj.split(' (<')[0]
        html.append(f"<th class='hm-th'>{clean_subj}</th>")
    html.append("</tr>")
    
    def get_excel_b64(records, label_text, subj_clean, logo_bytes_data):
        if not records: return ""
        temp_df = pd.DataFrame(records)
        temp_df = temp_df[['Reg/USN', 'Name', 'Specialization', 'Section', 'Subject (Name & Code)', 'Classes Attended', 'Classes Conducted', 'Percentage (%)']]
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            temp_df.to_excel(writer, index=False, startrow=5, sheet_name='Shortages')
            worksheet = writer.sheets['Shortages']
            workbook = writer.book
            
            title_fmt = workbook.add_format({'bold': True, 'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'text_wrap': True, 'align': 'center'})
            data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            
            worksheet.merge_range(1, 1, 3, 7, f"Attendance Shortages: {subj_clean}\n{label_text}", title_fmt)
            
            if logo_bytes_data:
                logo_io = io.BytesIO(logo_bytes_data)
                worksheet.insert_image('A1', 'logo.png', {'image_data': logo_io, 'x_scale': 0.45, 'y_scale': 0.45, 'positioning': 1})
            
            worksheet.set_row(5, 30)
            for i, col in enumerate(temp_df.columns):
                worksheet.write(5, i, col, header_fmt)
                max_len = max(temp_df[col].astype(str).map(len).max(), len(col)) + 4
                worksheet.set_column(i, i, max_len, data_fmt)
                
        return base64.b64encode(out.getvalue()).decode()
        
    for _, row in data_df.iterrows():
        spec = row['Specialization']
        sec = row['Section']
        label = f"{str(spec).replace('2026_BBA ', '')} ({sec})"
        
        html.append("<tr>")
        html.append(f"<th class='hm-th-left'>{label}</th>")
        
        for subj_col in subjects_to_plot:
            subj_clean = subj_col.split(' (<')[0]
            above_col = subj_col.replace('(<', '(>=')
            
            b_val = row.get(subj_col, 'NA')
            a_val = row.get(above_col, 'NA')
            
            if b_val == 'NA' or pd.isna(b_val):
                html.append("<td class='hm-td' style='background-color:#ffffff;'><div class='hm-nolink'>NA</div></td>")
            else:
                b = int(float(b_val))
                a = int(float(a_val))
                total_subj = b + a
                
                if total_subj > 0:
                    pct = (b / total_subj) * 100
                    bg = get_interpolated_color(pct)
                    
                    if b > 0:
                        records = shortage_dict.get((spec, sec, subj_clean), [])
                        b64 = get_excel_b64(records, label, subj_clean, logo_bytes)
                        safe_name = "".join([c for c in subj_clean if c.isalnum()]).rstrip()
                        filename = f"Shortages_{sec}_{safe_name}.xlsx"
                        href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
                        html.append(f"<td class='hm-td' style='background-color:{bg};'><a href='{href}' download='{filename}' class='hm-link'><span>{b}/{total_subj}</span><span>({pct:.1f}%)</span></a></td>")
                    else:
                        html.append(f"<td class='hm-td' style='background-color:{bg};'><div class='hm-nolink'><span>{b}/{total_subj}</span><span>({pct:.1f}%)</span></div></td>")
                else:
                    html.append("<td class='hm-td' style='background-color:#ffffff;'><div class='hm-nolink'><span>0/0</span><span>(0.0%)</span></div></td>")
        html.append("</tr>")
        
    summary_rows = matrix_df.iloc[-4:].copy()
    for _, row in summary_rows.iterrows():
        label = row['Specialization']
        html.append("<tr>")
        html.append(f"<th class='hm-th-left' style='background-color: #e5e7eb;'>{label}</th>")
        for subj_col in subjects_to_plot:
            val = str(row.get(subj_col, ''))
            html.append(f"<td class='hm-td' style='background-color: #e5e7eb; font-weight: bold;'><div class='hm-nolink'>{val}</div></td>")
        html.append("</tr>")
        
    html.append("</table>")
    html.append("</div>")
    return "".join(html)

def generate_heatmap_image(df, threshold):
    data_df = df.iloc[:-4].copy()
    data_df['Label'] = data_df['Specialization'].apply(lambda x: str(x).replace('2026_BBA ', '')) + ' (' + data_df['Section'].astype(str) + ')'
    
    pct_matrix = pd.DataFrame(index=data_df['Label'].values)
    annot_matrix = pd.DataFrame(index=data_df['Label'].values, dtype=str)
    
    for col in df.columns:
        if '(<' in col and 'OVERALL' not in col:
            subj_name = col.split(' (<')[0]
            above_col = col.replace('(<', '(>=')
            
            below_vals = data_df[col].values
            above_vals = data_df[above_col].values
            
            p_vals = []
            a_vals = []
            
            for b_val, a_val in zip(below_vals, above_vals):
                if b_val == 'NA' or pd.isna(b_val):
                    p_vals.append(0.0)
                    a_vals.append("NA")
                else:
                    b = float(b_val)
                    a = float(a_val)
                    subj_strength = b + a
                    
                    if subj_strength > 0:
                        pct = (b / subj_strength) * 100
                        p_vals.append(pct)
                        a_vals.append(f"{int(b)}/{int(subj_strength)}\n({pct:.1f}%)")
                    else:
                        p_vals.append(0.0)
                        a_vals.append("0/0\n(0.0%)")
                        
            pct_matrix[subj_name] = p_vals
            annot_matrix[subj_name] = a_vals
            
    active_pct = pct_matrix.loc[:, (pct_matrix != 0).any(axis=0)]
    active_annot = annot_matrix.loc[:, (pct_matrix != 0).any(axis=0)]
    
    fig, ax_heat = plt.subplots(figsize=(24, 12), dpi=300)
    
    sns.heatmap(active_pct, annot=active_annot, fmt='', cmap='YlOrRd', linewidths=0.8, 
                cbar_kws={'label': 'Percentage of Students Below Threshold (%)'}, ax=ax_heat,
                annot_kws={'size': 8, 'weight': 'normal'})
    
    ax_heat.set_title(f'Student Attendance Shortages Heatmap (Threshold < {int(threshold)}%)', fontsize=18, fontweight='bold', pad=20)
    ax_heat.set_xlabel('Subjects', fontsize=13, fontweight='bold', labelpad=12)
    ax_heat.set_ylabel('Specialization & Section', fontsize=13, fontweight='bold', labelpad=12)
    plt.setp(ax_heat.get_xticklabels(), rotation=45, ha='right', fontsize=10, fontweight='semibold')
    
    plt.tight_layout()
    img_output = io.BytesIO()
    plt.savefig(img_output, format='png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    img_output.seek(0)
    return img_output.getvalue()

def generate_excel(df, logo_bytes, chart_img_bytes, threshold):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    df.to_excel(writer, sheet_name='Attendance Matrix', index=False, startrow=6)
    worksheet = writer.sheets['Attendance Matrix']
    workbook_obj = writer.book
    
    header_fmt = workbook_obj.add_format({
        'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2', 'font_size': 9
    })
    data_fmt = workbook_obj.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 9, 'text_wrap': True
    })
    bold_fmt = workbook_obj.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F2F2F2', 'font_size': 9, 'text_wrap': True
    })
    title_fmt = workbook_obj.add_format({
        'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'
    })
    
    last_col_idx = max(len(df.columns) - 1, 10)
    worksheet.merge_range(1, 2, 3, min(last_col_idx, 10), f"Student Attendance Shortage Matrix (< {int(threshold)}%)", title_fmt)
    
    tmp_path = None
    if logo_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            tmp.write(logo_bytes)
            tmp_path = tmp.name
        worksheet.insert_image('A1', tmp_path, {'x_scale': 0.5, 'y_scale': 0.5, 'positioning': 1})
    
    worksheet.set_row(6, 45)
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(6, col_num, col_name, header_fmt)
        max_len = max(len(str(col_name)), 8) + 1
        worksheet.set_column(col_num, col_num, min(max_len, 18), data_fmt)
            
    data_last_row = 6 + len(df) - 5
    
    for col_num, col_name in enumerate(df.columns):
        if "(<" in col_name:
            worksheet.conditional_format(7, col_num, data_last_row, col_num, {
                'type': '3_color_scale',
                'min_color': '#63BE7B', 'min_type': 'num', 'min_value': 0,
                'mid_color': '#FFEB84', 'mid_type': 'percentile', 'mid_value': 50,
                'max_color': '#F8696B', 'max_type': 'percentile', 'max_value': 100
            })
            
    for summary_row_idx in range(data_last_row + 1, data_last_row + 5):
        worksheet.set_row(summary_row_idx, 28)
        for col_num in range(len(df.columns)):
            val = df.iloc[summary_row_idx - 6, col_num]
            worksheet.write(summary_row_idx, col_num, val, bold_fmt)
            
    for r_idx in range(7, data_last_row + 1):
        worksheet.set_row(r_idx, 24)
    
    legend_row = data_last_row + 6
    worksheet.write(legend_row, 0, "Color Legend:", workbook_obj.add_format({'bold': True, 'font_size': 9}))
    worksheet.write(legend_row, 1, "0 Shortages", workbook_obj.add_format({'bg_color': '#63BE7B', 'border': 1, 'align': 'center', 'font_size': 9}))
    worksheet.write(legend_row, 2, "Moderate", workbook_obj.add_format({'bg_color': '#FFEB84', 'border': 1, 'align': 'center', 'font_size': 9}))
    worksheet.write(legend_row, 3, "High Shortages", workbook_obj.add_format({'bg_color': '#F8696B', 'border': 1, 'align': 'center', 'font_size': 9}))
    
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 1)
    worksheet.set_margins(0.2, 0.2, 0.3, 0.3)

    ws2 = workbook_obj.add_worksheet('Attendance Matrix (2)')
    ws2.merge_range(1, 2, 3, 7, f"Student Attendance Shortage Matrix (< {int(threshold)}%)", title_fmt)
    
    if logo_bytes and tmp_path:
        ws2.insert_image('A1', tmp_path, {'x_scale': 0.5, 'y_scale': 0.5, 'positioning': 1})
        
    all_subject_pairs = []
    col_names = df.columns.tolist()
    i = 3
    while i < len(col_names) - 2:
        c1 = col_names[i]
        c2 = col_names[i+1]
        if '(<' in c1 and '(>=' in c2:
            subj_title = c1.split(' (<')[0]
            all_subject_pairs.append((subj_title, c1, c2))
            i += 2
        else:
            i += 1
            
    ws2.set_column(0, 0, 22)
    ws2.set_column(1, 1, 12)
    ws2.set_column(2, 2, 14)
    
    data_len = len(df) - 4
    subjects_per_block = 2
    num_blocks = (len(all_subject_pairs) + subjects_per_block - 1) // subjects_per_block
    
    for block_idx in range(num_blocks):
        start_row = 6 + block_idx * (data_len + 6)
        
        ws2.write(start_row, 0, 'Specialization', header_fmt)
        ws2.write(start_row, 1, 'Section', header_fmt)
        ws2.write(start_row, 2, 'Total Strength', header_fmt)
        ws2.set_row(start_row, 45)
        
        for r_idx in range(data_len):
            ws2.write(start_row + 1 + r_idx, 0, df.iloc[r_idx, 0], data_fmt)
            ws2.write(start_row + 1 + r_idx, 1, df.iloc[r_idx, 1], data_fmt)
            ws2.write(start_row + 1 + r_idx, 2, df.iloc[r_idx, 2], data_fmt)
            ws2.set_row(start_row + 1 + r_idx, 24)
            
        summary_labels = ['TOTAL', 'AVERAGE', 'PERCENTAGE', 'RATIO SUM (NUM/DEN)']
        for s_idx, s_lab in enumerate(summary_labels):
            ws2.write(start_row + 1 + data_len + s_idx, 0, s_lab, bold_fmt)
            ws2.write(start_row + 1 + data_len + s_idx, 1, '', bold_fmt)
            ws2.write(start_row + 1 + data_len + s_idx, 2, df.iloc[data_len + s_idx, 2], bold_fmt)
            ws2.set_row(start_row + 1 + data_len + s_idx, 28)
            
        block_pairs = all_subject_pairs[block_idx * subjects_per_block : (block_idx + 1) * subjects_per_block]
        for idx_in_block, (subj_title, c_below, c_above) in enumerate(block_pairs):
            col_c1 = 3 + idx_in_block * 2
            col_c2 = col_c1 + 1
            
            ws2.write(start_row, col_c1, c_below, header_fmt)
            ws2.write(start_row, col_c2, c_above, header_fmt)
            ws2.set_column(col_c1, col_c2, 18)
            
            for r_idx in range(data_len):
                ws2.write(start_row + 1 + r_idx, col_c1, df.iloc[r_idx][c_below], data_fmt)
                ws2.write(start_row + 1 + r_idx, col_c2, df.iloc[r_idx][c_above], data_fmt)
                
            for s_idx in range(4):
                sum_row_src_idx = data_len + s_idx
                ws2.write(start_row + 1 + data_len + s_idx, col_c1, df.iloc[sum_row_src_idx][c_below], bold_fmt)
                ws2.write(start_row + 1 + data_len + s_idx, col_c2, df.iloc[sum_row_src_idx][c_above], bold_fmt)

    ws2.set_landscape()
    ws2.fit_to_pages(1, 0)
    ws2.set_margins(0.2, 0.2, 0.3, 0.3)

    chart_sheet = workbook_obj.add_worksheet('Shortage Heatmap')
    chart_tmp_path = None
    if chart_img_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as ctmp:
            ctmp.write(chart_img_bytes)
            chart_tmp_path = ctmp.name
        chart_sheet.insert_image('B2', chart_tmp_path, {'x_scale': 0.7, 'y_scale': 0.7})
        
    writer.close()
    
    for p in [tmp_path, chart_tmp_path]:
        if p:
            try:
                os.remove(p)
            except OSError:
                pass
                
    return output.getvalue()

if uploaded_files:
    if st.button("Generate Matrix & Interactive Heatmap", type="primary"):
        with st.spinner("Processing documents, configuring interactive layout & Excel workbook..."):
            df_subj, df_stud = process_attendance(uploaded_files, threshold)
            logo_bytes = logo_file.read() if logo_file else None
            
            if df_subj.empty:
                st.warning("No valid data found in uploads. Please check your Excel formatting.")
            else:
                matrix_df = build_smart_matrix(df_subj, df_stud, threshold)
                
                interactive_html = generate_interactive_html_heatmap(matrix_df, df_subj, threshold, logo_bytes)
                static_heatmap_bytes = generate_heatmap_image(matrix_df, threshold)
                
                st.success("✅ Analysis and Interactive Heatmap Generation Complete!")
                st.subheader("🖱 Interactive Attendance Shortage Heatmap")
                st.info("💡 **Click on any cell containing shortages to instantly download the detailed Excel list (with embedded logos) of students below the threshold.**")
                
                components.html(interactive_html, height=850, scrolling=True)
                
                excel_data = generate_excel(matrix_df, logo_bytes, static_heatmap_bytes, threshold)
                
                st.download_button(
                    label="💾 Download Master Excel (Wrapped Statistics & Heatmap)",
                    data=excel_data,
                    file_name=f"Attendance_Interactive_Report_{int(threshold)}pct.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )