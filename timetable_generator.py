import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="🎓 Timetable Generator", page_icon="📚", layout="wide")

def normalize_class_name(class_name):
    """VI A → VIA, IX B → IXB, XII A → XIIA"""
    if pd.isna(class_name):
        return None
    # Remove spaces and normalize
    clean = str(class_name).strip().upper().replace(' ', '')
    return clean if len(clean) > 2 else None

@st.cache_data
def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name='SCHOOL TIMETABLE')
    
    data_start = 0
    for i, row in df.iterrows():
        if pd.notna(row.iloc[1]) and len(str(row.iloc[1]).strip()) > 3:
            data_start = i
            break
    
    df_clean = df.iloc[data_start:].reset_index(drop=True)
    teachers = []
    classes_normalized = set()
    has_subject = len(df_clean.columns) > 3
    
    for idx, row in df_clean.iterrows():
        name = str(row.iloc[1]).strip()
        if len(name) < 3: continue
        
        # EXACT 48 periods (columns 4-51)
        periods_raw = pd.Series(row.iloc[4:52]).dropna().tolist()
        
        subject = str(row.iloc[3]).strip() if has_subject and pd.notna(row.iloc[3]) else name[:4]
        
        teacher = {
            'name': name,
            'subject': subject,
            'periods_raw': periods_raw,  # Keep raw for mapping
            'periods_normalized': []     # Normalized classes
        }
        
        # Normalize ALL classes for this teacher
        normalized_periods = []
        for period_raw in periods_raw:
            normalized_class = normalize_class_name(period_raw)
            if normalized_class:
                classes_normalized.add(normalized_class)
                normalized_periods.append(normalized_class)
            else:
                normalized_periods.append('')
        
        teacher['periods_normalized'] = normalized_periods
        teachers.append(teacher)
    
    return teachers, sorted(list(classes_normalized))

def create_perfect_timetable(teachers, classes, school_name):
    wb = Workbook()
    wb.remove(wb.active)
    
    COLORS = {
        'school': '2E8B57', 'header': '32CD32', 'teacher': 'FFD700',
        'day': '1E90FF', 'period': '4169E1', 'class_cell': 'E0F2F1', 'subject_cell': 'FFF2CC'
    }
    
    DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT']
    PERIODS = ['1', '2', '3', '4', '5', '6', '7', '8']
    
    def style_cell(ws, cell_ref, fill_color, size=10, bold=False, text_color='000000'):
        cell = ws[cell_ref]
        cell.font = Font(bold=bold, size=size, color=text_color)
        cell.fill = PatternFill(start_color=COLORS[fill_color], fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                       top=Side(style='thin'), bottom=Side(style='thin'))
        cell.border = border
        return cell
    
    # === TEACHER SHEET: Normalized Class Names ===
    ws_teacher = wb.create_sheet('👨‍🏫 Teacher Schedule', 0)
    ws_teacher.column_dimensions['A'].width = 30
    for col in range(2, 10): ws_teacher.column_dimensions[get_column_letter(col)].width = 13
    
    row = 1
    ws_teacher.merge_cells(f'A{row}:I{row}')
    style_cell(ws_teacher, f'A{row}', 'school', 14, True, 'FFFFFF').value = f"👨‍🏫 TEACHER SCHEDULE - {school_name}"
    ws_teacher.row_dimensions[row].height = 35
    row += 2
    
    for teacher in teachers:
        ws_teacher.merge_cells(f'A{row}:I{row}')
        style_cell(ws_teacher, f'A{row}', 'teacher', 11, True).value = f"{teacher['name']}\n📚 {teacher['subject']}"
        ws_teacher.row_dimensions[row].height = 35
        row += 1
        
        style_cell(ws_teacher, f'A{row}', 'period', 10, True, 'FFFFFF').value = 'DAY'
        for p_idx in range(8):
            style_cell(ws_teacher, get_column_letter(p_idx+2)+f'{row}', 'period', 9, True, 'FFFFFF').value = f'P{PERIODS[p_idx]}'
        row += 1
        
        # ALL 6 DAYS with NORMALIZED classes (VI A → VIA)
        normalized_periods = teacher['periods_normalized']
        for d_idx, day in enumerate(DAYS):
            style_cell(ws_teacher, f'A{row}', 'day', 10, True, 'FFFFFF').value = day
            
            day_start = d_idx * 8
            for p_idx in range(8):
                period_idx = day_start + p_idx
                cell_ref = get_column_letter(p_idx+2) + f'{row}'
                cell = style_cell(ws_teacher, cell_ref, 'class_cell')
                
                if period_idx < len(normalized_periods) and normalized_periods[period_idx]:
                    cell.value = normalized_periods[period_idx]  # VIA (not VI A)
            
            ws_teacher.row_dimensions[row].height = 22
            row += 1
        row += 1
    
    # === CLASS SHEET: SUBJECTS ===
    ws_class = wb.create_sheet('📚 Class Schedule', 1)
    ws_class.column_dimensions['A'].width = 25
    for col in range(2, 10): ws_class.column_dimensions[get_column_letter(col)].width = 14
    
    row = 1
    ws_class.merge_cells(f'A{row}:I{row}')
    style_cell(ws_class, f'A{row}', 'school', 14, True, 'FFFFFF').value = f"📚 CLASS SCHEDULE - {school_name}"
    ws_class.row_dimensions[row].height = 35
    row += 2
    
    for cls in classes:  # Normalized classes (VIA, not VI A)
        ws_class.merge_cells(f'A{row}:I{row}')
        style_cell(ws_class, f'A{row}', 'header', 12, True, 'FFFFFF').value = f"Class {cls}"  # VIA
        ws_class.row_dimensions[row].height = 32
        row += 1
        
        style_cell(ws_class, f'A{row}', 'period', 10, True, 'FFFFFF').value = 'DAY'
        for p_idx in range(8):
            style_cell(ws_class, get_column_letter(p_idx+2)+f'{row}', 'period', 9, True, 'FFFFFF').value = f'P{PERIODS[p_idx]}'
        row += 1
        
        for d_idx, day in enumerate(DAYS):
            style_cell(ws_class, f'A{row}', 'day', 10, True, 'FFFFFF').value = day
            
            day_start = d_idx * 8
            for p_idx in range(8):
                cell_ref = get_column_letter(p_idx+2) + f'{row}'
                cell = style_cell(ws_class, cell_ref, 'subject_cell')
                
                period_idx = day_start + p_idx
                subjects = []
                for teacher in teachers:
                    norm_periods = teacher['periods_normalized']
                    if (period_idx < len(norm_periods) and 
                        norm_periods[period_idx] == cls):
                        subjects.append(teacher['subject'][:4])
                
                if subjects:
                    cell.value = '/'.join(subjects)
            
            ws_class.row_dimensions[row].height = 22
            row += 1
        row += 1
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# === UI ===
st.title("🎓 School Timetable Generator")
st.markdown("**VI A = VIA • IX B = IXB • ALL normalized!**")

school_name = st.text_input("🏫 School Name", "Jawahar Navodaya Vidyalaya Baksa")

uploaded_file = st.file_uploader("📁 Upload Excel", type=['xlsx'])

if uploaded_file:
    with st.spinner("🔍 Normalizing classes (VI A → VIA)..."):
        teachers, classes = process_file(uploaded_file)
    
    col1, col2, col3 = st.columns(3)
    col1.metric("👨‍🏫 Teachers", len(teachers))
    col2.metric("📚 Unique Classes", len(classes))
    col3.metric("📅 Days", "6")
    
    st.success(f"""
    ✅ **Classes normalized!** ({len(classes)} unique)
    Sample: {', '.join(classes[:8])}{'...' if len(classes)>8 else ''}
    • VI A → **VIA** ✓
    • IX B → **IXB** ✓  
    • XII A → **XIIA** ✓
    """)
    
    if st.button("🚀 GENERATE", type="primary"):
        excel_data = create_perfect_timetable(teachers, classes, school_name)
        st.balloons()
        st.download_button(
            label="📥 DOWNLOAD NORMALIZED TIMETABLE",
            data=excel_data,
            file_name=f"Timetable_{school_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Upload file!")

st.markdown("""
**✨ Normalization:** VI A=**VIA** • IX B=**IXB** • XII A=**XIIA**
**📊 2 Sheets:** Teachers→Classes | Classes→Subjects
**🎨 7 Colors:** Professional formatting
""")
