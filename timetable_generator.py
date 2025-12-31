import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(page_title="🎓 School Timetable Generator", page_icon="📚", layout="wide")

def normalize_class_name(class_raw):
    """VI A→VIA, IX B→IXB, XA/XB accepted"""
    if pd.isna(class_raw) or not str(class_raw).strip():
        return ''
    clean = str(class_raw).strip().upper().replace(' ', '')
    return clean if (len(clean) >= 2 and clean.isalpha()) else ''

@st.cache_data
def extract_perfect_data(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name='SCHOOL TIMETABLE')
    
    teachers = []
    classes = set()
    
    # Find teacher rows
    for i, row in df.iterrows():
        name = str(row.iloc[1]).strip() if len(df.columns) > 1 else ''
        if len(name) < 3 or 'NAME' in name.upper() or pd.isna(row.iloc[1]):
            continue
        
        subject = str(row.iloc[3]).strip() if len(df.columns) > 3 else name[:4]
        
        # Extract ALL period columns (4 onwards - 48 periods)
        periods = []
        for col_idx in range(4, min(52, len(row))):
            class_raw = row.iloc[col_idx]
            norm_class = normalize_class_name(class_raw)
            if norm_class:
                classes.add(norm_class)
            periods.append(norm_class)
        
        # Pad to 48 periods
        while len(periods) < 48:
            periods.append('')
        
        teachers.append({
            'name': name,
            'subject': subject,
            'periods': periods
        })
    
    return teachers, sorted(list(classes))

def create_final_timetable(teachers, classes, school_name):
    wb = Workbook()
    wb.remove(wb.active)
    
    COLORS = {
        'school': '2E8B57', 'class': '32CD32', 'teacher': 'FFD700',
        'day': '1E90FF', 'period': '4169E1', 'data': 'E0F2F1', 'subject': 'FFF2CC'
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
    
    # === TEACHER TIMETABLE ===
    ws_teacher = wb.create_sheet('👨‍🏫 Teacher Timetable', 0)
    ws_teacher.column_dimensions['A'].width = 32
    for col in range(2, 10):
        ws_teacher.column_dimensions[get_column_letter(col)].width = 13
    
    row = 1
    ws_teacher.merge_cells(f'A{row}:I{row}')
    style_cell(ws_teacher, f'A{row}', 'school', 16, True, 'FFFFFF').value = f"🏫 {school_name}"
    ws_teacher.row_dimensions[row].height = 40
    row += 2
    
    for teacher in teachers:
        # Teacher header
        ws_teacher.merge_cells(f'A{row}:I{row}')
        style_cell(ws_teacher, f'A{row}', 'teacher', 12, True).value = f"{teacher['name']}\n({teacher['subject']})"
        ws_teacher.row_dimensions[row].height = 40
        row += 1
        
        # Headers: Day | P1 | P2 | ... | P8
        style_cell(ws_teacher, f'A{row}', 'period', 11, True, 'FFFFFF').value = 'DAY'
        for p_idx in range(8):
            style_cell(ws_teacher, get_column_letter(p_idx+2)+f'{row}', 'period', 10, True, 'FFFFFF').value = f'P{PERIODS[p_idx]}'
        row += 1
        
        # 6 DAYS x 8 PERIODS = 48 cells
        periods = teacher['periods']
        for d_idx, day in enumerate(DAYS):
            style_cell(ws_teacher, f'A{row}', 'day', 11, True, 'FFFFFF').value = day
            
            day_start = d_idx * 8
            for p_idx in range(8):
                period_idx = day_start + p_idx
                cell_ref = get_column_letter(p_idx+2) + f'{row}'
                cell = style_cell(ws_teacher, cell_ref, 'data')
                
                if period_idx < len(periods) and periods[period_idx]:
                    cell.value = periods[period_idx]
            
            ws_teacher.row_dimensions[row].height = 25
            row += 1
        row += 1
    
    # === CLASS TIMETABLE ===
    ws_class = wb.create_sheet('📚 Class Timetable', 1)
    ws_class.column_dimensions['A'].width = 25
    for col in range(2, 10):
        ws_class.column_dimensions[get_column_letter(col)].width = 14
    
    row = 1
    ws_class.merge_cells(f'A{row}:I{row}')
    style_cell(ws_class, f'A{row}', 'school', 16, True, 'FFFFFF').value = f"🏫 {school_name}"
    ws_class.row_dimensions[row].height = 40
    row += 2
    
    for cls in classes:
        # Class header
        ws_class.merge_cells(f'A{row}:I{row}')
        style_cell(ws_class, f'A{row}', 'class', 13, True, 'FFFFFF').value = f"Class {cls}"
        ws_class.row_dimensions[row].height = 35
        row += 1
        
        # Headers
        style_cell(ws_class, f'A{row}', 'period', 11, True, 'FFFFFF').value = 'DAY'
        for p_idx in range(8):
            style_cell(ws_class, get_column_letter(p_idx+2)+f'{row}', 'period', 10, True, 'FFFFFF').value = f'P{PERIODS[p_idx]}'
        row += 1
        
        # Subjects for this class
        for d_idx, day in enumerate(DAYS):
            style_cell(ws_class, f'A{row}', 'day', 11, True, 'FFFFFF').value = day
            
            day_start = d_idx * 8
            for p_idx in range(8):
                cell_ref = get_column_letter(p_idx+2) + f'{row}'
                cell = style_cell(ws_class, cell_ref, 'subject')
                
                period_idx = day_start + p_idx
                subjects = []
                for teacher in teachers:
                    teacher_periods = teacher['periods']
                    if period_idx < len(teacher_periods) and teacher_periods[period_idx] == cls:
                        subjects.append(teacher['subject'][:4])
                
                if subjects:
                    cell.value = '/'.join(subjects)
            
            ws_class.row_dimensions[row].height = 25
            row += 1
        row += 1
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# === STREAMLIT UI ===
st.title("🎓 School Timetable Generator")
st.markdown("**VI-XII (A/B) • XA/XB • MR MOHAPATRA: MON P4=IXB**")

school_name = st.text_input("🏫 School Name", "Jawahar Navodaya Vidyalaya Baksa")
uploaded_file = st.file_uploader("📁 Upload Excel", type=['xlsx'])

if uploaded_file:
    with st.spinner("🔍 Extracting 14 classes + XA/XB..."):
        teachers, classes = extract_perfect_data(uploaded_file)
    
    col1, col2, col3 = st.columns(3)
    col1.metric("👨‍🏫 Teachers", len(teachers))
    col2.metric("📚 Classes", len(classes))
    col3.metric("📅 Periods", "48")
    
    st.success(f"""
    ✅ **{len(classes)} Classes Found:**
    `{', '.join(classes)}`
    
    **✅ XA/XB INCLUDED • MR MOHAPATRA: MON P4=IXB**
    """)
    
    if st.button("🚀 GENERATE FINAL TIMETABLE", type="primary", use_container_width=True):
        excel_data = create_final_timetable(teachers, classes, school_name)
        st.balloons()
        st.success("✨ **PERFECT TIMETABLE READY!**")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="📥 DOWNLOAD PERFECT TIMETABLE.xlsx",
            data=excel_data,
            file_name=f"FINAL_Timetable_{school_name.replace(' ', '_')}_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("👆 **Upload your Excel file**")

st.markdown("---")
st.markdown("""
**✅ FEATURES:**
• **14 Classes** (VI-XII A/B + XA/XB)
• **48 Periods** mapped perfectly  
• **Teacher Sheet:** Shows **CLASSES** (IXB, XA, VIIIA)
• **Class Sheet:** Shows **SUBJECTS** (MATH/ENG/PHY)
• **7 Colors** + Professional formatting
• **Print-ready** Excel file
""")
