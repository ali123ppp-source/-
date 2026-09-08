import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import re

# إعدادات واجهة المستخدم
st.set_page_config(page_title="نظام تنسيق وتدقيق كشوفات الوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2E4053; color: white; width: 100%; font-weight: bold; border-radius: 8px; font-size: 18px;}
    .report-box { background-color: #F4F6F7; padding: 15px; border-radius: 8px; border-right: 5px solid #2E4053; text-align: right; margin-bottom: 10px;}
    div.row-widget.stRadio > div { flex-direction: row-reverse; justify-content: flex-start; gap: 20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام تنسيق وتدقيق كشوفات الوكلاء المطور 📄💎</h1>", unsafe_allow_html=True)

_default_state = {
    "processing_done": False,
    "df_final": None,
    "output_filename": "",
    "selected_card": "",
    "template_choice": "",
    "input_format_choice": "",
    "reference_filename": "",
    "corrections_df": None,
    "unresolved_df": None,
    "matched_count": 0,
}
for _key, _default_val in _default_state.items():
    if _key not in st.session_state:
        st.session_state[_key] = _default_val

# -----------------------------------------------------------------------------
# مساعدات التنسيق المتقدمة لملفات Word عبر الـ XML
# -----------------------------------------------------------------------------
def set_table_borders(table, color_hex="2A4B7C"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="{color_hex}"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="{color_hex}"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color_hex}"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="{color_hex}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_cell_advanced(cell, text, bold=False, color_rgb=None, size_pt=16, font_name="Calibri", align="center"):
    cell.text = str(text)
    p = cell.paragraphs[0]
    
    if align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    pPr = p.paragraph_format.element.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    
    for run in p.runs:
        run.bold = bold
        if color_rgb:
            run.font.color.rgb = color_rgb
            
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)
        rPr.append(rFonts)
        run.font.size = Pt(size_pt)

# -----------------------------------------------------------------------------
# محرك قراءة وتنظيف البيانات المطور (يدعم Word و Excel)
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj, card_choice):
    raw_records = []
    rows_data = []
    
    file_ext = file_obj.name.split('.')[-1].lower()
    
    if file_ext == 'docx':
        doc = Document(file_obj)
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                rows_data.append(cells)
                
    elif file_ext == 'xlsx':
        xls = pd.ExcelFile(file_obj)
        for sheet_name in xls.sheet_names:
            df_excel = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            for row in df_excel.values:
                cells = []
                for cell in row:
                    if pd.isna(cell):
                        continue
                    if isinstance(cell, float) and cell.is_integer():
                        cells.append(str(int(cell)))
                    else:
                        cells.append(str(cell).strip().replace('\n', ' '))
                rows_data.append(cells)
    
    for cells in rows_data:
        if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
            continue
        
        name_idx = -1
        max_len = 0
        for i, c in enumerate(cells):
            if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                if len(c) > max_len:
                    max_len = len(c)
                    name_idx = i
        if name_idx == -1: continue
        
        card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
        if not card_indices: continue
        
        old_card_num = cells[card_indices[0]]
        new_card_num = cells[card_indices[1]] if len(card_indices) > 1 else old_card_num
        selected_card_num = new_card_num if card_choice == "رقم البطاقة الحديث" else old_card_num
        
        digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
        if len(digit_cells) >= 3:
            withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
        elif len(digit_cells) == 2:
            withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
        else:
            continue
        
        full_name = cells[name_idx]
        name_parts = full_name.split()
        three_part_name = " ".join(name_parts[:3])
            
        raw_records.append({
            "اسم رب الأسرة": three_part_name,
            "رقم البطاقة": selected_card_num,
            "الكلي": total,
            "مستحق": eligible,
            "محجوب": withheld
        })
        
    df = pd.DataFrame(raw_records)
    if not df.empty:
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
    return df

# -----------------------------------------------------------------------------
# محرك قراءة إضافي مخصص لكشوفات "القطع الغذائية" 
# -----------------------------------------------------------------------------
def extract_ration_list_data(file_obj, card_choice):
    file_ext = file_obj.name.split('.')[-1].lower()
    header_row = None
    data_rows = []

    if file_ext == 'docx':
        doc = Document(file_obj)
        for table in doc.tables:
            table_rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            if not table_rows:
                continue
            if header_row is None:
                header_row = table_rows[0]
                data_rows.extend(table_rows[1:])
            else:
                start = 1 if table_rows[0] == header_row else 0
                data_rows.extend(table_rows[start:])

    elif file_ext == 'xlsx':
        xls = pd.ExcelFile(file_obj)
        for sheet_name in xls.sheet_names:
            df_excel = pd.read_excel(xls, sheet_name=sheet_name, header=0, dtype=str)
            df_excel.columns = [str(c).strip() for c in df_excel.columns]
            if header_row is None:
                header_row = list(df_excel.columns)
            for row in df_excel.values:
                data_rows.append(["" if pd.isna(v) else str(v).strip() for v in row])

    if not header_row:
        return pd.DataFrame()

    def find_col(possible_names):
        for i, h in enumerate(header_row):
            for name in possible_names:
                if name in h:
                    return i
        return -1

    idx_name = find_col(["اسم العائلة", "اسم رب"])
    idx_old = find_col(["البطاقة القديمة"])
    idx_new = find_col(["البطاقة الجديدة"])
    idx_total = find_col(["الأفراد الكلية", "كلي"])
    idx_eligible = find_col(["الأفراد المستحقة", "مستحق"])
    idx_withheld = find_col(["المحجوبين", "محجوب"])

    if -1 in (idx_name, idx_total, idx_eligible, idx_withheld):
        return pd.DataFrame()

    raw_records = []
    for cells in data_rows:
        if cells == header_row:
            continue
        if not any(str(c).strip() for c in cells):
            continue
        if idx_name >= len(cells):
            continue
            
        # تجاوز صف الإجماليات/المجاميع الختامية إن وُجد في نهاية الجدول
        if any("الإجمالي" in str(c) for c in cells) or any("المجموع" in str(c) for c in cells):
            continue
            
        # استخراج وتصحيح الأسماء والأرقام
        full_name = cells[idx_name]
        name_parts = full_name.split()
        three_part_name = " ".join(name_parts[:3])
        
        old_card_num = cells[idx_old] if idx_old != -1 else ""
        new_card_num = cells[idx_new] if idx_new != -1 else ""
        selected_card_num = new_card_num if card_choice == "رقم البطاقة الحديث" and new_card_num else old_card_num
        if not selected_card_num:
            selected_card_num = old_card_num or new_card_num
            
        try:
            total = int(cells[idx_total]) if str(cells[idx_total]).isdigit() else 0
            eligible = int(cells[idx_eligible]) if str(cells[idx_eligible]).isdigit() else 0
            withheld = int(cells[idx_withheld]) if str(cells[idx_withheld]).isdigit() else 0
            
            raw_records.append({
                "اسم رب الأسرة": three_part_name,
                "رقم البطاقة": selected_card_num,
                "الكلي": total,
                "مستحق": eligible,
                "محجوب": withheld
            })
        except ValueError:
            continue

    df = pd.DataFrame(raw_records)
    if not df.empty:
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
    return df

# -----------------------------------------------------------------------------
# دالة توليد الوورد المنسق للتحميل
# -----------------------------------------------------------------------------
def generate_word_report(df, agent_name="كشف الوكيل"):
    doc = Document()
    doc.add_heading(f"بيانات الوكيل: {agent_name} - مرتبة أبجدياً", level=1)
    
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    set_table_borders(table)
    
    # رأس الجدول
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        set_cell_background(hdr_cells[i], "2A4B7C")
        format_cell_advanced(hdr_cells[i], col_name, bold=True, color_rgb=RGBColor(255, 255, 255))
        
    # تعبئة البيانات
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            format_cell_advanced(row_cells[i], str(val))
            
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة Streamlit الرئيسية
# -----------------------------------------------------------------------------
st.write("قم برفع ملف الكشف (Word أو Excel) ليتم تحليله، ترتيبه أبجدياً، وتجهيزه للتحميل المباشر.")

uploaded_file = st.file_uploader("📂 ارفع الملف هنا", type=["docx", "xlsx"])
card_choice = st.radio("اختر نوع رقم البطاقة المراد اعتماده:", ["رقم البطاقة الحديث", "رقم البطاقة القديم"])

if uploaded_file is not None:
    if st.button("🚀 معالجة البيانات وتحليلها"):
        with st.spinner("جاري تحليل ومعالجة البيانات..."):
            # محاولة المعالجة باستخدام المحرك المخصص أولاً، إن فشل نستخدم المحرك العام
            df_result = extract_ration_list_data(uploaded_file, card_choice)
            if df_result.empty:
                df_result = extract_and_clean_data(uploaded_file, card_choice)
                
            if not df_result.empty:
                st.session_state.df_final = df_result
                st.session_state.processing_done = True
                
                # عرض التحليل بشكل مبسط في الواجهة
                st.success("✅ تمت المعالجة بنجاح!")
                st.markdown(f"**إجمالي العوائل:** {len(df_result)} | **إجمالي الأفراد:** {df_result['الكلي'].sum()} | **الأفراد المستحقون:** {df_result['مستحق'].sum()}")
                st.dataframe(df_result.head(10))
            else:
                st.error("⚠️ لم يتم العثور على بيانات قابلة للاستخراج، تأكد من تنسيق الملف.")

# زر التحميل المباشر (يظهر فقط بعد اكتمال المعالجة)
if st.session_state.processing_done and st.session_state.df_final is not None:
    st.markdown("### 📥 تحميل الملف النهائي")
    word_buffer = generate_word_report(st.session_state.df_final)
    
    st.download_button(
        label="💾 اضغط هنا لتحميل الكشف المنسق (Word)",
        data=word_buffer,
        file_name="الكشف_النهائي_المرتب.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
