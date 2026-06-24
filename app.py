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

# -----------------------------------------------------------------------------
# إعدادات واجهة المستخدم
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام تنسيق وتدقيق كشوفات الوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2E4053; color: white; width: 100%; font-weight: bold; border-radius: 8px; font-size: 18px;}
    .report-box { background-color: #F4F6F7; padding: 15px; border-radius: 8px; border-right: 5px solid #2E4053; text-align: right; margin-bottom: 10px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام تنسيق وتدقيق كشوفات الوكلاء المطور 📄💎</h1>", unsafe_allow_html=True)

if "processing_done" not in st.session_state:
    st.session_state.processing_done = False
    st.session_state.df_final = None
    st.session_state.rejected_rows = []
    st.session_state.output_filename = ""

# -----------------------------------------------------------------------------
# مساعدات التنسيق المتقدمة لملفات Word
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

def set_cell_vertical_text(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    text_dir = parse_xml(f'<w:textDirection {nsdecls("w")} w:val="btLr"/>')
    tcPr.append(text_dir)

def set_cell_no_wrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    no_wrap = parse_xml(f'<w:noWrap {nsdecls("w")}/>')
    tcPr.append(no_wrap)

def format_cell_advanced(cell, text, bold=False, color_rgb=None, size_pt=16, font_name="Calibri", align="center"):
    cell.text = str(text)
    p = cell.paragraphs[0]
    
    if align == "right": p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "left": p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    pPr = p.paragraph_format.element.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    
    for run in p.runs:
        run.bold = bold
        if color_rgb: run.font.color.rgb = color_rgb
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)
        rPr.append(rFonts)
        run.font.size = Pt(size_pt)

# -----------------------------------------------------------------------------
# محرك قراءة وتنظيف البيانات (المعدل بالكامل بدون حذف السجلات للتشخيص)
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj):
    doc = Document(file_obj)
    rejected_rows = []
    raw_records = []
    
    lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text.strip())
            
    for table in doc.tables:
        for row in table.rows:
            row_text = "   ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
            if row_text.strip():
                lines.append(row_text)

    for line in lines:
        parts = [p.strip() for p in re.split(r'\s{2,}|,', line) if p.strip()]
        
        if not parts or "اسم رب" in line or "الوكالة" in line or "ةلاكولا" in line or "الوكيل" in line or "المركز" in line:
            continue
            
        if len(parts) < 4:
            rejected_rows.append(line)
            continue
            
        name_parts = [p for p in parts if re.search(r'[أ-ي]', p)]
        if not name_parts:
            rejected_rows.append(line)
            continue
        name = name_parts[0]
        
        numeric_parts = [p for p in parts if p.isdigit()]
        if len(numeric_parts) < 2:
            rejected_rows.append(line)
            continue
            
        try:
            old_card_num = numeric_parts[1] if len(numeric_parts) > 1 else numeric_parts[0]
            total_m = numeric_parts[-3] if len(numeric_parts) >= 3 else "0"
            sub_m = numeric_parts[-2] if len(numeric_parts) >= 3 else "0"
            rej_m = numeric_parts[-1] if len(numeric_parts) >= 3 else "0"
        except Exception:
            rejected_rows.append(line)
            continue

        # 1- تصحيح الأسماء المقلوبة تلقائياً بدلاً من حذفها
        fixed_words = []
        for w in name.split():
            if w.startswith('ة') and len(w) > 1:
                fixed_words.append(w[::-1]) 
            else:
                fixed_words.append(w)
        name = " ".join(fixed_words)
        
        # 2- الاحتفاظ بكافة القيود (تمت إزالة الفلترة لتشخيص الأعداد الحقيقية)
        raw_records.append({
            "اسم رب الأسرة": name,
            "رقم البطاقة القديم": old_card_num,
            "الكلي": total_m,
            "مستحق": sub_m,
            "محجوب": rej_m
        })

    df = pd.DataFrame(raw_records)
    if not df.empty:
        # ترتيب أبجدي وإعطاء تسلسل (بدون drop_duplicates)
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
        
    return df, rejected_rows

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word الاحترافي
# -----------------------------------------------------------------------------
def build_professional_word_report(df, filename_base):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(0.3)
        section.right_margin = Cm(0.3)
        
    clean_name = filename_base
    words_to_remove = ["مستكشف", "معدل", "كشف", "منسق", "جاهز"]
    for w in words_to_remove: clean_name = clean_name.replace(w, "")
    clean_name = re.sub(r'[a-zA-Z]', '', clean_name)
    clean_name = re.sub(r'[\-_+_.]', '', clean_name)
    clean_name = " ".join(clean_name.split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name = "Segoe UI Semibold"
    title_run.font.size = Pt(14)
    title_run.bold = True
    
    headers = ["ت", "اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", "رقم البطاقة القديم", "ملاحظات"]
    
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    
    col_widths = [Cm(0.9), dynamic_name_width, Cm(0.35), Cm(0.9), Cm(0.9), Cm(0.9), Cm(3.0), Cm(1.75)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        if i in [3, 4, 5]:
            set_cell_vertical_text(hdr_cells[i])
            format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        else:
            cell_align = "left" if i == 1 else "center"
            format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align=cell_align, color_rgb=COLOR_NAVY_BLUE)
            
    HEX_ELEGANT_BLUE, HEX_LIGHT_BLUE, HEX_LIGHT_RED, HEX_LIGHT_GREEN, HEX_ALERT_RED = "D4E6F1", "EBF5FB", "FADBD8", "E8F8F5", "EC7063"
    
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        table.rows[idx+1]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        # حماية من القيم الفارغة أو غير المتوقعة
        try:
            is_eligible_zero = int(row["مستحق"]) == 0
        except:
            is_eligible_zero = False
            
        for i in range(8): row_cells[i].width = col_widths[i]
        set_cell_no_wrap(row_cells[1])
        
        for i in range(8):
            val, text_color, cell_align, font_size = "", None, "center", 16
            
            if i == 0: val = row["ت"]
            elif i == 1: val, cell_align = row["اسم رب الأسرة"], "left" 
            elif i == 2: val = "x" if is_eligible_zero else "" 
            elif i == 3: val = row["الكلي"]
            elif i == 4: val = row["مستحق"]
            elif i == 5: val, font_size = row["محجوب"], 10 
            elif i == 6: val = row["رقم البطاقة القديم"]
            elif i == 7: 
                val = "محجوب" if is_eligible_zero else "" 
                if is_eligible_zero: text_color, font_size = RGBColor(203, 67, 53), 12
                    
            format_cell_advanced(row_cells[i], val, size_pt=font_size, font_name="Calibri", color_rgb=text_color, align=cell_align)
            
            if is_eligible_zero: set_cell_background(row_cells[i], HEX_ALERT_RED)
            else:
                if i == 0: set_cell_background(row_cells[i], HEX_ELEGANT_BLUE)
                elif i == 3: set_cell_background(row_cells[i], HEX_LIGHT_BLUE)
                elif i == 4: set_cell_background(row_cells[i], HEX_LIGHT_GREEN)
                elif i == 5: set_cell_background(row_cells[i], HEX_LIGHT_RED)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة Streamlit 
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف المراد تدقيقه</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("ارفع كشف الوكلاء", type=['docx'], key="doc_input_v6", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename:
        st.session_state.processing_done = False

if st.button("⚙️ تشغيل محرك التنظيم والتنسيق المتقدم الكلي"):
    if uploaded_file:
        with st.spinner('جاري قراءة البيانات ومعالجتها...'):
            try:
                df_res, rejected_rows = extract_and_clean_data(uploaded_file)
                if not df_res.empty:
                    st.session_state.df_final = df_res
                    st.session_state.rejected_rows = rejected_rows
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.processing_done = True
                else:
                    st.error("لم يتم العثور على بيانات متوافقة في الملف.")
            except Exception as e:
                st.error(f"خطأ غير متوقع أثناء المعالجة: {e}")
    else:
        st.warning("الرجاء رفع ملف docx أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    rejected = st.session_state.rejected_rows
    output_filename = st.session_state.output_filename
    
    # عرض إحصائيات المقارنة للمستخدم فوراً
    col1, col2 = st.columns(2)
    with col1:
        st.success(f"✅ إجمالي السجلات المستخرجة بنجاح: **{len(df_final)}** قيد.")
    with col2:
        if len(rejected) > 0:
            st.error(f"⚠️ إجمالي الأسطر/السجلات المرفوضة: **{len(rejected)}** سطر.")
        else:
            st.success("✅ لم يتم رفض أي سطر (0 مرفوض).")
    
    # تحميل ملف التشخيص في حال وجود نواقص
    if len(rejected) > 0:
        st.markdown("<div class='report-box'><strong>يوجد أسطر لم يقرأها الكود (مرفوضة)، حملها من هنا لمعرفة الخلل:</strong></div>", unsafe_allow_html=True)
        df_rej = pd.DataFrame({"السطر_المرفوض": rejected})
        
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df_rej.to_excel(writer, index=False, sheet_name='المرفوضات')
        
        st.download_button(
            label="📉 تحميل ملف الأسطر المرفوضة (Excel)",
            data=excel_buffer.getvalue(),
            file_name=f"مرفوضات_{output_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("<hr>", unsafe_allow_html=True)
    
    with st.spinner('جاري صياغة وهيكلة مستند Word المطور...'):
        word_output = build_professional_word_report(df_final, output_filename)
        
    st.download_button(
        label="📥 تحميل كشف الوكلاء المنسق والجاهز للطباعة فوراً (Word)",
        data=word_output,
        file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
