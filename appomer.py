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
# إعدادات الواجهة
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام كشوفات الوكلاء الخارق", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #1A5276; color: white; width: 100%; font-weight: bold; border-radius: 8px; font-size: 18px;}
    .report-box { background-color: #F4F6F7; padding: 15px; border-radius: 8px; border-right: 5px solid #1A5276; text-align: right; margin-bottom: 10px;}
    div.row-widget.stRadio > div { flex-direction: row-reverse; justify-content: flex-start; gap: 20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right; color: #1A5276;'>نظام تدقيق وتنسيق الوكلاء الخارق 👁️⚡</h1>", unsafe_allow_html=True)

if "processing_done" not in st.session_state:
    st.session_state.processing_done = False
    st.session_state.df_final = None
    st.session_state.df_duplicates = None
    st.session_state.output_filename = ""
    st.session_state.selected_card = ""

# -----------------------------------------------------------------------------
# دوال التنسيق المتقدمة لـ Word
# -----------------------------------------------------------------------------
def set_table_borders(table, color_hex="1A5276"):
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
# وحدات الاستخراج المعزولة (الأساس والأركان كما هي)
# -----------------------------------------------------------------------------
def _parse_docx_to_list(file_obj, card_choice):
    """نفس المحرك الجبري الأصلي دون مساس (يقرأ Word)"""
    doc = Document(file_obj)
    raw_records = []
    lines = []
    
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.replace('\n', ' ').strip() for c in row.cells if c.text.strip()]
            lines.append(" | ".join(cells))
            
    for para in doc.paragraphs:
        if para.text.strip() and "|" not in para.text:
            lines.append(para.text.strip())
            
    for line in lines:
        line_clean = line.replace(',', ' ').replace('،', ' ').replace('-', ' ')
        if not line_clean: continue
        
        if any(w in line_clean for w in ["المركز", "اسم رب", "ملاحظات", "الوكيل", "الافراد", "الكلية", "المحجوبين", "FOOD", "نوع الوكالة"]):
            continue
            
        text_only = re.sub(r'[^\u0600-\u06FF\s\|]', ' ', line_clean)
        segments = [s.strip() for s in text_only.split('|')]
        if len(segments) <= 1:
            segments = [s.strip() for s in text_only.split('  ')]
            
        valid_names = []
        for s in segments:
            s_clean = re.sub(r'\s+', ' ', s).strip()
            if len(s_clean.split()) >= 2:
                valid_names.append(s_clean)
                
        if not valid_names:
            fallback = re.findall(r'[\u0600-\u06FF]{2,}(?:\s+[\u0600-\u06FF]{2,})+', line_clean)
            if fallback: valid_names = fallback
            else: continue
                
        full_name = max(valid_names, key=len).strip()
        
        all_nums = re.findall(r'\d+', line_clean)
        cards = [n for n in all_nums if len(n) >= 5]
        
        old_card = str(cards[0]) if cards else "غير متوفر"
        new_card = str(cards[-1]) if len(cards) >= 2 else old_card
        selected_card = new_card if card_choice == "رقم البطاقة الحديث" else old_card
        
        idx = line_clean.find(full_name.split()[0])
        str_before = line_clean[:idx]
        str_after = line_clean[idx:]
        
        smalls_before = [int(n) for n in re.findall(r'\d+', str_before) if len(n) < 5]
        smalls_after = [int(n) for n in re.findall(r'\d+', str_after) if len(n) < 5]
        
        total = eligible = withheld = 0
        
        if len(smalls_after) >= 3:
            total, eligible, withheld = smalls_after[0], smalls_after[1], smalls_after[2]
        elif len(smalls_before) >= 3:
            withheld, eligible, total = smalls_before[-3], smalls_before[-2], smalls_before[-1]
        elif len(smalls_after) == 2:
            total, eligible = smalls_after[0], smalls_after[1]
            withheld = 0
        elif len(smalls_before) == 2:
            eligible, total = smalls_before[-2], smalls_before[-1]
            withheld = 0
        else:
            all_smalls = smalls_before + smalls_after
            if len(all_smalls) >= 3:
                total, eligible, withheld = all_smalls[-3], all_smalls[-2], all_smalls[-1]
            elif len(all_smalls) == 2:
                total, eligible, withheld = max(all_smalls), min(all_smalls), 0
                
        raw_records.append({
            "اسم رب الأسرة": full_name,
            "رقم البطاقة": selected_card,
            "الكلي": total,
            "محجوب": withheld,
            "مستحق": eligible
        })
        
    return raw_records

def _parse_excel_to_list(file_obj, file_ext):
    """وحدة الاستخراج المرنة لقراءة جداول الـ Excel أو الـ CSV ودمجها"""
    records = []
    try:
        df_in = pd.read_excel(file_obj) if file_ext == 'xlsx' else pd.read_csv(file_obj)
        df_in.columns = df_in.columns.astype(str).str.strip()
        
        # التقاط الأعمدة ديناميكياً لتجنب توقف الكود بسبب اختلاف تسميات الأعمدة في الإكسل
        col_name = next((c for c in df_in.columns if 'اسم' in c), df_in.columns[1] if len(df_in.columns)>1 else df_in.columns[0])
        col_card = next((c for c in df_in.columns if 'بطاق' in c), df_in.columns[6] if len(df_in.columns)>6 else df_in.columns[-1])
        col_total = next((c for c in df_in.columns if 'كلي' in c), None)
        col_eligible = next((c for c in df_in.columns if 'مستحق' in c), None)
        col_withheld = next((c for c in df_in.columns if 'محجوب' in c), None)

        for _, row in df_in.iterrows():
            name = str(row[col_name]).strip()
            if not name or name == 'nan' or 'اسم' in name: continue
                
            card = str(row[col_card]).strip()
            if card == 'nan': card = "غير متوفر"
            
            def safe_int(val):
                try: return int(float(val))
                except: return 0
                
            records.append({
                "اسم رب الأسرة": name,
                "رقم البطاقة": card,
                "الكلي": safe_int(row[col_total]) if col_total else 0,
                "محجوب": safe_int(row[col_withheld]) if col_withheld else 0,
                "مستحق": safe_int(row[col_eligible]) if col_eligible else 0
            })
    except Exception as e:
        st.warning(f"تنبيه: حدث خطأ أثناء تحليل ملف الإكسل/الجدول الإضافي ({e})")
    return records

# -----------------------------------------------------------------------------
# الماستر: الدمج والفرز والتنظيف (مع الحفاظ على الكيان)
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj, card_choice, file_obj2=None, file2_ext=None):
    # 1. استخراج الأساسي
    raw_records = _parse_docx_to_list(file_obj, card_choice)
    
    # 2. استخراج ودمج الإضافي (إن وجد)
    if file_obj2:
        if file2_ext == 'docx':
            raw_records.extend(_parse_docx_to_list(file_obj2, card_choice))
        elif file2_ext in ['xlsx', 'csv']:
            raw_records.extend(_parse_excel_to_list(file_obj2, file2_ext))
            
    df = pd.DataFrame(raw_records)
    df_duplicates = pd.DataFrame()
    
    if not df.empty:
        duplicate_mask = df.duplicated(keep='first')
        df_duplicates = df[duplicate_mask].copy()
        
        if not df_duplicates.empty:
            df_duplicates = df_duplicates.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
            df_duplicates.insert(0, "ت", df_duplicates.index + 1)

        # الكشف النهائي المعقم والمدمج
        df = df.drop_duplicates(keep='first')
        # الفرز الأبجدي يتم هنا لجميع الملفات المدمجة تلقائياً
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
        
    return df, df_duplicates, len(raw_records)

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word
# -----------------------------------------------------------------------------
def build_professional_word_report(df, filename_base, card_choice, is_duplicate_report=False):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(0.3)
        section.right_margin = Cm(0.3)
        
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز"]: clean_name = clean_name.replace(w, "")
    clean_name = re.sub(r'[a-zA-Z]', '', clean_name)
    clean_name = re.sub(r'[\-_+_.]', '', clean_name)
    clean_name = " ".join(clean_name.split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    report_title = f"كشف المتكررات المعزولة: {clean_name}" if is_duplicate_report else f"الكشف الإحصائي المنسق للوكيل: {clean_name}"
    
    title_run = title_p.add_run(report_title)
    title_run.font.name = "Segoe UI Semibold"
    title_run.font.size = Pt(14)
    title_run.bold = True
    title_run.font.color.rgb = RGBColor(220, 20, 60) if is_duplicate_report else RGBColor(26, 82, 118)
    
    footer = doc.sections[0].footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f_run = footer_p.add_run("صفحة ")
    f_run.font.size = Pt(10)
    fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    fldChar3 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    f_run._r.extend([fldChar1, instrText, fldChar2, fldChar3])
    
    headers = ["ت", "اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    border_color = "DC143C" if is_duplicate_report else "1A5276"
    set_table_borders(table, color_hex=border_color)
    
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    
    col_widths = [Cm(0.9), dynamic_name_width, Cm(0.44), Cm(0.9), Cm(0.9), Cm(0.9), Cm(3.0), Cm(2.19)]
    THEME_COLOR = RGBColor(220, 20, 60) if is_duplicate_report else RGBColor(26, 82, 118)
    
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        if i in [3, 4, 5]:
            set_cell_vertical_text(hdr_cells[i])
            format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=THEME_COLOR)
        else:
            cell_align = "left" if i == 1 else "center"
            format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align=cell_align, color_rgb=THEME_COLOR)
            
    HEX_ELEGANT_BLUE = "D4E6F1"
    HEX_LIGHT_BLUE = "EBF5FB"
    HEX_LIGHT_RED = "FADBD8"
    HEX_LIGHT_GREEN = "E8F8F5"
    HEX_ALERT_RED = "EC7063"
    
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        r_trPr = table.rows[idx+1]._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        is_eligible_zero = int(row["مستحق"]) == 0
        
        for i in range(8): row_cells[i].width = col_widths[i]
        set_cell_no_wrap(row_cells[1])
        
        for i in range(8):
            val, text_color, cell_align, font_size = "", None, "center", 16
            
            if i == 0: val = row["ت"]
            elif i == 1: 
                val = row["اسم رب الأسرة"]
                cell_align = "left" 
            elif i == 2: val = "x" if is_eligible_zero else "" 
            elif i == 3: val = row["الكلي"]
            elif i == 4: val = row["مستحق"]
            elif i == 5: 
                val = row["محجوب"]
                font_size = 14  
            elif i == 6: val = row["رقم البطاقة"] 
            elif i == 7: 
                val = "محجوب" if is_eligible_zero else "" 
                if is_eligible_zero:
                    text_color = RGBColor(203, 67, 53)
                    font_size = 12  
                    
            format_cell_advanced(row_cells[i], val, size_pt=font_size, font_name="Calibri", color_rgb=text_color, align=cell_align)
            
            if is_eligible_zero: set_cell_background(row_cells[i], HEX_ALERT_RED)
            else:
                if i == 0: set_cell_background(row_cells[i], HEX_ELEGANT_BLUE)
                elif i == 3: set_cell_background(row_cells[i], HEX_LIGHT_BLUE)   
                elif i == 4: set_cell_background(row_cells[i], HEX_LIGHT_GREEN)  
                elif i == 5: set_cell_background(row_cells[i], HEX_LIGHT_RED)    

    total_all = df["الكلي"].astype(int).sum()
    total_eligible = df["مستحق"].astype(int).sum()
    total_withheld = df["محجوب"].astype(int).sum()
    
    doc.add_paragraph()  
    stats_p = doc.add_paragraph()
    stats_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    stats_p.paragraph_format.element.get_or_add_pPr().append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    
    stats_text = (f"العدد الكلي للافراد = {total_all}\n"
                  f"العدد الكلي للمستحقين = {total_eligible}\n"
                  f"العدد الكلي للمحجوبين = {total_withheld}")
    stats_run = stats_p.add_run(stats_text)
    stats_run.font.name = "Segoe UI Semibold"
    stats_run.font.size = Pt(13)
    stats_run.bold = True
    stats_run.font.color.rgb = THEME_COLOR

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة الاستخدام المُحسّنة (رفع ملفين)
# -----------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col2:
    st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف الأساسي (Word)</h3>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("ارفع كشف الوكلاء الأساسي", type=['docx'], key="doc_input_master", label_visibility="collapsed")

with col1:
    st.markdown("<h3 style='text-align: right;'>➕ ملف للدمج (اختياري)</h3>", unsafe_allow_html=True)
    uploaded_file2 = st.file_uploader("ارفع جدول (Excel/Word/CSV) لدمجه أبجدياً", type=['docx', 'xlsx', 'csv'], key="doc_input_merge", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)
selected_card = st.radio("اختر نوع رقم البطاقة المراد اعتماده في الكشف:", ["رقم البطاقة القديم", "رقم البطاقة الحديث"], index=0, horizontal=True)
st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename or st.session_state.selected_card != selected_card:
        st.session_state.processing_done = False

if st.button("🚀 تشغيل المحرك الجبري (استخراج ودمج وفرز)"):
    if uploaded_file:
        with st.spinner('يتم الآن سحق الجداول وتحليل الشيفرات ودمج البيانات أبجدياً...'):
            try:
                # معرفة صيغة الملف الثاني إن وجد
                file2_ext = uploaded_file2.name.split('.')[-1].lower() if uploaded_file2 else None
                
                df_res, df_dup, total_scanned = extract_and_clean_data(uploaded_file, selected_card, uploaded_file2, file2_ext)
                
                if not df_res.empty:
                    st.session_state.df_final = df_res
                    st.session_state.df_duplicates = df_dup
                    st.session_state.total_scanned = total_scanned
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.selected_card = selected_card
                    st.session_state.processing_done = True
                else:
                    st.error("لم يتم العثور على بيانات جداول متوافقة. يرجى التأكد من محتوى الملف.")
            except Exception as e:
                st.error(f"خطأ غير متوقع: {e}")
    else:
        st.warning("الرجاء رفع ملف الكشف الأساسي (Word) أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    df_duplicates = st.session_state.df_duplicates
    output_filename = st.session_state.output_filename
    used_card_type = st.session_state.selected_card
    
    st.balloons()
    merge_text = " (شاملة البيانات المدمجة)" if uploaded_file2 else ""
    st.success(f"🏆 المهمة أُنجزت! تم استخراج وترتيب ({len(df_final)}) قيد صافٍ أبجدياً{merge_text}.")
    
    with st.spinner('جاري صياغة الكشف الذهبي الأساسي...'):
        word_output = build_professional_word_report(df_final, output_filename, used_card_type, is_duplicate_report=False)
        
    st.download_button(
        label=f"📥 تحميل الكشف الأساسي المدمج ({len(df_final)} اسم) جاهز للطباعة",
        data=word_output,
        file_name=f"كشف_منسق_مدمج_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    
    if not df_duplicates.empty:
        st.warning(f"⚠️ تم عزل ({len(df_duplicates)}) قيد كـ (نسخ ولصق حرفي).")
        with st.spinner('جاري بناء ملف التكرارات...'):
            word_output_dup = build_professional_word_report(df_duplicates, output_filename, used_card_type, is_duplicate_report=True)
            
        st.download_button(
            label="🚨 تحميل ملف القيود المتكررة المعزولة",
            data=word_output_dup,
            file_name=f"سجلات_متكررة_معزولة_{output_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
