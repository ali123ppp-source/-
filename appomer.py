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
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام تنسيق وتدقيق كشوفات الوكلاء المطور 📄💎</h1>", unsafe_allow_html=True)

if "processing_done" not in st.session_state:
    st.session_state.processing_done = False
    st.session_state.df_final = None
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
# المحرك الذكي لقراءة وتنظيف البيانات (مقاوم للتمزق والتشوه)
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj):
    doc = Document(file_obj)
    
    # 1. سحب كل النصوص من المستند بغض النظر عن الجداول لمنع مشكلة التمزق
    all_text = []
    for p in doc.paragraphs:
        if p.text.strip(): all_text.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip().replace('\n', ' ')
                if txt: all_text.append(txt)
                
    full_text = " ".join(all_text)
    tokens = full_text.split()
    
    raw_records = []
    seen_cards = set()
    
    # كلمات محجوزة يجب استبعادها من الأسماء
    stop_words = {"اسم", "رب", "الاسرة", "نوع", "الوكالة", "رقم", "المركز", "ت", "البطاقة", "القديم", "الافراد", "الكلية", "المستحقة", "المحجوبين"}
    
    i = 0
    while i < len(tokens):
        token = tokens[i]
        
        # التقاط رقم بطاقة (الجديدة أو القديمة طولها بين 5 إلى 8 أرقام)
        if token.isdigit() and 5 <= len(token) <= 8:
            card = token
            
            if card not in seen_cards:
                # تجميع البيانات المحيطة بالبطاقة (نافذة من 15 كلمة قبل وبعد)
                start_idx = max(0, i - 15)
                end_idx = min(len(tokens), i + 15)
                window = tokens[start_idx:end_idx]
                
                # استخراج الاسم العربي
                name_words = [w for w in window if re.match(r'^[أ-ي]+$', w) and w not in stop_words]
                clean_name = " ".join(name_words)
                
                # استخراج الأرقام الصغيرة (أعداد الأفراد)
                nums = [int(w) for w in window if w.isdigit() and len(w) <= 2]
                
                # استخراج التسلسل الأصلي
                orig_seq = ""
                for j in range(i-1, max(0, i-10), -1):
                    if tokens[j].isdigit() and 1 <= len(tokens[j]) <= 3:
                        orig_seq = tokens[j]
                        break
                
                # التحقق من اكتمال القيد
                if clean_name and len(nums) >= 3:
                    total = nums[0]
                    eligible = nums[1]
                    withheld = nums[2]
                    
                    # ترتيب منطقي للأفراد: الرقم الكلي يجب أن يكون الأكبر دائماً
                    if total < eligible:
                        total, eligible = eligible, total
                        
                    raw_records.append({
                        "ت": orig_seq if orig_seq else str(len(raw_records) + 1),
                        "اسم رب الأسرة": clean_name,
                        "رقم البطاقة القديم": card,
                        "الكلي": total,
                        "مستحق": eligible,
                        "محجوب": withheld
                    })
                    
                    # منع التكرار: إضافة البطاقة وأي بطاقة أخرى مجاورة في نفس القيد إلى الذاكرة
                    for w in window:
                        if w.isdigit() and 5 <= len(w) <= 8:
                            seen_cards.add(w)
        i += 1
        
    # البناء الآمن للجدول (لمنع خطأ ValueError نهائياً)
    if raw_records:
        df = pd.DataFrame(raw_records)
        return df
    else:
        # إرجاع جدول فارغ مهيأ بالأعمدة في حال كان الملف غير مقروء
        return pd.DataFrame(columns=["ت", "اسم رب الأسرة", "رقم البطاقة القديم", "الكلي", "مستحق", "محجوب"])

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
    for w in words_to_remove:
        clean_name = clean_name.replace(w, "")
    clean_name = re.sub(r'[a-zA-Z]', '', clean_name)
    clean_name = re.sub(r'[\-_+_.]', '', clean_name)
    clean_name = " ".join(clean_name.split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name = "Segoe UI Semibold"
    title_run.font.size = Pt(14)
    title_run.bold = True
    
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
    
    headers = ["ت", "اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", "رقم البطاقة القديم", "ملاحظات"]
    
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    set_table_borders(table, color_hex="2A4B7C")
    
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    
    col_widths = [
        Cm(0.9),              # ت
        dynamic_name_width,   # اسم رب الأسرة
        Cm(0.35),             # حقل فارغ
        Cm(0.9),              # الكلي
        Cm(0.9),              # مستحق
        Cm(0.9),              # محجوب
        Cm(3.0),              # رقم البطاقة القديم
        Cm(1.75)              # ملاحظات
    ]
    
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
            
    HEX_ELEGANT_BLUE = "D4E6F1"
    HEX_LIGHT_BLUE = "EBF5FB"
    HEX_LIGHT_RED = "FADBD8"
    HEX_LIGHT_GREEN = "E8F8F5"
    HEX_ALERT_RED = "EC7063"
    
    # تعبئة صفوف البيانات
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        
        r_trPr = table.rows[idx+1]._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        is_eligible_zero = int(row["مستحق"]) == 0
        
        for i in range(8):
            row_cells[i].width = col_widths[i]
            
        set_cell_no_wrap(row_cells[1])
        
        for i in range(8):
            val = ""
            text_color = None
            cell_align = "center"
            font_size = 16
            
            if i == 0: val = row["ت"]
            elif i == 1: 
                val = row["اسم رب الأسرة"]
                cell_align = "left" 
            elif i == 2: val = "x" if is_eligible_zero else "" 
            elif i == 3: val = row["الكلي"]
            elif i == 4: val = row["مستحق"]
            elif i == 5: 
                val = row["محجوب"]
                font_size = 11  
            elif i == 6: val = row["رقم البطاقة القديم"]
            elif i == 7: 
                val = "محجوب" if is_eligible_zero else "" 
                if is_eligible_zero:
                    text_color = RGBColor(203, 67, 53)
                    font_size = 11
                    
            format_cell_advanced(row_cells[i], val, size_pt=font_size, font_name="Calibri", color_rgb=text_color, align=cell_align)
            
            if is_eligible_zero:
                set_cell_background(row_cells[i], HEX_ALERT_RED)
            else:
                if i == 0: set_cell_background(row_cells[i], HEX_ELEGANT_BLUE)
                elif i == 3: set_cell_background(row_cells[i], HEX_LIGHT_BLUE)   
                elif i == 4: set_cell_background(row_cells[i], HEX_LIGHT_GREEN)  
                elif i == 5: set_cell_background(row_cells[i], HEX_LIGHT_RED)    

    # محرك الإحصاء والمجموع الكلي
    total_all = df["الكلي"].astype(int).sum()
    total_eligible = df["مستحق"].astype(int).sum()
    total_withheld = df["محجوب"].astype(int).sum()
    
    doc.add_paragraph()  
    stats_p = doc.add_paragraph()
    stats_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    stats_p.paragraph_format.element.get_or_add_pPr().append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    
    stats_text = (
        f"العدد الكلي للافراد = {total_all}\n"
        f"العدد الكلي للمستحقين = {total_eligible}\n"
        f"العدد الكلي للمحجوبين = {total_withheld}"
    )
    stats_run = stats_p.add_run(stats_text)
    stats_run.font.name = "Segoe UI Semibold"
    stats_run.font.size = Pt(13)
    stats_run.bold = True
    stats_run.font.color.rgb = COLOR_NAVY_BLUE

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة الاستخدام (Streamlit Interface)
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف المراد تدقيقه وتنسيقه للمطبعة</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("ارفع كشف الوكلاء", type=['docx'], key="doc_input_v9", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename:
        st.session_state.processing_done = False

if st.button("⚙️ تشغيل محرك التنظيم والتنسيق المتقدم الكلي"):
    if uploaded_file:
        with st.spinner('جاري قراءة البيانات الممزقة وإعادة تجميعها بذكاء مع الحفاظ على التسلسل...'):
            try:
                df_res = extract_and_clean_data(uploaded_file)
                if not df_res.empty:
                    st.session_state.df_final = df_res
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.processing_done = True
                else:
                    st.error("لم يتم العثور على بيانات جداول قابلة للقراءة في هذا الملف.")
            except Exception as e:
                st.error(f"خطأ غير متوقع: {e}")
    else:
        st.warning("الرجاء رفع ملف docx أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    output_filename = st.session_state.output_filename
    
    st.success(f"✅ تم معالجة وتنسيق المستند واستعادة الأرقام الأصلية لـ ({len(df_final)}) قيد اسم بنجاح.")
    
    with st.spinner('جاري صياغة وهيكلة مستند Word المطور...'):
        word_output = build_professional_word_report(df_final, output_filename)
        
    st.download_button(
        label="📥 تحميل كشف الوكلاء المنسق والجاهز للطباعة فوراً (Word)",
        data=word_output,
        file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
