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
# إعدادات الواجهة (The Console)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام كشوفات الوكلاء الخارق", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #1A5276; color: white; width: 100%; font-weight: bold; border-radius: 8px; font-size: 18px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);}
    div.stButton > button:hover { background-color: #154360; border-color: #F1C40F; color: #F1C40F;}
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
    st.session_state.total_scanned = 0

# -----------------------------------------------------------------------------
# دوال التنسيق المتقدمة لـ Word (The Forges)
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
# المحرك الخارق: خوارزمية الرياضيات الموجهة (Directional Math AI)
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj, card_choice):
    doc = Document(file_obj)
    raw_records = []
    
    # 1. التجميع الشامل (حصد كل نصوص المستند وتدمير الجداول الوهمية)
    lines = []
    for para in doc.paragraphs:
        if para.text.strip(): lines.append(para.text.strip())
            
    for table in doc.tables:
        for row in table.rows:
            # دمج الخلايا أفقياً لسحق تشوهات المطبعة
            cells_text = [re.sub(r'\s+', ' ', cell.text.strip()) for cell in row.cells if cell.text.strip()]
            lines.append(" ".join(cells_text))
            
    total_scanned_counter = 0

    # 2. الاختراق اللغوي والرياضي
    for line in lines:
        # توحيد الرموز الغريبة
        line_clean = re.sub(r'[،,\-\_]', ' ', line)
        line_clean = re.sub(r'\s+', ' ', line_clean).strip()
        
        if not line_clean: continue
        
        # تخطي الترويسات التي تخدع الكود
        if any(w in line_clean for w in ["المركز", "نوع الوكالة", "الافراد الكلية", "اسم رب", "رقم البطاقة"]): continue
        
        # أ- صيد الاسم (أقوى رادار للأسماء العربية)
        arabic_words = re.findall(r'[\u0600-\u06FF]{2,}', line_clean)
        stop_words = ["ملاحظات", "الوكيل", "الافراد", "المحجوبين", "المستحقة", "رقم", "البطاقة", "القديم", "الحديث", "ت", "حقل", "فارغ"]
        name_candidates = [w for w in arabic_words if w not in stop_words]
        
        if len(name_candidates) < 2: continue # يجب أن يكون الاسم مقطعين على الأقل
            
        # حماية الهوية العراقية (عبد، أبو، الـ)
        prefixes = ["عبد", "ابو", "أبو", "ام", "أم", "آل", "ال", "بني", "ذو"]
        final_parts = []
        real_words = 0
        for p in name_candidates:
            final_parts.append(p)
            if p not in prefixes: real_words += 1
            if real_words >= 4: break # نأخذ 4 مقاطع لضمان الأسماء الطويلة
        three_part_name = " ".join(final_parts)
        
        # ب- فصل الأرقام عن النص
        line_without_name = line_clean
        for word in name_candidates:
            line_without_name = line_without_name.replace(word, ' ')
            
        all_nums = [int(n) for n in re.findall(r'\d+', line_without_name)]
        
        cards = [n for n in all_nums if n > 10000]
        smalls = [n for n in all_nums if n < 100] # العوائل والتسلسلات
        
        if not cards: 
            selected_card = "غير متوفر" 
        else:
            old_card = str(cards[0])
            new_card = str(cards[-1]) if len(cards) >= 2 else old_card
            selected_card = new_card if card_choice == "رقم البطاقة الحديث" else old_card
        
        # ج- الرياضيات الموجهة (Directional Math Algorithm)
        total = eligible = withheld = 0
        found_math = False
        
        # نبحث في كل 3 أرقام متجاورة عن معادلة منطقية
        for i in range(len(smalls) - 2):
            window = smalls[i:i+3]
            a, b, c = window
            
            # النسق الأول (كلي، مستحق، محجوب) -> أ = ب + ج
            if a == b + c:
                total, eligible, withheld = a, b, c
                found_math = True; break
            
            # النسق الثاني (محجوب، مستحق، كلي) -> ج = أ + ب
            elif c == a + b:
                total, eligible, withheld = c, b, a
                found_math = True; break
                
        # وضع طوارئ (إذا كان المحجوب 0 أحياناً لا يكتبونه)
        if not found_math and len(smalls) >= 2:
            safe_smalls = [n for n in smalls if n <= 20] # تجنب أرقام التسلسل الكبيرة
            if len(safe_smalls) >= 2:
                total = max(safe_smalls)
                safe_smalls.remove(total)
                eligible = max(safe_smalls)
                withheld = total - eligible
            elif len(safe_smalls) == 1:
                total = eligible = safe_smalls[0]
                withheld = 0

        total_scanned_counter += 1

        raw_records.append({
            "اسم رب الأسرة": three_part_name,
            "رقم البطاقة": selected_card,
            "الكلي": total,
            "محجوب": withheld,
            "مستحق": eligible
        })
        
    df = pd.DataFrame(raw_records)
    df_duplicates = pd.DataFrame()
    
    if not df.empty:
        # 1. استخراج المتكررات الحرفية وإرسالها للسجن (ملف العزل)
        duplicate_mask = df.duplicated(subset=["اسم رب الأسرة", "رقم البطاقة"], keep='first')
        df_duplicates = df[duplicate_mask].copy()
        
        if not df_duplicates.empty:
            df_duplicates = df_duplicates.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
            df_duplicates.insert(0, "ت", df_duplicates.index + 1)

        # 2. تنقية الكشف الأصلي من التكرار وإعداده للطباعة
        df = df.drop_duplicates(subset=["اسم رب الأسرة", "رقم البطاقة"], keep='first')
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
        
    return df, df_duplicates, total_scanned_counter

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word (The Document Forge)
# -----------------------------------------------------------------------------
def build_professional_word_report(df, filename_base, card_choice, is_duplicate_report=False):
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
    
    report_title = f"كشف المتكررات المعزولة: {clean_name}" if is_duplicate_report else f"الكشف الإحصائي المنسق للوكيل: {clean_name}"
    
    title_run = title_p.add_run(report_title)
    title_run.font.name = "Segoe UI Semibold"
    title_run.font.size = Pt(14)
    title_run.bold = True
    if is_duplicate_report:
        title_run.font.color.rgb = RGBColor(220, 20, 60) # أحمر للمتكررات
    else:
        title_run.font.color.rgb = RGBColor(26, 82, 118)
    
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
# قمرة القيادة (Streamlit Console)
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف المراد تدقيقه واختراقه</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("ارفع كشف الوكلاء", type=['docx'], key="doc_input_master", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)
selected_card = st.radio("اختر نوع رقم البطاقة المراد اعتماده في الكشف:", ["رقم البطاقة القديم", "رقم البطاقة الحديث"], index=0, horizontal=True)
st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename or st.session_state.selected_card != selected_card:
        st.session_state.processing_done = False

if st.button("🚀 إطلاق محرك الرياضيات الموجهة (حصد كل البيانات)"):
    if uploaded_file:
        with st.spinner('يتم الآن سحق الجداول وتحليل الشيفرات بالذكاء الرياضي الباليستي...'):
            try:
                df_res, df_dup, total_scanned = extract_and_clean_data(uploaded_file, selected_card)
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
        st.warning("الرجاء رفع ملف docx أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    df_duplicates = st.session_state.df_duplicates
    total_found = st.session_state.total_scanned
    output_filename = st.session_state.output_filename
    used_card_type = st.session_state.selected_card
    
    st.balloons()
    st.success(f"🏆 المهمة أُنجزت بقوة ساحقة! المحرك الخفي قام بمسح ({total_found}) قيد من أعماق الملف الممزق.\nبعد تدمير التكرارات، صفا لك ({len(df_final)}) قيد اسم صافٍ ونقي!")
    
    # تحميل الكشف النظيف
    with st.spinner('جاري صياغة الكشف الذهبي الأساسي...'):
        word_output = build_professional_word_report(df_final, output_filename, used_card_type, is_duplicate_report=False)
        
    st.download_button(
        label=f"📥 تحميل الكشف الأساسي ({len(df_final)} اسم) جاهز للطباعة",
        data=word_output,
        file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    
    # تحميل كشف العزل للمتكررات
    if not df_duplicates.empty:
        st.warning(f"⚠️ لقد كشفنا المستور! تم اصطياد وعزل ({len(df_duplicates)}) قيد متكرر كان يختبئ في الملف.")
        with st.spinner('جاري بناء سجن المتكررات...'):
            word_output_dup = build_professional_word_report(df_duplicates, output_filename, used_card_type, is_duplicate_report=True)
            
        st.download_button(
            label="🚨 تحميل ملف القيود المتكررة المعزولة (للتدقيق)",
            data=word_output_dup,
            file_name=f"سجلات_متكررة_معزولة_{output_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
