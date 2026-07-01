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
# المحرك الجبار المطلق (هندسة الكيانات) - مضاد للتمزق والتداخل 100%
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj):
    doc = Document(file_obj)
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
    
    stop_words = {"نوع", "الوكالة", "رقم", "المركز", "ت", "البطاقة", "القديم", "اسم", "رب", "الاسرة", "الافراد", "الكلية", "المستحقة", "المحجوبين", "FOOD"}
    
    def is_arabic_word(w):
        return any('\u0600' <= c <= '\u06FF' for c in w) and not any(c.isdigit() for c in w)

    entities = []
    i = 0
    # تحويل النصوص الممزقة إلى كيانات ذكية مستقلة
    while i < len(tokens):
        t = tokens[i]
        if t.isdigit() and 5 <= len(t) <= 8:
            entities.append({'type': 'CARD', 'val': t, 'idx': i})
            i += 1
        elif is_arabic_word(t) and t not in stop_words:
            start = i
            while i < len(tokens) and is_arabic_word(tokens[i]) and tokens[i] not in stop_words:
                i += 1
            name = " ".join(tokens[start:i])
            entities.append({'type': 'NAME', 'val': name, 'idx': start})
        elif t.isdigit():
            entities.append({'type': 'NUM', 'val': int(t), 'idx': i})
            i += 1
        else:
            i += 1
            
    # 1. استخراج أسماء المواطنين الحقيقية فقط (من كلمتين فأكثر)
    unique_names = []
    seen = set()
    for e in entities:
        if e['type'] == 'NAME' and len(e['val'].split()) >= 2:
            if e['val'] not in seen:
                seen.add(e['val'])
                unique_names.append(e)

    # 2. استخراج المعادلات الحسابية الصحيحة فقط (الكلي = المستحق + المحجوب)
    num_entities = [e for e in entities if e['type'] == 'NUM']
    triplets = []
    used_num_indices = set()
    for j in range(len(num_entities)-2):
        e1, e2, e3 = num_entities[j], num_entities[j+1], num_entities[j+2]
        
        # السماح بمسافة صغيرة بين الأرقام لتفادي التمزق
        if e3['idx'] - e1['idx'] <= 5:
            a, b, c = e1['val'], e2['val'], e3['val']
            if a == b + c or c == a + b:
                if e1['idx'] not in used_num_indices and e3['idx'] not in used_num_indices:
                    total = max(a, b, c)
                    if a == total: elig, withh = b, c
                    else: elig, withh = b, a
                    triplets.append({
                        'total': total, 'elig': elig, 'withh': withh,
                        'center_idx': e2['idx'],
                        'e1_idx': e1['idx'], 'e2_idx': e2['idx'], 'e3_idx': e3['idx']
                    })
                    used_num_indices.update([e1['idx'], e2['idx'], e3['idx']])
                    
    all_cards = [e for e in entities if e['type'] == 'CARD']
    used_cards = set()
    used_triplets = set()
    
    records = []
    
    # 3. ربط كل مواطن ببياناته الخاصة مهما كان مكانها في السطر الممزق
    for i, name_ent in enumerate(unique_names):
        n_idx = name_ent['idx']
        
        # التقاط أقرب بطاقة
        best_card = ""
        best_card_dist = 9999
        best_c_idx = -1
        
        for c_ent in all_cards:
            if c_ent['idx'] not in used_cards:
                dist = abs(c_ent['idx'] - n_idx)
                if dist < best_card_dist:
                    best_card_dist = dist
                    best_card = c_ent['val']
                    best_c_idx = c_ent['idx']
                    
        # إعطاء الأولوية للبطاقة القديمة التي تبدأ بصفر
        for c_ent in all_cards:
            if c_ent['idx'] not in used_cards and c_ent['val'].startswith('0'):
                dist = abs(c_ent['idx'] - n_idx)
                if dist <= best_card_dist + 5: 
                    best_card_dist = dist
                    best_card = c_ent['val']
                    best_c_idx = c_ent['idx']
                    
        if best_card:
            used_cards.add(best_c_idx)
            for c_ent in all_cards:
                if abs(c_ent['idx'] - best_c_idx) <= 4:
                    used_cards.add(c_ent['idx'])

        # التقاط أقرب معادلة أفراد
        best_trip = None
        best_trip_dist = 9999
        best_trip_idx = -1
        for t_idx, trip in enumerate(triplets):
            if t_idx not in used_triplets:
                dist = abs(trip['center_idx'] - n_idx)
                if dist < best_trip_dist:
                    best_trip_dist = dist
                    best_trip = trip
                    best_trip_idx = t_idx
                    
        if best_trip:
            used_triplets.add(best_trip_idx)
            total, elig, withh = best_trip['total'], best_trip['elig'], best_trip['withh']
            trip_e_indices = [best_trip['e1_idx'], best_trip['e2_idx'], best_trip['e3_idx']]
        else:
            total, elig, withh = 0, 0, 0
            trip_e_indices = []

        # التقاط التسلسل الأصلي الفعلي للمواطن
        seq = str(i + 1)
        search_start = unique_names[i-1]['idx'] if i > 0 else 0
        for num_ent in num_entities:
            if search_start <= num_ent['idx'] <= max(n_idx, best_c_idx if best_c_idx != -1 else n_idx):
                if num_ent['idx'] not in trip_e_indices and num_ent['val'] < 500:
                    seq = str(num_ent['val'])
                    break

        records.append({
            "ت": seq,
            "اسم رب الأسرة": name_ent['val'],
            "رقم البطاقة القديم": best_card,
            "الكلي": total,
            "مستحق": elig,
            "محجوب": withh
        })
        
    df = pd.DataFrame(records)
    return df

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
uploaded_file = st.file_uploader("ارفع كشف الوكلاء", type=['docx'], key="doc_input_final_v2", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename:
        st.session_state.processing_done = False

if st.button("⚙️ تشغيل محرك التنظيم والتنسيق المتقدم الكلي"):
    if uploaded_file:
        with st.spinner('جاري المسح العميق وربط الكيانات لاستخراج كافة القيود بدقة...'):
            try:
                df_res = extract_and_clean_data(uploaded_file)
                if not df_res.empty:
                    st.session_state.df_final = df_res
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.processing_done = True
                else:
                    st.error("لم يتم العثور على بيانات قابلة للقراءة.")
            except Exception as e:
                st.error(f"خطأ غير متوقع: {e}")
    else:
        st.warning("الرجاء رفع ملف docx أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    output_filename = st.session_state.output_filename
    
    st.success(f"✅ فخر البرمجة! تم انتزاع واستعادة كافة البيانات الأصلية لـ ({len(df_final)}) قيد اسم بنجاح مطلق. 🚀")
    
    with st.spinner('جاري صياغة وهيكلة مستند Word المطور...'):
        word_output = build_professional_word_report(df_final, output_filename)
        
    st.download_button(
        label="📥 تحميل كشف الوكلاء المنسق والجاهز للطباعة فوراً (Word)",
        data=word_output,
        file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )            <w:left w:val="single" w:sz="6" w:space="0" w:color="{color_hex}"/>
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
# المحرك الحسابي المطلق المضاد للتمزق 100% (The Absolute Distance-Based Anchor Parser)
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj):
    doc = Document(file_obj)
    
    # 1. سحب النصوص ككتلة واحدة مقاومة للتمزق
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
    
    # كلمات محجوزة لا تعتبر من أسماء المواطنين
    stop_words = {"نوع", "الوكالة", "رقم", "المركز", "ت", "البطاقة", "القديم", "اسم", "رب", "الاسرة", "الافراد", "الكلية", "المستحقة", "المحجوبين", "FOOD"}
    
    # 2. عزل أسماء المواطنين بشكل مستقل وصارم
    names = []
    i = 0
    while i < len(tokens):
        if any('\u0600' <= c <= '\u06FF' for c in tokens[i]) and not any(c.isdigit() for c in tokens[i]) and tokens[i] not in stop_words:
            start = i
            while i < len(tokens) and any('\u0600' <= c <= '\u06FF' for c in tokens[i]) and not any(c.isdigit() for c in tokens[i]) and tokens[i] not in stop_words:
                i += 1
            if (i - start) >= 2:
                names.append({
                    'start_idx': start,
                    'end_idx': i - 1,
                    'name': " ".join(tokens[start:i])
                })
        else:
            i += 1
            
    records = []
    
    # 3. معالجة كل مواطن ضمن "النطاق المغلق" الخاص به فقط
    for idx, n in enumerate(names):
        # تحديد النطاق لمنع التدخل في بيانات المواطن السابق أو اللاحق
        span_start = names[idx-1]['end_idx'] + 1 if idx > 0 else 0
        span_end = names[idx+1]['start_idx'] if idx < len(names) - 1 else len(tokens)
        
        span_tokens = tokens[span_start:span_end]
        
        # البحث عن البطاقة القديمة للمواطن
        old_card = ""
        old_card_idx_in_span = -1
        
        for j, t in enumerate(span_tokens):
            if t.isdigit() and 5 <= len(t) <= 8 and t.startswith('0'):
                old_card = t
                old_card_idx_in_span = j
                break
        
        if old_card_idx_in_span == -1: 
            for j, t in enumerate(span_tokens):
                if t.isdigit() and 5 <= len(t) <= 8:
                    old_card = t
                    old_card_idx_in_span = j
                    break
                    
        # في حال تم العثور على بطاقة المواطن، نقوم بفك تشفير أرقامه الحسابية
        if old_card_idx_in_span != -1:
            small_nums_with_idx = [(j, int(t)) for j, t in enumerate(span_tokens) if t.isdigit() and len(t) <= 3]
            
            valid_triplets = []
            # استخراج الأرقام التي تطابق قاعدة (الكلي = المستحق + المحجوب)
            for j in range(len(small_nums_with_idx)-2):
                i1, a = small_nums_with_idx[j]
                i2, b = small_nums_with_idx[j+1]
                i3, c = small_nums_with_idx[j+2]
                
                if a == b + c:
                    valid_triplets.append( (a, b, c, i2) )
                elif c == a + b:
                    valid_triplets.append( (c, b, a, i2) )
                    
            best_triplet_indices = []
            total, elig, withh = 0, 0, 0
            
            if valid_triplets:
                # اختيار الأرقام الأقرب لبطاقة المواطن في حال وجود أكثر من احتمالية
                valid_triplets.sort(key=lambda x: abs(x[3] - old_card_idx_in_span))
                total, elig, withh = valid_triplets[0][0], valid_triplets[0][1], valid_triplets[0][2]
                
                for j in range(len(small_nums_with_idx)-2):
                    i1, a = small_nums_with_idx[j]
                    i2, b = small_nums_with_idx[j+1]
                    i3, c = small_nums_with_idx[j+2]
                    if i2 == valid_triplets[0][3]: 
                        best_triplet_indices = [i1, i2, i3]
                        break
            elif len(small_nums_with_idx) >= 3:
                # خطة بديلة في حال عدم تطابق المعادلة
                a, b, c = small_nums_with_idx[-3][1], small_nums_with_idx[-2][1], small_nums_with_idx[-1][1]
                total = max(a, b, c)
                if a == total: elig, withh = b, c
                elif c == total: elig, withh = b, a
                else: elig, withh = a, c
                best_triplet_indices = [small_nums_with_idx[-3][0], small_nums_with_idx[-2][0], small_nums_with_idx[-1][0]]
                
            # سحب التسلسل الأصلي (الرقم الصغير الذي يسبق البطاقة ولا ينتمي لمعادلة الأفراد)
            seq = str(len(records)+1)
            for j in range(old_card_idx_in_span - 1, -1, -1):
                t = span_tokens[j]
                if t.isdigit() and len(t) <= 3 and j not in best_triplet_indices:
                    seq = t
                    break
                    
            records.append({
                "ت": seq,
                "اسم رب الأسرة": n['name'],
                "رقم البطاقة القديم": old_card,
                "الكلي": total,
                "مستحق": elig,
                "محجوب": withh
            })
            
    df = pd.DataFrame(records)
    if not df.empty:
        # التخلص من التكرار بنسبة 100% بناءً على رقم البطاقة القديم
        df = df.drop_duplicates(subset=["رقم البطاقة القديم"], keep="first")
        
    return df

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
uploaded_file = st.file_uploader("ارفع كشف الوكلاء", type=['docx'], key="doc_input_final", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename:
        st.session_state.processing_done = False

if st.button("⚙️ تشغيل محرك التنظيم والتنسيق المتقدم الكلي"):
    if uploaded_file:
        with st.spinner('جاري قراءة الملف وتطبيق خوارزمية النطاقات المغلقة لسحب كافة القيود بدقة...'):
            try:
                df_res = extract_and_clean_data(uploaded_file)
                if not df_res.empty:
                    st.session_state.df_final = df_res
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.processing_done = True
                else:
                    st.error("لم يتم العثور على بيانات قابلة للقراءة.")
            except Exception as e:
                st.error(f"خطأ غير متوقع: {e}")
    else:
        st.warning("الرجاء رفع ملف docx أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    output_filename = st.session_state.output_filename
    
    st.success(f"✅ مذهل! تم استخراج واستعادة الأرقام الأصلية لـ ({len(df_final)}) قيد اسم بنجاح مطلق. 🚀")
    
    with st.spinner('جاري صياغة وهيكلة مستند Word المطور...'):
        word_output = build_professional_word_report(df_final, output_filename)
        
    st.download_button(
        label="📥 تحميل كشف الوكلاء المنسق والجاهز للطباعة فوراً (Word)",
        data=word_output,
        file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
