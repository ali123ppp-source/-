# import streamlit as st.txt  -- نسخة نهائية مدموجة وشاملة (Word + Excel/CSV) معالجة الجذور والأسطر المبعثرة
import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import re
import csv
import traceback

# -------------------------------------------------------------------------
# إعدادات الواجهة
# -------------------------------------------------------------------------
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
    st.session_state.failed_lines = []
    st.session_state.merge_groups = []
    st.session_state.stats = {}

# -------------------------------------------------------------------------
# دوال تنسيق Word
# -------------------------------------------------------------------------
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

# -------------------------------------------------------------------------
# دالة تحليل DOCX الجبرية
# -------------------------------------------------------------------------
def _parse_docx_to_list(file_obj, card_choice, stop_on_error=False, debug_limit=0):
    doc = Document(file_obj)
    parsed_records = []
    failed_lines = []
    raw_lines = []
    merge_groups = []

    for p in doc._element.xpath('.//w:p'):
        txt = "".join(node.text for node in p.xpath('.//w:t') if node.text).strip()
        if txt: raw_lines.append(txt)

    final_records = []
    current_name, current_cards, current_smalls = [], [], []

    for line in raw_lines:
        line_clean = line.replace(',', ' ').replace('،', ' ').replace('-', ' ')
        cards_in_line = re.findall(r'\b\d{4,}\b', line_clean)
        smalls_in_line = [int(n) for n in re.findall(r'\b\d{1,3}\b', line_clean)]
        arabic_in_line = re.findall(r'[\u0600-\u06FF]{2,}', line_clean)
        
        if any(w in line_clean for w in ["المركز", "نوع الوكالة", "الافراد الكلية", "ت", "البطاقة"]):
            continue

        if arabic_in_line and cards_in_line:
            if current_name:
                final_records.append({"name": " ".join(current_name), "cards": current_cards, "smalls": current_smalls, "raw": "مقطع مجمع"})
                merge_groups.append((("تجميع", "ذكي"), " ".join(current_name), "سجل مكتمل"))
                current_name, current_cards, current_smalls = [], [], []
            final_records.append({"name": " ".join(arabic_in_line), "cards": cards_in_line, "smalls": smalls_in_line, "raw": line})
        else:
            if arabic_in_line:
                is_new = False
                if current_name:
                    if current_cards or (len(current_smalls) >= 3 and len(arabic_in_line) >= 2) or (len(current_name) >= 3 and len(arabic_in_line) >= 3):
                        is_new = True
                if is_new:
                    final_records.append({"name": " ".join(current_name), "cards": current_cards, "smalls": current_smalls, "raw": "مقطع مجمع"})
                    merge_groups.append((("تجميع", "ذكي"), " ".join(current_name), "سجل مكتمل"))
                    current_name, current_cards, current_smalls = [], [], []
                current_name.extend(arabic_in_line)
            if cards_in_line: current_cards.extend(cards_in_line)
            if smalls_in_line: current_smalls.extend(smalls_in_line)

    if current_name:
        final_records.append({"name": " ".join(current_name), "cards": current_cards, "smalls": current_smalls, "raw": "مقطع نهائي"})

    for record in final_records:
        try:
            name = record["name"].strip()
            cards = record["cards"]
            smalls = record["smalls"]

            if not name: continue
            if not cards: cards = ["غير متوفر"]

            old_card = cards[0]
            new_card = cards[-1] if len(cards) >= 2 else old_card
            selected_card = new_card if card_choice == "رقم البطاقة الحديث" else old_card

            total = eligible = withheld = 0
            found_counts = False

            for i in range(len(smalls) - 2):
                triplet = smalls[i:i+3]
                if triplet[0] == triplet[1] + triplet[2] and triplet[0] > 0:
                    total, eligible, withheld = triplet[0], triplet[1], triplet[2]
                    found_counts = True; break
                if triplet[2] == triplet[0] + triplet[1] and triplet[2] > 0:
                    withheld, eligible, total = triplet[0], triplet[1], triplet[2]
                    found_counts = True; break

            if not found_counts:
                if len(smalls) >= 3: total, eligible, withheld = smalls[-3], smalls[-2], smalls[-1]
                elif len(smalls) == 2: total, eligible, withheld = max(smalls), min(smalls), 0
                elif len(smalls) == 1: total, eligible, withheld = smalls[0], 0, 0

            if any(w in name for w in ["اجمالي", "المجموع", "مجموع"]): continue

            is_reversed = False
            for pr in parsed_records[-15:]:
                if pr["اسم رب الأسرة"] == name[::-1]:
                    is_reversed = True; break
            if is_reversed: continue

            parsed_records.append({"اسم رب الأسرة": name, "رقم البطاقة": selected_card, "الكلي": total, "محجوب": withheld, "مستحق": eligible})
        except Exception as e:
            failed_lines.append((record.get("raw", ""), f"exception:{str(e)}"))

    return parsed_records, failed_lines, len(raw_lines), merge_groups

# -------------------------------------------------------------------------
# وحدة استخراج Excel / CSV (الجديدة والمطورة)
# -------------------------------------------------------------------------
def _parse_excel_to_list(file_obj, file_ext, card_choice, stop_on_error=False, debug_limit=0):
    records, failed_lines, raw_count, merge_groups = [], [], 0, []
    try:
        if file_ext == 'xlsx': df_in = pd.read_excel(file_obj, header=None)
        else: df_in = pd.read_csv(file_obj, header=None)
        
        raw_count = len(df_in)
        
        # البحث الديناميكي عن سطر العناوين (الهيدر)
        header_idx = 0
        for i, row in df_in.iterrows():
            row_str = ' '.join(str(x) for x in row.values)
            if 'اسم' in row_str and ('بطاق' in row_str or 'كلي' in row_str):
                header_idx = i
                break
        
        df_in.columns = df_in.iloc[header_idx].astype(str).str.strip()
        df_in = df_in.iloc[header_idx+1:].reset_index(drop=True)
        cols = df_in.columns
        
        col_name = next((c for c in cols if 'اسم' in c), None)
        
        col_new_card, col_old_card = None, None
        for c in cols:
            if c == 'رقم البطاقة': col_new_card = c
            if 'قديم' in c: col_old_card = c
        
        if not col_new_card:
            card_cols = [c for c in cols if 'بطاق' in c]
            if card_cols:
                col_new_card = card_cols[-1] if 'قديم' not in card_cols[-1] else card_cols[0]
                col_old_card = card_cols[0] if len(card_cols) > 1 else col_new_card
                
        col_total = next((c for c in cols if 'كلي' in c), None)
        col_eligible = next((c for c in cols if 'مستحق' in c), None)
        col_withheld = next((c for c in cols if 'محجوب' in c), None)

        for _, row in df_in.iterrows():
            try:
                if pd.isna(row[col_name]): continue
                
                # علاج فواصل الأسطر داخل الخلايا Excel
                name = str(row[col_name]).strip()
                if name in ['nan', 'None'] or not name or 'اسم' in name or 'اجمالي' in name or 'مجموع' in name:
                    continue
                name = name.replace('\n', ' ').replace('\r', ' ').replace('  ', ' ')
                
                c_new = str(row[col_new_card]).strip() if col_new_card else "غير متوفر"
                c_old = str(row[col_old_card]).strip() if col_old_card else "غير متوفر"
                c_new = "غير متوفر" if c_new == 'nan' else c_new
                c_old = "غير متوفر" if c_old == 'nan' else c_old
                selected_card = c_new if card_choice == "رقم البطاقة الحديث" else c_old

                def safe_int(val):
                    try: return int(float(val)) if not pd.isna(val) else 0
                    except: return 0

                total = safe_int(row[col_total]) if col_total else 0
                eligible = safe_int(row[col_eligible]) if col_eligible else 0
                withheld = safe_int(row[col_withheld]) if col_withheld else 0
                
                if total == 0 and (eligible > 0 or withheld > 0):
                    total = eligible + withheld
                    
                records.append({"اسم رب الأسرة": name, "رقم البطاقة": selected_card, "الكلي": total, "محجوب": withheld, "مستحق": eligible})
            except Exception as e:
                failed_lines.append((str(row.values), f"row_error: {e}"))
                
    except Exception as e:
        failed_lines.append(("Excel Parsing", f"fatal_error: {e}"))
        
    return records, failed_lines, raw_count, merge_groups

# -------------------------------------------------------------------------
# الماستر: الدمج والفرز
# -------------------------------------------------------------------------
def extract_and_clean_data(file_obj, file_ext, card_choice, file_obj2=None, file2_ext=None, stop_on_error=False):
    if file_ext == 'docx':
        parsed_primary, failed_primary, raw_count_primary, merge_groups_primary = _parse_docx_to_list(file_obj, card_choice, stop_on_error)
    else:
        parsed_primary, failed_primary, raw_count_primary, merge_groups_primary = _parse_excel_to_list(file_obj, file_ext, card_choice, stop_on_error)
    
    parsed_secondary, failed_secondary, raw_count_secondary, merge_groups_secondary = [], [], 0, []
    if file_obj2:
        if file2_ext == 'docx':
            parsed_secondary, failed_secondary, raw_count_secondary, merge_groups_secondary = _parse_docx_to_list(file_obj2, card_choice, stop_on_error)
        elif file2_ext in ['xlsx', 'csv']:
            parsed_secondary, failed_secondary, raw_count_secondary, merge_groups_secondary = _parse_excel_to_list(file_obj2, file2_ext, card_choice, stop_on_error)

    all_parsed = parsed_primary + parsed_secondary
    all_failed = failed_primary + failed_secondary
    all_merge_groups = merge_groups_primary + merge_groups_secondary

    df = pd.DataFrame(all_parsed)
    df_duplicates = pd.DataFrame()
    
    if not df.empty:
        duplicate_mask = df.duplicated(keep='first')
        df_duplicates = df[duplicate_mask].copy()
        if not df_duplicates.empty:
            df_duplicates = df_duplicates.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
            df_duplicates.insert(0, "ت", df_duplicates.index + 1)

        df = df.drop_duplicates(keep='first')
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
        
    stats = {
        "raw_lines_primary": raw_count_primary, "raw_lines_secondary": raw_count_secondary,
        "parsed_count": len(all_parsed), "failed_count": len(all_failed),
        "duplicates_count": len(df_duplicates), "merge_groups_count": len(all_merge_groups)
    }
    return df, df_duplicates, all_failed, all_merge_groups, stats

# -------------------------------------------------------------------------
# محرك التقرير (Word)
# -------------------------------------------------------------------------
def build_professional_word_report(df, filename_base, card_choice, is_duplicate_report=False):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Cm(0.5), Cm(0.5)
        section.left_margin, section.right_margin = Cm(0.3), Cm(0.3)
        
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز"]: clean_name = clean_name.replace(w, "")
    clean_name = re.sub(r'[a-zA-Z]', '', clean_name)
    clean_name = re.sub(r'[\-_+_.]', '', clean_name)
    clean_name = " ".join(clean_name.split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"كشف المتكررات المعزولة: {clean_name}" if is_duplicate_report else f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    title_run.font.color.rgb = RGBColor(220, 20, 60) if is_duplicate_report else RGBColor(26, 82, 118)
    
    footer_p = doc.sections[0].footer.paragraphs[0]
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
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="DC143C" if is_duplicate_report else "1A5276")
    
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) if not df.empty else 15
    col_widths = [Cm(0.9), Cm(max_name_len * 0.22 + 0.5), Cm(0.44), Cm(0.9), Cm(0.9), Cm(0.9), Cm(3.0), Cm(2.19)]
    THEME_COLOR = RGBColor(220, 20, 60) if is_duplicate_report else RGBColor(26, 82, 118)
    
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        if i in [3, 4, 5]:
            set_cell_vertical_text(hdr_cells[i])
            format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=12, font_name="Segoe UI Semibold", color_rgb=THEME_COLOR)
        else:
            format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i == 1 else "center", color_rgb=THEME_COLOR)
            
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        table.rows[idx+1]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0
        
        for i in range(8): row_cells[i].width = col_widths[i]
        set_cell_no_wrap(row_cells[1])
        
        for i in range(8):
            val, text_color, cell_align, font_size = "", None, "center", 16
            if i == 0: val = row["ت"]
            elif i == 1: val, cell_align = row["اسم رب الأسرة"], "left" 
            elif i == 2: val = "x" if is_eligible_zero else "" 
            elif i == 3: val = row["الكلي"]
            elif i == 4: val = row["مستحق"]
            elif i == 5: val, font_size = row["محجوب"], 14  
            elif i == 6: val = row["رقم البطاقة"] 
            elif i == 7: 
                val = "محجوب" if is_eligible_zero else "" 
                if is_eligible_zero: text_color, font_size = RGBColor(203, 67, 53), 12  
                    
            format_cell_advanced(row_cells[i], val, size_pt=font_size, color_rgb=text_color, align=cell_align)
            
            if is_eligible_zero: set_cell_background(row_cells[i], "EC7063")
            else:
                if i == 0: set_cell_background(row_cells[i], "D4E6F1")
                elif i == 3: set_cell_background(row_cells[i], "EBF5FB")   
                elif i == 4: set_cell_background(row_cells[i], "E8F8F5")  
                elif i == 5: set_cell_background(row_cells[i], "FADBD8")    

    doc.add_paragraph()  
    stats_p = doc.add_paragraph()
    stats_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    stats_p.paragraph_format.element.get_or_add_pPr().append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    
    total_all, total_eligible, total_withheld = (df[c].astype(int).sum() if not df.empty else 0 for c in ["الكلي", "مستحق", "محجوب"])
    stats_run = stats_p.add_run(f"العدد الكلي للافراد = {total_all}\nالعدد الكلي للمستحقين = {total_eligible}\nالعدد الكلي للمحجوبين = {total_withheld}")
    stats_run.font.name, stats_run.font.size, stats_run.bold, stats_run.font.color.rgb = "Segoe UI Semibold", Pt(13), True, THEME_COLOR

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -------------------------------------------------------------------------
# واجهة الاستخدام (محدثة لدعم Excel)
# -------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col2:
    st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف الأساسي (Word أو Excel)</h3>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("ارفع كشف الوكلاء الأساسي", type=['docx', 'xlsx', 'csv'], key="doc_input_master", label_visibility="collapsed")

with col1:
    st.markdown("<h3 style='text-align: right;'>➕ ملف للدمج (اختياري)</h3>", unsafe_allow_html=True)
    uploaded_file2 = st.file_uploader("ارفع جدول للدمج", type=['docx', 'xlsx', 'csv'], key="doc_input_merge", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)
selected_card = st.radio("اختر نوع رقم البطاقة المراد اعتماده:", ["رقم البطاقة القديم", "رقم البطاقة الحديث"], index=0, horizontal=True)
st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename or st.session_state.selected_card != selected_card:
        st.session_state.processing_done = False

if st.button("🚀 تشغيل المحرك (استخراج ودمج وفرز)"):
    if uploaded_file:
        with st.spinner('يتم الآن سحق الجداول وتحليل الشيفرات...'):
            try:
                ext1 = uploaded_file.name.split('.')[-1].lower()
                ext2 = uploaded_file2.name.split('.')[-1].lower() if uploaded_file2 else None
                
                df_res, df_dup, failed_lines, merge_groups, stats = extract_and_clean_data(uploaded_file, ext1, selected_card, uploaded_file2, ext2)
                
                st.session_state.df_final = df_res
                st.session_state.df_duplicates = df_dup
                st.session_state.failed_lines = failed_lines
                st.session_state.merge_groups = merge_groups
                st.session_state.stats = stats
                st.session_state.output_filename = current_filename
                st.session_state.selected_card = selected_card
                st.session_state.processing_done = True

                if df_res.empty: st.error("لم يتم العثور على بيانات. يرجى التأكد من الملف.")
                else: st.success(f"تم الاستخراج بنجاح: {len(df_res)} سجلّاً.")
            except Exception as e:
                st.error(f"خطأ: {e}")
                st.error(traceback.format_exc())
    else: st.warning("الرجاء رفع الملف أولاً.")

# -------------------------------------------------------------------------
# التنزيلات
# -------------------------------------------------------------------------
if st.session_state.processing_done:
    df_final, df_duplicates, output_filename = st.session_state.df_final, st.session_state.df_duplicates, st.session_state.output_filename
    
    st.balloons()
    st.success(f"🏆 المهمة أُنجزت! تم استخراج وترتيب ({len(df_final)}) قيد صافٍ.")

    with st.spinner('جاري بناء ملف Word...'):
        word_output = build_professional_word_report(df_final, output_filename, st.session_state.selected_card)
        
    st.download_button(label=f"📥 تحميل الكشف الأساسي المنسق", data=word_output, file_name=f"كشف_منسق_{output_filename}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    
    if not df_duplicates.empty:
        st.warning(f"⚠️ تم عزل ({len(df_duplicates)}) قيد مكرر.")
        with st.spinner('جاري بناء ملف التكرارات...'):
            word_output_dup = build_professional_word_report(df_duplicates, output_filename, st.session_state.selected_card, True)
        st.download_button(label="🚨 تحميل ملف التكرارات المعزولة", data=word_output_dup, file_name=f"تكرارات_{output_filename}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
