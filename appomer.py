# import streamlit as st.txt  -- نسخة نهائية مدموجة مع دالة دمج الأسطر المبعثرة
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
    st.session_state.stats = {}

# -------------------------------------------------------------------------
# دوال تنسيق Word (ثابتة)
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
# دالة تصنيف سبب الفشل (لوج)
# -------------------------------------------------------------------------
def classify_failure_reason(line):
    """تعطي سبب محتمل لفشل استخراج السطر"""
    if not line or not line.strip():
        return "empty_line"
    if '<td' in line.lower() or '<table' in line.lower():
        return "html_table_unparsed"
    if not re.search(r'[\u0600-\u06FF]{2,}', line):
        return "no_arabic_name"
    if len(re.findall(r'\b\d{1,}\b', line)) < 1:
        return "no_numbers"
    return "ambiguous_format"

# -------------------------------------------------------------------------
# دالة مساعدة: توليد خلايا فريدة لتفادي مشاكل الخلايا المدموجة
# -------------------------------------------------------------------------
def row_cells_unique(row):
    """يتجاهل الخلايا المكررة الناتجة عن merged cells أفقياً"""
    seen = set()
    for cell in row.cells:
        tc = cell._tc
        if tc in seen:
            continue
        seen.add(tc)
        yield cell

# -------------------------------------------------------------------------
# دالة لدمج الأسطر المبعثرة قبل التحليل
# -------------------------------------------------------------------------
def merge_fragmented_lines(raw_lines, max_group=3):
    """
    يدمج الأسطر المبعثرة التي تحتوي أرقاماً فقط أو أجزاء سجلات.
    - raw_lines: قائمة الأسطر الخام (قائمة سترينج)
    - max_group: الحد الأقصى لعدد الأسطر التي نسمح بدمجها معاً
    يعيد قائمة جديدة من الأسطر المدمجة.
    """
    merged = []
    i = 0
    n = len(raw_lines)
    arabic_re = re.compile(r'[\u0600-\u06FF]')
    numeric_only_re = re.compile(r'^\s*(?:\d+\s*)+$')  # سطر يحتوي أرقام فقط
    short_nums_re = re.compile(r'^\s*(?:\d{1,3}\s+){1,}\d{1,3}\s*$')  # أرقام قصيرة 1-3 خانات
    while i < n:
        line = raw_lines[i].strip()
        # حالة: سطر أرقام فقط → حاول دمجه مع السابق أو التالي الذي يحتوي اسم عربي
        if numeric_only_re.match(line) or (short_nums_re.match(line) and not arabic_re.search(line)):
            # نفضّل الدمج مع السطر السابق إذا يحتوي اسم عربي
            if merged and arabic_re.search(merged[-1]):
                merged[-1] = merged[-1] + " " + line
                i += 1
                continue
            # وإلا ندمجه مع التالي إذا التالي يحتوي اسم عربي
            if i + 1 < n and arabic_re.search(raw_lines[i+1]):
                combined = raw_lines[i+1].strip()
                merged.append(line + " " + combined)
                i += 2
                continue
            # وإلا نحتفظ به كخط مستقل (قد يكون ملخصاً)
            merged.append(line)
            i += 1
            continue

        # حالة: السطر يبدأ بأرقام/مؤشر ثم لا يحتوي اسم واضح، والصف التالي يبدأ باسم → ادمجهما
        parts = re.split(r'\||\t', line)
        has_arabic = any(arabic_re.search(p) for p in parts)
        if not has_arabic and i + 1 < n and arabic_re.search(raw_lines[i+1]):
            merged.append(line + " " + raw_lines[i+1].strip())
            i += 2
            continue

        # حالة: السطر يحتوي اسم جزئي (قصير) والأرقام موزعة على السطر التالي → دمج إذا التالي أرقام
        if arabic_re.search(line) and i + 1 < n and numeric_only_re.match(raw_lines[i+1].strip()):
            merged.append(line + " " + raw_lines[i+1].strip())
            i += 2
            continue

        # حالة عامة: لا دمج
        merged.append(line)
        i += 1

    # خطوة ثانية: محاولة دمج مجموعات صغيرة متبقية (حد أقصى max_group)
    final = []
    i = 0
    n = len(merged)
    while i < n:
        group = [merged[i]]
        j = i + 1
        while j < n and len(group) < max_group:
            # إذا المجموعة الحالية لا تحتوي اسم عربي لكن العنصر التالي يحتوي اسمًا، ندمج
            if not re.search(r'[\u0600-\u06FF]', " ".join(group)) and re.search(r'[\u0600-\u06FF]', merged[j]):
                group.append(merged[j])
                j += 1
                break
            # إذا العنصر التالي عبارة عن أرقام فقط، نضمّه
            if numeric_only_re.match(merged[j]):
                group.append(merged[j])
                j += 1
                continue
            break
        final.append(" ".join(group))
        i = j
    return final

# -------------------------------------------------------------------------
# دالة تحليل DOCX محسّنة (النسخة النهائية المدموجة)
# -------------------------------------------------------------------------
def _parse_docx_to_list(file_obj, card_choice, stop_on_error=False, debug_limit=0):
    """
    نسخة محسّنة نهائية:
    - تتعامل مع merged cells أفقياً وعمودياً (vMerge)
    - تحلل HTML-like tables داخل الفقرات
    - تجمع فقرات مبعثرة
    - تدمج الأسطر المبعثرة قبل التحليل
    - ترجع parsed_records, failed_lines, raw_lines_count
    """
    doc = Document(file_obj)
    parsed_records = []
    failed_lines = []
    raw_lines = []

    def parse_html_table_text(text):
        tds = re.findall(r'<td[^>]*>(.*?)</td>', text, flags=re.DOTALL | re.IGNORECASE)
        tds = [re.sub(r'<.*?>', '', td).replace('\n', ' ').strip() for td in tds]
        rows = []
        if not tds:
            return rows
        for cols in (8,7,6,5,4):
            if len(tds) % cols == 0:
                for i in range(0, len(tds), cols):
                    rows.append(' | '.join(tds[i:i+cols]))
                return rows
        rows.append(' | '.join(tds))
        return rows

    def classify_failure_reason_local(line):
        if not line or not line.strip(): return "empty_line"
        if '<td' in line.lower() or '<table' in line.lower(): return "html_table_unparsed"
        if not re.search(r'[\u0600-\u06FF]{2,}', line): return "no_arabic_name"
        if len(re.findall(r'\b\d{1,}\b', line)) < 1: return "no_numbers"
        return "ambiguous_format"

    # --- 1) استخراج صفوف الجداول الحقيقية مع معالجة vMerge عمودي و merged أفقياً ---
    for table in doc.tables:
        prev_row_texts = []
        for row in table.rows:
            cells = []
            seen = set()
            for cell in row.cells:
                tc = cell._tc
                if tc in seen:
                    continue
                seen.add(tc)
                # فحص vMerge داخل خصائص الخلية
                vmerge = None
                try:
                    tcPr = tc.tcPr
                    if tcPr is not None:
                        vm = tcPr.find(qn('w:vMerge'))
                        if vm is not None:
                            vmerge = vm.get(qn('w:val')) if vm.get(qn('w:val')) is not None else "continue"
                except Exception:
                    vmerge = None
                text = cell.text.replace('\n', ' ').strip()
                cells.append((text, vmerge))
            # حل مشكلة vMerge عمودي
            resolved = []
            for i, (txt, vmerge) in enumerate(cells):
                if vmerge and str(vmerge).lower() == 'continue':
                    if i < len(prev_row_texts):
                        resolved.append(prev_row_texts[i])
                    else:
                        resolved.append(txt)
                else:
                    resolved.append(txt)
            prev_row_texts = resolved[:]  # للاستخدام في الصف التالي
            line = " | ".join([r for r in resolved])
            if line.strip():
                raw_lines.append(line)

    # --- 2) فحص الفقرات (نحلل HTML-like أو نجمع فقرات قصيرة) ---
    buffer_para = ""
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            if buffer_para:
                raw_lines.append(buffer_para.strip())
                buffer_para = ""
            continue
        low = txt.lower()
        if '<table' in low or '<td' in low:
            rows = parse_html_table_text(txt)
            if rows:
                raw_lines.extend(rows)
            else:
                raw_lines.append(txt)
            continue
        if len(txt.split()) < 4:
            buffer_para += " " + txt
        else:
            if buffer_para:
                combined = (buffer_para + " " + txt).strip()
                raw_lines.append(combined)
                buffer_para = ""
            else:
                raw_lines.append(txt)
    if buffer_para:
        raw_lines.append(buffer_para.strip())

    # --- 2.5) دمج الأسطر المبعثرة قبل التحليل ---
    raw_lines = merge_fragmented_lines(raw_lines)

    # --- 3) تحليل كل raw_line لاستخراج الاسم والأرقام مع قواعد مرنة ---
    header_keywords = ["المركز", "اسم رب", "ملاحظات", "الوكيل", "الافراد", "الكلية", "المحجوبين", "FOOD", "نوع الوكالة", "ت رقم البطاقة", "اجمالي", "المجموع"]
    for idx, line in enumerate(raw_lines):
        try:
            line_clean = line.replace(',', ' ').replace('،', ' ').replace('-', ' ')
            if any(re.search(r'^\s*' + re.escape(w) + r'\b', line_clean, flags=re.IGNORECASE) for w in header_keywords):
                if not (re.search(r'[\u0600-\u06FF]{2,}', line_clean) and re.search(r'\d', line_clean)):
                    continue

            all_nums = re.findall(r'\b0*\d{4,}\b', line_clean)
            smalls_all = re.findall(r'\b\d{1,3}\b', line_clean)

            name_matches = re.findall(r'[\u0600-\u06FF]{2,}(?:\s+[\u0600-\u06FF]{2,})*', line_clean)
            name = None
            if name_matches:
                name = max(name_matches, key=len).strip()

            if not name:
                parts = [p.strip() for p in re.split(r'\||\t', line_clean) if p.strip()]
                for p in parts:
                    if re.search(r'[\u0600-\u06FF]{2,}', p):
                        name = p
                        break

            if not name:
                reason = classify_failure_reason_local(line_clean)
                failed_lines.append((line, reason))
                continue

            cards = [n for n in all_nums if len(n) >= 4]
            old_card = cards[0] if cards else "غير متوفر"
            new_card = cards[-1] if len(cards) >= 2 else old_card
            selected_card = new_card if card_choice == "رقم البطاقة الحديث" else old_card

            idx_name = line_clean.find(name.split()[0])
            before = line_clean[:idx_name] if idx_name >= 0 else ""
            after = line_clean[idx_name + len(name):] if idx_name >= 0 else line_clean

            smalls_before = [int(n) for n in re.findall(r'\b\d{1,3}\b', before)]
            smalls_after = [int(n) for n in re.findall(r'\b\d{1,3}\b', after)]

            total = eligible = withheld = 0
            nums = smalls_after if len(smalls_after) >= 1 else smalls_before
            if len(nums) >= 3:
                total, eligible, withheld = nums[0], nums[1], nums[2]
            elif len(nums) == 2:
                total, eligible = nums[0], nums[1]
                withheld = 0
            elif len(nums) == 1:
                total = nums[0]
                eligible = withheld = 0
            else:
                all_smalls = [int(n) for n in re.findall(r'\b\d{1,3}\b', line_clean)]
                if len(all_smalls) >= 3:
                    total, eligible, withheld = all_smalls[-3], all_smalls[-2], all_smalls[-1]
                elif len(all_smalls) == 2:
                    total, eligible = max(all_smalls), min(all_smalls)
                    withheld = 0
                else:
                    reason = classify_failure_reason_local(line_clean)
                    failed_lines.append((line, reason))
                    continue

            if re.search(r'^\s*(اجمالي|المجموع|مجموع)\b', name):
                continue

            parsed_records.append({
                "اسم رب الأسرة": name,
                "رقم البطاقة": selected_card,
                "الكلي": total,
                "محجوب": withheld,
                "مستحق": eligible
            })

            if debug_limit and len(parsed_records) >= debug_limit:
                break

        except Exception as e:
            failed_lines.append((line, f"exception:{str(e)}"))
            if stop_on_error:
                raise

    raw_count = len(raw_lines)
    return parsed_records, failed_lines, raw_count

# -------------------------------------------------------------------------
# وحدة الاستخراج من Excel/CSV (ثابتة مع تحسين طفيف)
# -------------------------------------------------------------------------
def _parse_excel_to_list(file_obj, file_ext):
    records = []
    try:
        df_in = pd.read_excel(file_obj) if file_ext == 'xlsx' else pd.read_csv(file_obj)
        df_in.columns = df_in.columns.astype(str).str.strip()
        
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

# -------------------------------------------------------------------------
# الماستر: الدمج والفرز والتنظيف (محدّث لاستقبال failed_lines)
# -------------------------------------------------------------------------
def extract_and_clean_data(file_obj, card_choice, file_obj2=None, file2_ext=None, stop_on_error=False):
    # 1. استخراج الأساسي (الآن يعيد parsed و failed_lines و raw_count)
    parsed_primary, failed_primary, raw_count_primary = _parse_docx_to_list(file_obj, card_choice, stop_on_error=stop_on_error)
    
    # 2. استخراج ودمج الإضافي (إن وجد)
    parsed_secondary = []
    failed_secondary = []
    raw_count_secondary = 0
    if file_obj2:
        if file2_ext == 'docx':
            parsed_secondary, failed_secondary, raw_count_secondary = _parse_docx_to_list(file_obj2, card_choice, stop_on_error=stop_on_error)
        elif file2_ext in ['xlsx', 'csv']:
            parsed_secondary = _parse_excel_to_list(file_obj2, file2_ext)
            failed_secondary = []
            raw_count_secondary = len(parsed_secondary)
        else:
            pass

    all_parsed = parsed_primary + parsed_secondary
    all_failed = failed_primary + failed_secondary

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
        "raw_lines_primary": raw_count_primary,
        "raw_lines_secondary": raw_count_secondary,
        "parsed_count": len(all_parsed),
        "failed_count": len(all_failed),
        "duplicates_count": len(df_duplicates),
    }
    return df, df_duplicates, all_failed, stats

# -------------------------------------------------------------------------
# محرك بناء تقرير Word (ثابت)
# -------------------------------------------------------------------------
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
    
    if not df.empty:
        max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    else:
        max_name_len = 15
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

    total_all = df["الكلي"].astype(int).sum() if not df.empty else 0
    total_eligible = df["مستحق"].astype(int).sum() if not df.empty else 0
    total_withheld = df["محجوب"].astype(int).sum() if not df.empty else 0
    
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

# -------------------------------------------------------------------------
# واجهة الاستخدام (معدّلة لإظهار لوج الفشل وإحصاءات)
# -------------------------------------------------------------------------
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
                file2_ext = uploaded_file2.name.split('.')[-1].lower() if uploaded_file2 else None
                
                df_res, df_dup, failed_lines, stats = extract_and_clean_data(uploaded_file, selected_card, uploaded_file2, file2_ext)
                
                st.session_state.df_final = df_res
                st.session_state.df_duplicates = df_dup
                st.session_state.failed_lines = failed_lines
                st.session_state.stats = stats
                st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                st.session_state.selected_card = selected_card
                st.session_state.processing_done = True

                if df_res.empty:
                    st.error("لم يتم العثور على بيانات جداول متوافقة. يرجى التأكد من محتوى الملف.")
                else:
                    st.success(f"استخراج أولي: تم استخراج {len(df_res)} سجلّاً؛ فشل في استخراج {len(failed_lines)} سطر (سجّل للتدقيق).")
            except Exception as e:
                st.error(f"خطأ غير متوقع أثناء المعالجة: {e}")
                st.error(traceback.format_exc())
    else:
        st.warning("الرجاء رفع ملف الكشف الأساسي (Word) أولاً.")

# -------------------------------------------------------------------------
# بعد المعالجة: تنزيل التقارير + لوج الفشل
# -------------------------------------------------------------------------
if st.session_state.processing_done:
    df_final = st.session_state.df_final
    df_duplicates = st.session_state.df_duplicates
    failed_lines = st.session_state.failed_lines
    stats = st.session_state.stats
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

    # إظهار إحصاءات موجزة
    st.markdown("### إحصاءات الاستخراج")
    st.write({
        "عدد الأسطر الخام (أساسي)": stats.get("raw_lines_primary", 0),
        "عدد الأسطر الخام (ثانوي)": stats.get("raw_lines_secondary", 0),
        "عدد السجلات المستخرجة": stats.get("parsed_count", 0),
        "عدد الأسطر الفاشلة": stats.get("failed_count", 0),
        "عدد التكرارات المعزولة": stats.get("duplicates_count", 0),
    })

    # عرض عيّنة من failed_lines مع سبب الفشل
    if failed_lines:
        st.markdown("### أمثلة على الأسطر التي فشل استخراجها (مع سبب الفشل)")
        sample = failed_lines[:200]
        df_failed = pd.DataFrame(sample, columns=["السطر الخام", "سبب الفشل"])
        st.dataframe(df_failed)

        # زر تنزيل لوج الفشل كـ CSV
        csv_buffer = StringIO()
        writer = csv.writer(csv_buffer)
        writer.writerow(["السطر الخام", "سبب الفشل"])
        for ln, reason in failed_lines:
            writer.writerow([ln, reason])
        csv_buffer.seek(0)
        st.download_button(
            label="📥 تحميل لوج الأسطر الفاشلة (CSV)",
            data=csv_buffer.getvalue(),
            file_name=f"failed_lines_{output_filename}.csv",
            mime="text/csv"
        )

    # عرض نصي: نصائح سريعة بناءً على أسباب الفشل الشائعة
    st.markdown("### توصيات سريعة لتحسين الاستخراج")
    st.markdown("""
    - إذا كانت معظم أسباب الفشل `html_table_unparsed` فالمستند يحتوي على جداول داخلية بصيغة HTML-like؛ أعد حفظ المستند كـ Word حقيقي (بدون HTML داخل الفقرات) أو زودني بنسخة من الصفحات الأولى لأقوم بتخصيص parser.
    - إذا كانت الأسباب `no_arabic_name` أو `not_enough_numbers` فراجع تنسيق الأعمدة في الملف الأصلي (قد تكون الأسماء موزعة على خلايا متعددة أو أرقام البطاقات قصيرة).
    - استخدم زر تنزيل لوج الفشل لمراجعة الأسطر يدوياً وإضافة قواعد استثنائية إذا لزم.
    """)

# نهاية الملف
