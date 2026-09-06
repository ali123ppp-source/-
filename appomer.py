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
    
    # قراءة الملف حسب صيغته وتحويله إلى صفوف موحدة
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
                    # إزالة الفواصل العشرية للأرقام الصحيحة القادمة من الإكسل
                    if isinstance(cell, float) and cell.is_integer():
                        cells.append(str(int(cell)))
                    else:
                        cells.append(str(cell).strip().replace('\n', ' '))
                rows_data.append(cells)
    
    # تطبيق نفس منطق التنظيف والاستخراج المعتاد على الأسطر
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
        
        # قص اللقب والإبقاء على الاسم الثلاثي فقط بصرامة
        full_name = cells[name_idx]
        name_parts = full_name.split()
        three_part_name = " ".join(name_parts[:3])
            
        raw_records.append({
            "اسم رب الأسرة": three_part_name,
            "رقم البطاقة": selected_card_num,
            "الكلي": total,
            "محجوب": withheld,
            "مستحق": eligible
        })
        
    df = pd.DataFrame(raw_records)
    if not df.empty:
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
    return df

# -----------------------------------------------------------------------------
# محرك قراءة إضافي مخصص لكشوفات "القطع الغذائية" ذات الأعمدة السبعة الثابتة
# (نموذج قراءة جديد يُضاف بجانب المحرك العام دون أي تعديل عليه)
# الأعمدة المعروفة لهذا الشكل:
# ت الأصلي | رقم البطاقة القديمة | رقم البطاقة الجديدة | اسم العائلة |
# الأفراد الكلية | الأفراد المستحقة | عدد المحجوبين
# -----------------------------------------------------------------------------
def extract_ration_list_data(file_obj, card_choice):
    """
    قارئ خاص يعتمد على عناوين الأعمدة نفسها (وليس التخمين بالموقع)،
    مناسب لهذا الشكل تحديداً من الكشوفات ذات السبعة أعمدة الثابتة.
    يُعيد نفس هيكلة البيانات (DataFrame) التي يعيدها المحرك العام،
    لذا يعمل تلقائياً مع كل نماذج تنسيق الوورد الأربعة الحالية دون أي تعديل عليها.
    """
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
                # في حال وجود أكثر من جدول بنفس ترويسة الأعمدة، نتجاهل تكرار الترويسة
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
    idx_total = find_col(["الأفراد الكلية"])
    idx_eligible = find_col(["الأفراد المستحقة"])
    idx_withheld = find_col(["المحجوبين"])

    # إن لم تتطابق ترويسة الملف مع هذا الشكل المتوقع، لا نكمل (لتفادي نتائج خاطئة)
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
        if any("الإجمالي" in str(c) or "المجموع" in str(c) for c in cells):
            continue

        full_name = str(cells[idx_name]).strip()
        if not full_name:
            continue

        old_val = str(cells[idx_old]).strip() if idx_old != -1 and idx_old < len(cells) else ""
        new_val = str(cells[idx_new]).strip() if idx_new != -1 and idx_new < len(cells) else ""

        if card_choice == "رقم البطاقة الحديث":
            selected_card_num = new_val if new_val.isdigit() else (old_val if old_val.isdigit() else new_val)
        else:
            selected_card_num = old_val if old_val.isdigit() else (new_val if new_val.isdigit() else old_val)

        try:
            total_txt = cells[idx_total].strip() if idx_total < len(cells) else ""
            eligible_txt = cells[idx_eligible].strip() if idx_eligible < len(cells) else ""
            withheld_txt = cells[idx_withheld].strip() if idx_withheld < len(cells) else ""
            total = int(float(total_txt)) if total_txt else 0
            eligible = int(float(eligible_txt)) if eligible_txt else 0
            withheld = int(float(withheld_txt)) if withheld_txt else 0
        except ValueError:
            continue

        # الإبقاء على نفس منطق النظام العام: الاسم الثلاثي فقط
        three_part_name = " ".join(full_name.split()[:3])

        raw_records.append({
            "اسم رب الأسرة": three_part_name,
            "رقم البطاقة": selected_card_num if selected_card_num else "-",
            "الكلي": total,
            "محجوب": withheld,
            "مستحق": eligible
        })

    df = pd.DataFrame(raw_records)
    if not df.empty:
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
    return df

# -----------------------------------------------------------------------------
# محرك "قاعدة الأسماء الصحيحة" - لتصحيح الأسماء التالفة/المقلوبة اعتماداً على
# مطابقة رقم البطاقة التموينية فقط (وليس الاسم، لأن الاسم هو الحقل المشكوك فيه)
# -----------------------------------------------------------------------------
def normalize_card_number(card):
    """
    يوحّد صيغة رقم البطاقة للمقارنة فقط (يزيل الأصفار البادئة)، لأن بعض الملفات
    تخزّن الرقم نفسه بأصفار بادئة (مثال: 0000022) وأخرى بدونها (22) رغم أنه نفس الرقم.
    """
    s = str(card).strip()
    if s.isdigit():
        return str(int(s))
    return s

def build_name_reference_map(file_obj):
    """
    يقرأ ملف قاعدة الأسماء الصحيحة (xlsx أو docx) ويبني خريطة:
    رقم البطاقة -> الاسم الصحيح الكامل.
    يلتقط كل رقم بطاقة (5 أرقام فأكثر) موجود بالصف - سواء كان قديماً أو جديداً -
    وينسبه لنفس الاسم، لضمان المطابقة أياً كان نوع الرقم المستخدم في الملف الرئيسي.
    """
    file_ext = file_obj.name.split('.')[-1].lower()
    rows_data = []

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

    reference_map = {}
    for cells in rows_data:
        if not any(cells):
            continue
        joined = "".join(cells)
        if "المركز" in joined or "الوكيل" in joined or "اسم رب" in joined or "اسم العائلة" in joined:
            continue

        name_idx, max_len = -1, 0
        for i, c in enumerate(cells):
            if any('\u0600' <= ch <= '\u06FF' for ch in c) and not any(ch.isdigit() for ch in c):
                if len(c) > max_len:
                    max_len, name_idx = len(c), i
        if name_idx == -1:
            continue

        correct_name = cells[name_idx].strip()
        if not correct_name:
            continue

        for c in cells:
            if c.isdigit() and len(c) >= 5:
                reference_map[normalize_card_number(c)] = correct_name

    return reference_map

def normalize_name_for_compare(name):
    """توحيد بسيط لأشكال الحروف والمسافات فقط لغرض المقارنة (لا يُستخدم كقيمة نهائية)."""
    n = re.sub(r'\s+', ' ', str(name).strip())
    n = n.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ى', 'ي')
    return n

def apply_name_corrections(df, reference_map):
    """
    يفحص كل سجل: إن كان رقم بطاقته موجوداً في قاعدة الأسماء الصحيحة والاسم مختلف عنها،
    يستبدل حقل الاسم فقط بالاسم الصحيح من القاعدة (بقية الأعمدة تبقى كما هي دون أي مساس)،
    ويعيد الترتيب الأبجدي وإعادة ترقيم "ت"، مع تسجيل كل استبدال في تقرير منفصل.
    """
    if df.empty or not reference_map:
        return df, pd.DataFrame(columns=["رقم البطاقة", "الاسم قبل التصحيح", "الاسم بعد التصحيح"])

    df = df.copy()
    corrections = []

    for idx, row in df.iterrows():
        card = normalize_card_number(row["رقم البطاقة"])
        if card in reference_map:
            correct_three_part = " ".join(reference_map[card].split()[:3])
            current_name = str(row["اسم رب الأسرة"]).strip()

            if normalize_name_for_compare(current_name) != normalize_name_for_compare(correct_three_part):
                corrections.append({
                    "رقم البطاقة": card,
                    "الاسم قبل التصحيح": current_name,
                    "الاسم بعد التصحيح": correct_three_part
                })
                df.at[idx, "اسم رب الأسرة"] = correct_three_part

    df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
    df["ت"] = df.index + 1

    return df, pd.DataFrame(corrections)

def light_auto_clean_name(name):
    """
    تنظيف سطحي آمن فقط: إزالة أي رموز/أرقام/حروف غير عربية عالقة بالاسم وضغط المسافات.
    لا يحاول إطلاقاً تخمين حروف ناقصة أو إعادة ترتيب حروف مبعثرة - فقط إزالة الشوائب الواضحة.
    """
    if not name:
        return ""
    n = re.sub(r'[^\u0600-\u06FF\s]', ' ', str(name))
    return re.sub(r'\s+', ' ', n).strip()

def build_known_name_tokens(reference_map):
    """يبني مجموعة كل مقاطع الأسماء (الكلمات) الظاهرة في قاعدة الأسماء الصحيحة، لاستخدامها كقاموس مرجعي."""
    tokens = set()
    for name in reference_map.values():
        tokens.update(name.split())
    return tokens

def detect_suspicious_name(name, known_tokens=None):
    """
    فحص إرشادي لاكتشاف الأسماء المشتبه بتلفها/عدم وضوحها لسجلات لا يوجد لها
    أي مرجع في قاعدة الأسماء الصحيحة (غالباً سجلات جديدة لم تُضف للقاعدة بعد).
    إن تم تمرير known_tokens (قاموس مقاطع الأسماء المعروفة من القاعدة)، يُضاف فحص إضافي
    يكتشف الأحرف المخربطة التي تشكّل كلمة تبدو سليمة هيكلياً لكنها غير موجودة في القاموس.
    يُرجع: (هل الاسم مشتبه به، قائمة أسباب الاشتباه)
    """
    reasons = []
    original = str(name).strip()

    if not original:
        return True, ["الاسم فارغ"]

    if any(not ('\u0600' <= ch <= '\u06FF' or ch.isspace()) for ch in original):
        reasons.append("يحتوي على رموز/أرقام/حروف غير عربية")

    words = original.split()
    if len(words) < 3:
        reasons.append(f"الاسم غير مكتمل ({len(words)} من 3 مقاطع متوقعة)")

    if any(len(w) <= 1 for w in words):
        reasons.append("يحتوي على مقطع من حرف واحد فقط (قد يكون ناقصاً)")

    if re.search(r'(.)\1{2,}', original):
        reasons.append("يحتوي على تكرار غير طبيعي لحرف واحد")

    if not light_auto_clean_name(original):
        reasons.append("الاسم بعد إزالة الشوائب أصبح فارغاً - على الأغلب غير مقروء بالكامل")

    if known_tokens:
        unknown_words = [w for w in words if len(w) > 1 and w not in known_tokens]
        if unknown_words:
            reasons.append(f"يحتوي على مقطع/مقاطع غير موجودة ضمن قائمة الأسماء المعروفة بالقاعدة: {', '.join(unknown_words)} (قد يكون محرَّفاً بتبديل حروف)")

    return (len(reasons) > 0), reasons

def flag_unresolved_suspicious_names(df, reference_map):
    """
    يفحص السجلات التي لم يُعثر لرقم بطاقتها على أي مقابل في قاعدة الأسماء الصحيحة،
    ويكتشف من بينها ما يبدو اسمه تالفاً أو غير مفهوم، مع اقتراح تنظيف سطحي (بدون تخمين)،
    ويعيدها كتقرير مراجعة يدوية منفصل - هذه السجلات لا يمكن تصحيحها تلقائياً لعدم وجود مرجع لها.
    """
    known_tokens = build_known_name_tokens(reference_map) if reference_map else None
    flagged = []
    for _, row in df.iterrows():
        card_display = str(row["رقم البطاقة"]).strip()
        card = normalize_card_number(card_display)
        if card in reference_map:
            continue  # هذه عولجت بالفعل عبر المطابقة المباشرة
        name = str(row["اسم رب الأسرة"]).strip()
        is_suspicious, reasons = detect_suspicious_name(name, known_tokens)
        if is_suspicious:
            flagged.append({
                "رقم البطاقة": card_display,
                "الاسم الحالي": name,
                "اقتراح تنظيف سطحي (بدون تخمين)": light_auto_clean_name(name) or "—",
                "سبب الاشتباه": "، ".join(reasons)
            })
    return pd.DataFrame(flagged)

def build_corrections_report_excel(corrections_df, unresolved_df=None):
    """يبني ملف Excel بتقريرين: الأسماء المصحَّحة تلقائياً من القاعدة، والأسماء المشتبه بها التي تحتاج مراجعة يدوية."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        if corrections_df is None or corrections_df.empty:
            pd.DataFrame({"ملاحظة": ["لم يتم العثور على أي استبدالات لهذه الدفعة"]}).to_excel(
                writer, index=False, sheet_name="تم تصحيحها تلقائياً")
        else:
            corrections_df.to_excel(writer, index=False, sheet_name="تم تصحيحها تلقائياً")

        if unresolved_df is not None and not unresolved_df.empty:
            unresolved_df.to_excel(writer, index=False, sheet_name="تحتاج مراجعة يدوية")
        else:
            pd.DataFrame({"ملاحظة": ["لا توجد أسماء مشتبه بها غير موجودة بالقاعدة"]}).to_excel(
                writer, index=False, sheet_name="تحتاج مراجعة يدوية")
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word - النموذج الأول (الأصلي)
# -----------------------------------------------------------------------------
def build_professional_word_report(df, filename_base, card_choice):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(0.3)
        section.right_margin = Cm(0.3)
        
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز"]:
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
    
    headers = ["ت", "اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(0.44), Cm(0.9), Cm(0.9), Cm(0.9), Cm(3.0), Cm(2.19)]
    
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    
    for i, title in enumerate(headers):
        table.rows[0].cells[i].width = col_widths[i]
        if i in [3, 4, 5]:
            set_cell_vertical_text(table.rows[0].cells[i])
            format_cell_advanced(table.rows[0].cells[i], title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        else:
            format_cell_advanced(table.rows[0].cells[i], title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i==1 else "center", color_rgb=COLOR_NAVY_BLUE)
            
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        table.rows[idx+1]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0
        set_cell_no_wrap(row_cells[1])
        
        for i in range(8):
            row_cells[i].width = col_widths[i]
            val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else "x" if i==2 and is_eligible_zero else "" if i==2 else row["الكلي"] if i==3 else row["مستحق"] if i==4 else row["محجوب"] if i==5 else row["رقم البطاقة"] if i==6 else "محجوب" if i==7 and is_eligible_zero else ""
            font_size = 14 if i==5 else 12 if i==7 and is_eligible_zero else 16
            text_color = RGBColor(203, 67, 53) if i==7 and is_eligible_zero else None
            format_cell_advanced(row_cells[i], val, size_pt=font_size, font_name="Calibri", color_rgb=text_color, align="left" if i==1 else "center")
            
            if is_eligible_zero: set_cell_background(row_cells[i], "EC7063")
            else:
                if i==0: set_cell_background(row_cells[i], "D4E6F1")
                elif i==3: set_cell_background(row_cells[i], "EBF5FB")
                elif i==4: set_cell_background(row_cells[i], "E8F8F5")
                elif i==5: set_cell_background(row_cells[i], "FADBD8")
    return save_doc_buffer(doc, df)

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word - النموذج الثاني
# -----------------------------------------------------------------------------
def build_professional_word_report_v2(df, filename_base, card_choice):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Cm(0.5), Cm(0.5), Cm(0.3), Cm(0.3)
        
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    
    headers = ["ت", "اسم رب الأسرة", "حقل فارغ 1", "حقل فارغ 2", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    table = doc.add_table(rows=1, cols=9)
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(0.80), Cm(0.80), Cm(0.9), Cm(0.9), Cm(0.9), Cm(3.0), Cm(1.80)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    
    for i, title in enumerate(headers):
        table.rows[0].cells[i].width = col_widths[i]
        if i in [4, 5, 6]: set_cell_vertical_text(table.rows[0].cells[i])
        format_cell_advanced(table.rows[0].cells[i], title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i==1 else "center", color_rgb=COLOR_NAVY_BLUE)
            
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        table.rows[idx+1]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0
        set_cell_no_wrap(row_cells[1])
        
        for i in range(9):
            row_cells[i].width = col_widths[i]
            val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else "x" if i in [2,3] and is_eligible_zero else "" if i in [2,3] else row["الكلي"] if i==4 else row["مستحق"] if i==5 else row["محجوب"] if i==6 else row["رقم البطاقة"] if i==7 else "محجوب" if i==8 and is_eligible_zero else ""
            text_color = RGBColor(203, 67, 53) if i==8 and is_eligible_zero else None
            format_cell_advanced(row_cells[i], val, size_pt=14, font_name="Calibri", color_rgb=text_color, align="left" if i==1 else "center")
            
            if is_eligible_zero: set_cell_background(row_cells[i], "EC7063")
            else:
                if i==0: set_cell_background(row_cells[i], "D4E6F1")
                elif i==4: set_cell_background(row_cells[i], "EBF5FB")
                elif i==5: set_cell_background(row_cells[i], "E8F8F5")
            if i==6: set_cell_background(row_cells[i], "E5E7E9")
    return save_doc_buffer(doc, df)

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word - النموذج الثالث (عناوين 12، تفاصيل 16، 4 أشهر)
# -----------------------------------------------------------------------------
def build_professional_word_report_v3(df, filename_base, card_choice):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin, section.left_margin, section.right_margin = Cm(0.5), Cm(0.5), Cm(0.3), Cm(0.3)
        
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز"]:
        clean_name = clean_name.replace(w, "")
    clean_name = re.sub(r'[a-zA-Z]', '', clean_name)
    clean_name = re.sub(r'[\-_+_.]', '', clean_name)
    clean_name = " ".join(clean_name.split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    
    headers = ["ت", "اسم رب الأسرة", card_choice, "الكلي", "مستحق", "محجوب", "الشهر الأول", "الشهر الثاني", "الشهر الثالث", "الشهر الرابع"]
    table = doc.add_table(rows=1, cols=10)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(3.0), Cm(0.9), Cm(0.9), Cm(0.9), Cm(2.3), Cm(2.3), Cm(2.3), Cm(2.3)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    
    for i, title in enumerate(headers):
        table.rows[0].cells[i].width = col_widths[i]
        if i in [3, 4, 5]: set_cell_vertical_text(table.rows[0].cells[i])
        format_cell_advanced(table.rows[0].cells[i], title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 1 else "center", color_rgb=COLOR_NAVY_BLUE)
            
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        table.rows[idx+1]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0
        set_cell_no_wrap(row_cells[1])
        
        for i in range(10):
            row_cells[i].width = col_widths[i]
            val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else row["رقم البطاقة"] if i==2 else row["الكلي"] if i==3 else row["مستحق"] if i==4 else row["محجوب"] if i==5 else ""
            format_cell_advanced(row_cells[i], val, size_pt=16, font_name="Calibri", color_rgb=None, align="left" if i==1 else "center")
            
            if is_eligible_zero: set_cell_background(row_cells[i], "EC7063")
            else:
                if i == 0: set_cell_background(row_cells[i], "D4E6F1")
            
            if i == 3: set_cell_background(row_cells[i], "E5E7E9")
            if i == 5: set_cell_background(row_cells[i], "FCF3CF")
    return save_doc_buffer(doc, df)

# -----------------------------------------------------------------------------
# محرك بناء تقرير Word - النموذج الرابع (12 سلة، التعديل: العدد الكلي للأفراد)
# -----------------------------------------------------------------------------
def build_professional_word_report_v4(df, filename_base, card_choice):
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
    
    # تم تغيير عنوان العمود الرابع إلى "العدد الكلي" بدلاً من المستحق
    headers = ["ت", card_choice, "اسم المواطن", "العدد الكلي"] + [f"سلة {i}" for i in range(1, 13)]
    
    table = doc.add_table(rows=1, cols=16)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    set_table_borders(table, color_hex="2A4B7C")
    
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    
    # توزيع المقاسات للنموذج الرابع (12 سلة تملأ الورقة بأحجام متساوية)
    col_widths = [
        Cm(0.9),              # ت
        Cm(2.5),              # رقم البطاقة
        dynamic_name_width,   # اسم المواطن
        Cm(0.9)               # العدد الكلي
    ] + [Cm(1.05)] * 12       # 12 سلة بأحجام متساوية 
    
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        
        # جعل النصوص عمودية للأعمدة الضيقة من العدد الكلي إلى آخر سلة لتوفير مساحة
        if i >= 3:
            set_cell_vertical_text(hdr_cells[i])
        
        cell_align = "left" if i == 2 else "center"
        format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align=cell_align, color_rgb=COLOR_NAVY_BLUE)
            
    HEX_ELEGANT_BLUE = "D4E6F1"
    HEX_LIGHT_GREEN = "E8F8F5"
    HEX_ALERT_RED = "EC7063"
    
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        r_trPr = table.rows[idx+1]._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        is_eligible_zero = int(row["مستحق"]) == 0
        
        set_cell_no_wrap(row_cells[2])
        
        for i in range(16):
            row_cells[i].width = col_widths[i]
            
            val = ""
            cell_align = "center"
            font_size = 14  # حجم خط 14 ليتناسب مع 16 عمود
            
            if i == 0: val = row["ت"]
            elif i == 1: val = row["رقم البطاقة"]
            elif i == 2: 
                val = row["اسم رب الأسرة"]
                cell_align = "left" 
            elif i == 3: val = row["الكلي"] # تم التعديل لعرض الكلي بدلاً من مستحق
            elif i >= 4: val = "" # حقول السلات الـ 12 تبقى فارغة
                    
            format_cell_advanced(row_cells[i], val, size_pt=font_size, font_name="Calibri", color_rgb=None, align=cell_align)
            
            # التنسيق اللوني
            if is_eligible_zero:
                set_cell_background(row_cells[i], HEX_ALERT_RED)
            else:
                if i == 0: set_cell_background(row_cells[i], HEX_ELEGANT_BLUE)
                if i == 3: set_cell_background(row_cells[i], HEX_LIGHT_GREEN)

    return save_doc_buffer(doc, df)

# دالة لحفظ وعرض الإحصائيات المشتركة تحت الجدول
def save_doc_buffer(doc, df):
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
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
    stats_text_run = stats_p.add_run(stats_text)
    stats_text_run.font.name = "Segoe UI Semibold"
    stats_text_run.font.size = Pt(13)
    stats_text_run.bold = True
    stats_text_run.font.color.rgb = COLOR_NAVY_BLUE

    # الترقيم السفلي للصفحات
    footer = doc.sections[0].footer
    if len(footer.paragraphs) == 0:
        footer_p = footer.add_paragraph()
    else:
        footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.clear()
    f_run = footer_p.add_run("صفحة ")
    f_run.font.size = Pt(10)
    f_run._r.extend([
        parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'),
        parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'),
        parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'),
        parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    ])

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة استخدام التطبيق (Streamlit Interface)
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف المراد تدقيقه وتنسيقه للمطبعة</h3>", unsafe_allow_html=True)
# تعديل شريط الرفع ليقبل ملفات إكسل بصيغة xlsx إلى جانب الـ docx
uploaded_file = st.file_uploader("ارفع كشف الوكلاء", type=['docx', 'xlsx'], key="doc_input_v6", label_visibility="collapsed")

st.markdown("<h4 style='text-align: right;'>📚 قاعدة الأسماء الصحيحة (اختياري) - لتصحيح الأسماء التالفة أو المقلوبة</h4>", unsafe_allow_html=True)
reference_file = st.file_uploader(
    "ارفع ملف قاعدة الأسماء الصحيحة",
    type=['xlsx', 'docx'],
    key="ref_db_input_v1",
    label_visibility="collapsed",
    help="ملف يحتوي على أرقام البطاقات مع الأسماء الصحيحة الكاملة. سيتم استخدامه لمطابقة رقم البطاقة تلقائياً واستبدال أي اسم تالف أو مقلوب بالاسم الصحيح."
)

st.markdown("<br>", unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 1.4, 1.4])

with col1:
    selected_card = st.radio(
        "📄 اختر نوع رقم البطاقة:",
        ["رقم البطاقة القديم", "رقم البطاقة الحديث"],
        index=0,
        horizontal=False
    )

with col2:
    template_choice = st.radio(
        "🎨 اختر نموذج قالب الـ Word المطلوب:",
        [
            "النموذج الأول (الأصلي المطور)", 
            "النموذج الثاني (حجم 14 وحقلين فارغين)",
            "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)",
            "النموذج الرابع (12 سلة، العدد الكلي)"
        ],
        index=0,
        horizontal=False
    )

with col3:
    input_format_choice = st.radio(
        "🗂️ اختر طريقة قراءة الملف المرفوع:",
        [
            "تلقائي (يفحص الأعمدة ويختار الأنسب)",
            "فرض القراءة العامة",
            "فرض كشف القطع الغذائية (7 أعمدة ثابتة)"
        ],
        index=0,
        horizontal=False,
        help="الوضع التلقائي يجرّب أولاً قراءة الأعمدة السبعة الثابتة (ت الأصلي، رقم البطاقة القديمة، رقم البطاقة الجديدة، اسم العائلة، الأفراد الكلية، الأفراد المستحقة، عدد المحجوبين)، فإن لم تتطابق ترويسة الملف معها ينتقل تلقائياً للقراءة العامة. استخدم الخيارين اليدويين فقط إذا أردت إجبار طريقة معينة."
    )

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    current_ref_name = reference_file.name if reference_file else ""
    if (st.session_state.output_filename != current_filename or 
        st.session_state.selected_card != selected_card or 
        st.session_state.template_choice != template_choice or
        st.session_state.input_format_choice != input_format_choice or
        st.session_state.reference_filename != current_ref_name):
        st.session_state.processing_done = False

if st.button("⚙️ تشغيل محرك التنظيم والتنسيق المتقدم الكلي"):
    if uploaded_file:
        with st.spinner('جاري ترتيب القيود أبجدياً وإعداد التنسيق الشرطي والمقاييس...'):
            try:
                if input_format_choice == "فرض كشف القطع الغذائية (7 أعمدة ثابتة)":
                    df_res = extract_ration_list_data(uploaded_file, selected_card)
                    if df_res.empty:
                        st.error("لم يتم التعرف على أعمدة كشف القطع الغذائية المتوقعة في هذا الملف. تأكد من صحة الترويسة أو جرّب 'تلقائي' / 'فرض القراءة العامة'.")
                elif input_format_choice == "فرض القراءة العامة":
                    df_res = extract_and_clean_data(uploaded_file, selected_card)
                    if df_res.empty:
                        st.error("لم يتم العثور على بيانات جداول متوافقة.")
                else:
                    # الوضع التلقائي: نجرّب أولاً القارئ الخاص بالأعمدة السبعة الثابتة (يعتمد على تطابق أسماء
                    # الأعمدة نفسها، فهو آمن الفشل - إن لم تتطابق الترويسة يرجع جدولاً فارغاً دون أي تخمين خاطئ)
                    df_res = extract_ration_list_data(uploaded_file, selected_card)
                    if df_res.empty:
                        uploaded_file.seek(0)
                        df_res = extract_and_clean_data(uploaded_file, selected_card)
                        if df_res.empty:
                            st.error("لم يتم العثور على بيانات جداول متوافقة بأي من طريقتي القراءة.")

                if not df_res.empty:
                    corrections_df = pd.DataFrame(columns=["رقم البطاقة", "الاسم قبل التصحيح", "الاسم بعد التصحيح"])
                    unresolved_df = pd.DataFrame()
                    matched_count = 0
                    if reference_file:
                        with st.spinner('جاري مطابقة الأسماء مع قاعدة الأسماء الصحيحة...'):
                            reference_map = build_name_reference_map(reference_file)
                            matched_count = int(df_res["رقم البطاقة"].apply(
                                lambda c: normalize_card_number(c) in reference_map).sum())
                            df_res, corrections_df = apply_name_corrections(df_res, reference_map)
                            unresolved_df = flag_unresolved_suspicious_names(df_res, reference_map)

                    st.session_state.df_final = df_res
                    st.session_state.corrections_df = corrections_df
                    st.session_state.unresolved_df = unresolved_df
                    st.session_state.matched_count = matched_count
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.selected_card = selected_card
                    st.session_state.template_choice = template_choice
                    st.session_state.input_format_choice = input_format_choice
                    st.session_state.reference_filename = reference_file.name if reference_file else ""
                    st.session_state.processing_done = True
            except Exception as e:
                st.error(f"خطأ غير متوقع: {e}")
    else:
        st.warning("الرجاء رفع ملف docx أو xlsx أولاً.")

if st.session_state.processing_done:
    df_final = st.session_state.df_final
    output_filename = st.session_state.output_filename
    used_card_type = st.session_state.selected_card
    used_template = st.session_state.template_choice
    
    st.success(f"✅ تم التنظيم الأبجدي بنجاح لـ ({len(df_final)}) قيد اسم (تم قصرها على الاسم الثلاثي).")

    corrections_df = st.session_state.corrections_df
    unresolved_df = st.session_state.unresolved_df

    if reference_file:
        st.caption(f"🔎 تشخيص: تمت مطابقة رقم البطاقة لـ {st.session_state.matched_count} سجل من أصل {len(df_final)} مع قاعدة الأسماء المرفوعة.")

    if corrections_df is not None and not corrections_df.empty:
        st.markdown(
            f"<div class='report-box'>🛠️ تم تصحيح <b>{len(corrections_df)}</b> اسم تالف/مقلوب اعتماداً على مطابقة رقم البطاقة مع قاعدة الأسماء الصحيحة.</div>",
            unsafe_allow_html=True
        )
        st.dataframe(corrections_df, use_container_width=True)
    elif reference_file:
        st.info("ℹ️ لم يتم العثور على أي أسماء تحتاج تصحيحاً بمطابقة أرقام البطاقات مع القاعدة المرفوعة.")

    if unresolved_df is not None and not unresolved_df.empty:
        st.markdown(
            f"<div class='report-box' style='border-right-color:#CB4335;'>⚠️ يوجد <b>{len(unresolved_df)}</b> اسم مشتبه بتلفه لسجلات <b>غير موجودة في القاعدة</b> (على الأغلب مضافة حديثاً)، ولا يمكن تصحيحها تلقائياً لعدم وجود مرجع لها - يُرجى مراجعتها يدوياً ثم إضافتها للقاعدة مستقبلاً.</div>",
            unsafe_allow_html=True
        )
        st.dataframe(unresolved_df, use_container_width=True)

    if reference_file and ((corrections_df is not None and not corrections_df.empty) or (unresolved_df is not None and not unresolved_df.empty)):
        corrections_report = build_corrections_report_excel(corrections_df, unresolved_df)
        st.download_button(
            label="📥 تحميل تقرير تصحيح ومراجعة الأسماء (Excel - ورقتين)",
            data=corrections_report,
            file_name=f"تقرير_تصحيح_الاسماء_{output_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.spinner('جاري صياغة وهيكلة مستند Word المطور المختار...'):
        if used_template == "النموذج الأول (الأصلي المطور)":
            word_output = build_professional_word_report(df_final, output_filename, used_card_type)
        elif used_template == "النموذج الثاني (حجم 14 وحقلين فارغين)":
            word_output = build_professional_word_report_v2(df_final, output_filename, used_card_type)
        elif used_template == "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)":
            word_output = build_professional_word_report_v3(df_final, output_filename, used_card_type)
        else:
            word_output = build_professional_word_report_v4(df_final, output_filename, used_card_type)
        
    st.download_button(
        label="📥 تحميل كشف الوكلاء المنسق والجاهز للطباعة فوراً (Word)",
        data=word_output,
        file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
