import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import re
import os
import base64
from functools import lru_cache
from datetime import datetime

# تم تعديل الاستثناء هنا ليتجاهل خطأ OSError تماماً بدلاً من توقف التطبيق
try:
    from weasyprint import HTML
    WEASYPRINT_AVAILABLE = True
except Exception:  
    WEASYPRINT_AVAILABLE = False

try:
    import pdfkit
    PDFKIT_AVAILABLE = True
except Exception:
    PDFKIT_AVAILABLE = False

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

# -----------------------------------------------------------------------------
# سجل الملفات المعالجة (لعرضها لاحقاً كجدول ديناميكي، الأحدث أولاً)
# -----------------------------------------------------------------------------
PROCESSING_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "processing_log.csv")
LOG_COLUMNS = ["التاريخ والوقت", "اسم الملف", "عدد القيود", "نوع البطاقة", "طول الاسم", "القالب المستخدم", "ترتيب البيانات"]

def log_processed_file(filename, record_count, card_choice, name_length_choice, template_choice, sort_choice):
    entry = pd.DataFrame([{
        "التاريخ والوقت": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "اسم الملف": filename,
        "عدد القيود": record_count,
        "نوع البطاقة": card_choice,
        "طول الاسم": name_length_choice,
        "القالب المستخدم": template_choice,
        "ترتيب البيانات": sort_choice,
    }])
    header = not os.path.exists(PROCESSING_LOG_PATH)
    try:
        entry.to_csv(PROCESSING_LOG_PATH, mode='a', header=header, index=False, encoding='utf-8-sig')
    except Exception:
        pass  # تسجيل السجل ثانوي؛ لا يجب أن يوقف معالجة الملف الأساسية

def load_processing_log():
    if not os.path.exists(PROCESSING_LOG_PATH):
        return pd.DataFrame(columns=LOG_COLUMNS)
    try:
        log_df = pd.read_csv(PROCESSING_LOG_PATH, encoding='utf-8-sig')
    except Exception:
        return pd.DataFrame(columns=LOG_COLUMNS)
    log_df["_ts"] = pd.to_datetime(log_df["التاريخ والوقت"], errors='coerce')
    log_df = log_df.sort_values(by="_ts", ascending=False).drop(columns="_ts").reset_index(drop=True)
    return log_df

with st.expander("🕘 سجل الملفات المعالجة سابقاً (الأحدث أولاً)"):
    log_df = load_processing_log()
    if log_df.empty:
        st.info("لا يوجد أي ملفات تمت معالجتها بعد.")
    else:
        st.dataframe(log_df, use_container_width=True, hide_index=True)

if "processing_done" not in st.session_state:
    st.session_state.processing_done = False
    st.session_state.results = []
    st.session_state.uploaded_filenames = []
    st.session_state.selected_card = ""
    st.session_state.template_choice = ""
    st.session_state.name_choice = ""
    st.session_state.sort_choice = ""
    st.session_state.merge_choice = ""
    st.session_state.extraction_warnings = {}

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

def _get_or_create_char_style(doc, font_name, size_pt, bold, color_rgb):
    """يبني نمط حرف (character style) مرة واحدة لكل توليفة (خط/حجم/غامق/لون) ويُعاد استخدامه
    بالإشارة إليه (w:rStyle) بدل تكرار خصائص التنسيق الكاملة (rFonts/sz/b/color) داخل كل
    خلية على حدة. هذا التكرار وحده كان يُضخّم حجم XML الداخلي للمستند بشدة في الملفات
    الكبيرة (مثلاً 2.6 ميجابايت لـ1089 صف فقط) ويُثقل أداء Word عند الفتح والتمرير والتعديل،
    رغم أن حجم الملف على القرص يبدو صغيراً بسبب الضغط."""
    color_key = str(color_rgb) if color_rgb else "none"
    safe_font = re.sub(r'[^A-Za-z0-9]', '', font_name) or "Font"
    style_name = f"CellFmt_{safe_font}_{size_pt}_{int(bool(bold))}_{color_key}"
    try:
        return doc.styles[style_name]
    except KeyError:
        pass
    style = doc.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
    style.font.name = font_name
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    if color_rgb:
        style.font.color.rgb = color_rgb
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    return style

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

    style = _get_or_create_char_style(cell.part.document, font_name, size_pt, bold, color_rgb)
    for run in p.runs:
        run.style = style

def _add_styled_run(p, text, bold, color_rgb, font_name, size_pt):
    doc = p.part.document
    style = _get_or_create_char_style(doc, font_name, size_pt, bold, color_rgb)
    run = p.add_run(text)
    run.style = style
    return run

def format_name_cell_with_small_suffix(cell, text, base_size=16, small_size=10, font_name="Calibri", bold=False, color_rgb=None, align="left"):
    """اسم رباعي (4 مقاطع فأكثر): يُعرض المقطع الرابع بخط أصغر من باقي الاسم."""
    words = str(text).split()
    cell.text = ""
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align == "right" else WD_ALIGN_PARAGRAPH.LEFT if align == "left" else WD_ALIGN_PARAGRAPH.CENTER
    pPr = p.paragraph_format.element.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))

    if len(words) >= 4:
        _add_styled_run(p, " ".join(words[:-1]) + " ", bold, color_rgb, font_name, base_size)
        _add_styled_run(p, words[-1], bold, color_rgb, font_name, small_size)
    else:
        _add_styled_run(p, " ".join(words), bold, color_rgb, font_name, base_size)

def format_header_cell_two_lines(cell, full_text, bold=True, size_pt=14, font_name="Segoe UI Semibold", color_rgb=None):
    """يقسم عنوان الحقل على سطرين (آخر كلمة بسطر مستقل) للسماح بتقليل عرض العمود."""
    line1, _, line2 = str(full_text).rpartition(" ")
    if not line1:
        line1, line2 = line2, ""
    cell.text = ""
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p.paragraph_format.element.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))

    run1 = _add_styled_run(p, line1, bold, color_rgb, font_name, size_pt)
    if line2:
        run1.add_break()
        _add_styled_run(p, line2, bold, color_rgb, font_name, size_pt)

# -----------------------------------------------------------------------------
# بنرات الأحرف الأبجدية: فاصل ملوّن مميز فوق أول اسم بكل حرف عند الترتيب الأبجدي
# -----------------------------------------------------------------------------
ARABIC_LETTER_BANNER_COLORS = {
    "ا": "2E4053", "ب": "A93226", "ت": "117864", "ث": "B9770E", "ج": "512E5F",
    "ح": "1A5276", "خ": "6E2C00", "د": "196F3D", "ذ": "922B21", "ر": "154360",
    "ز": "7D6608", "س": "4A235A", "ش": "0E6251", "ص": "78281F", "ض": "1B4F72",
    "ط": "145A32", "ظ": "7B241C", "ع": "5B2C6F", "غ": "873600", "ف": "1F618D",
    "ق": "186A3B", "ك": "9A7D0A", "ل": "6C3483", "م": "1B2631", "ن": "A04000",
    "ه": "0B5345", "و": "7E5109", "ي": "283747",
}

def get_name_group_letter(name):
    """أول حرف من الاسم؛ تُجمّع كل أشكال الألف/الهمزة (أ إ آ ا) تحت حرف \"ا\" واحد."""
    text = str(name).strip()
    if not text:
        return "ا"
    first = text[0]
    if first in ("أ", "إ", "آ"):
        return "ا"
    return first

def get_letter_base_color(letter):
    """اللون الأساسي المشبع للحرف (يُستخدم لنص البانر)."""
    return ARABIC_LETTER_BANNER_COLORS.get(letter, "5D6D7E")

def blend_color_with_white(hex_color, alpha=0.25):
    """يمزج اللون مع الأبيض بنسبة شفافية alpha لإعطاء تأثير هايلايت خفيف (بديل الشفافية الحقيقية)."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    r = round(r * alpha + 255 * (1 - alpha))
    g = round(g * alpha + 255 * (1 - alpha))
    b = round(b * alpha + 255 * (1 - alpha))
    return f"{r:02X}{g:02X}{b:02X}"

def get_letter_banner_color(letter):
    """لون هايلايت شفاف (25%) لكل حرف، يُستخدم لخلفية البانر وخانة الترقيم معاً."""
    return blend_color_with_white(get_letter_base_color(letter), 0.25)

def add_letter_banner_row(table, letter, height_inches=0.45):
    """يضيف صف بانر مدموج على كامل عرض الجدول بخلفية هايلايت شفافة (25%) ونص بلون الحرف الأساسي."""
    banner_row = table.add_row()
    banner_row.height = Inches(height_inches)
    cells = banner_row.cells
    merged = cells[0]
    for extra_cell in cells[1:]:
        merged = merged.merge(extra_cell)
    color_hex = get_letter_banner_color(letter)
    set_cell_background(merged, color_hex)
    merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    format_cell_advanced(merged, letter, bold=True, size_pt=18, font_name="Segoe UI Semibold", align="center", color_rgb=RGBColor.from_string(get_letter_base_color(letter)))
    return color_hex

def setup_document_layout(doc, filename_base, is_a3=False):
    """إعداد هوامش الصفحة، الترويسة، والتذييل حسب المطلوب"""
    for section in doc.sections:
        if is_a3:
            section.page_width, section.page_height = Cm(29.7), Cm(42.0)
        section.top_margin, section.bottom_margin = Cm(0.5), Cm(0.5)
        section.left_margin, section.right_margin = Cm(0.3), Cm(0.3)
        
        # ارتفاع الترويسة والتذييل
        section.header_distance = Inches(0.2)
        section.footer_distance = Inches(0.2)
        
        # فحص اسم الملف لاضافة الكلمة المناسبة في الترويسة
        fname_lower = str(filename_base).lower()
        header_text = ""
        if "flour" in fname_lower:
            header_text = "طحين"
        elif "food" in fname_lower:
            header_text = "غذائية"
            
        if header_text:
            header = section.header
            hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hp.paragraph_format.element.get_or_add_pPr().append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
            run = hp.add_run(header_text)
            run.font.name = "Calibri"
            run._r.get_or_add_rPr().append(parse_xml(f'<w:cs {nsdecls("w")} w:val="Calibri"/>'))
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.bold = True

# -----------------------------------------------------------------------------
# محرك استخراج البيانات اعتماداً على عناوين الأعمدة الفعلية (تقارير منسّقة)
# -----------------------------------------------------------------------------
_ARABIC_DIACRITICS_RE = re.compile(r'[ؐ-ًؚ-ٰٟۖ-ۜ۟-۪ۨ-ۭ]')

def _strip_diacritics_and_unify_hamza(text):
    """يزيل التشكيل ويوحّد أشكال الهمزة/الألف، مع الحفاظ على المسافات (بعكس
    _normalize_header_cell) — مفيد حين تهم حدود الكلمات، كتمييز أول كلمة بحقل الاسم."""
    text = _ARABIC_DIACRITICS_RE.sub('', str(text))
    return text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")

def _normalize_header_cell(cell):
    return _strip_diacritics_and_unify_hamza(cell).replace(" ", "")

def _locate_header_row(rows_data):
    empty_idx_map = {"ت": -1, "اسم": -1, "كلي": -1, "مستحق": -1, "محجوب": -1, "بطاقة_قديم": -1, "بطاقة_حديث": -1}
    # "اسم" و"مستحق" فقط أساسيان في كل الملفات المعروفة؛ "كلي"/"محجوب"/رقم
    # البطاقة أعمدة اختيارية غير موجودة في بعض القوالب (مثل ملف توزيع مواد
    # غذائية بلا رقم بطاقة ولا عمود "كلي" منفصل).
    required = ["اسم", "مستحق"]

    for i, row in enumerate(rows_data):
        # صف ترويسة حقيقي يتكوّن من خلايا عناوين منفصلة، لا فقرة نصية طويلة
        # (مثل ملاحظات منهجية قد تحتوي بالصدفة على كلمات مشابهة لعناوين الأعمدة).
        # عدد الأحرف وحده لا يكفي: ملاحظة قد تُوزَّع على عدة خلايا يكون طول كل
        # منها معقولاً بمفرده لكنها فقرة شرح لا عنوان عمود، فنفحص أيضاً عدد
        # الكلمات بكل خلية — عنوان عمود حقيقي (حتى الطويل منه) نادراً ما يتجاوز
        # 6 كلمات، بعكس جملة توضيحية كاملة.
        if len(row) < 4 or any(len(cell) > 60 or len(cell.split()) > 6 for cell in row):
            continue

        row_joined = _normalize_header_cell("".join(row))
        if not ("اسم" in row_joined and ("كلي" in row_joined or "مجموع" in row_joined or "مستحق" in row_joined or "بطاق" in row_joined or "تموين" in row_joined)):
            continue

        idx_map = dict(empty_idx_map)
        for j, cell in enumerate(row):
            c = _normalize_header_cell(cell)
            if not c:
                continue
            if c in ["ت", "تسلسل", "التسلسل", "م"]:
                idx_map["ت"] = j
            elif "اسم" in c:
                idx_map["اسم"] = j
            elif "مستحق" in c:
                idx_map["مستحق"] = j
            elif "محجوب" in c:
                idx_map["محجوب"] = j
            elif "كلي" in c or "اجمالي" in c or "مجموع" in c:
                idx_map["كلي"] = j
            elif ("تسلسل" not in c
                  and not any(w in c for w in ["هاتف", "موبايل", "جوال", "تلفون", "فون", "ملف", "قيد"])
                  and ("بطاق" in c or "تموين" in c or "رقم" in c)):
                if "حديث" in c or "جديد" in c:
                    idx_map["بطاقة_حديث"] = j
                elif "قديم" in c or "سابق" in c:
                    idx_map["بطاقة_قديم"] = j
                elif idx_map["بطاقة_قديم"] == -1:
                    idx_map["بطاقة_قديم"] = j

        if all(idx_map[k] != -1 for k in required):
            return i, idx_map

    return -1, empty_idx_map

def _extract_records_by_headers(rows_data, card_choice, name_length_choice):
    header_idx, idx_map = _locate_header_row(rows_data)
    required = ["اسم", "مستحق"]
    if header_idx == -1 or any(idx_map[k] == -1 for k in required):
        return None

    take = 4 if name_length_choice == "الاسم الرباعي (إن وجد)" else 3
    footer_label_re = re.compile(r'(ال)?(مجموع|اجمالي|وكيل)$')
    records = []
    for i in range(header_idx + 1, len(rows_data)):
        row = rows_data[i]
        row_joined = "".join(row)
        if not row_joined:
            continue

        def get_val(key):
            idx = idx_map[key]
            return row[idx] if 0 <= idx < len(row) else ""

        def get_num(key):
            digits = ''.join(filter(str.isdigit, get_val(key)))
            return int(digits) if digits else 0

        name_val = get_val("اسم")
        if not name_val or len(name_val) < 3:
            continue

        # صف تذييل/توقيع (مثل "المجموع" أو "الوكيل خالد ياسين") يُميَّز عن اسم عائلة
        # حقيقي بأن الكلمة الدالة تقع في *أول* حقل الاسم، بعكس اسم شخص حقيقي قد
        # ينتهي بلقب عائلة مثل "خالد ياسين الوكيل" — الفحص هنا على أول كلمة فقط،
        # لا الحقل كاملاً، حتى لا يُستبعد اسم حقيقي أو كلمة أطول تبدأ بنفس الحروف
        # ("مجموعة سكنية" مثلاً).
        first_word = _strip_diacritics_and_unify_hamza(name_val).split()[0]
        if footer_label_re.fullmatch(first_word):
            continue

        # عائلة حقيقية تضم فرداً واحداً على الأقل دائماً؛ "الكلي" = 0 يعني عملياً
        # أن هذا الصف ليس سجل عائلة (كصف توقيع/ملاحظة انزلق من جدول لاحق بالملف
        # وصادف أن قيمة الاسم فيه بدت كاسم شخص حقيقي). لا ينطبق هذا الفحص إن لم
        # يوجد عمود "كلي" أصلاً بالملف (مثل قوالب توزيع مواد غذائية بلا عمود كلي).
        if idx_map["كلي"] != -1 and get_num("كلي") == 0:
            continue

        final_name = " ".join(name_val.split()[:take])
        old_card = ''.join(filter(str.isdigit, get_val("بطاقة_قديم")))
        new_card = ''.join(filter(str.isdigit, get_val("بطاقة_حديث")))

        card_fields = {}
        if card_choice == "رقم البطاقة الحديث":
            card_fields["رقم البطاقة"] = new_card or old_card
        elif card_choice == "القديم والحديث":
            card_fields["رقم البطاقة القديم"] = old_card
            card_fields["رقم البطاقة الحديث"] = new_card
        else:
            card_fields["رقم البطاقة"] = old_card or new_card

        records.append({
            "اسم رب الأسرة": final_name,
            **card_fields,
            "الكلي": get_num("كلي"),
            "محجوب": get_num("محجوب"),
            "مستحق": get_num("مستحق"),
        })

    return records

def _validate_extracted_records(records):
    """فحص أمان: يبحث عن مؤشرات على خطأ بتحديد مكان الأعمدة (كأن ينزلق رقم البطاقة إلى عمود الكلي)
    ويُرجع رسائل تحذير واضحة للمستخدم بدل تمرير بيانات خاطئة بصمت."""
    warnings = []
    if not records:
        return warnings

    n = len(records)

    huge_total = sum(1 for r in records if r.get("الكلي", 0) >= 1000)
    if huge_total:
        warnings.append(
            f"يحتوي عمود \"الكلي\" على قيمة كبيرة جداً (1000 فأكثر) في {huge_total} من {n} سجل — "
            f"هذا غير منطقي لعدد أفراد الأسرة، ويُحتمل أن عمود \"الكلي\" أُخذ خطأً من عمود رقم البطاقة. "
            f"يرجى مراجعة ترويسة الملف الأصلي."
        )

    # عمود "الكلي" قد لا يوجد أصلاً في بعض القوالب (مثل توزيع مواد غذائية بلا
    # عمود كلي منفصل)، وفي هذه الحالة تكون قيمته 0 لكل السجلات بلا استثناء —
    # هذا لا يعني خطأ بتحديد الأعمدة، فلا داعي لفحص "مستحق > الكلي" عليه.
    has_kuli_field = any(r.get("الكلي", 0) > 0 for r in records)
    exceeds = sum(1 for r in records if r.get("مستحق", 0) > r.get("الكلي", 0)) if has_kuli_field else 0
    if exceeds:
        warnings.append(
            f"في {exceeds} من {n} سجل، \"مستحق\" أكبر من \"الكلي\" — وهذا غير منطقي (لا يمكن أن يتجاوز "
            f"عدد المستحقين عدد أفراد الأسرة الكلي). يُحتمل وجود خطأ في تحديد مكان الأعمدة."
        )

    card_keys = [k for k in records[0].keys() if k.startswith("رقم البطاقة")]
    for key in card_keys:
        missing_or_short = sum(1 for r in records if len(str(r.get(key, ""))) < 4)
        # إن كان العمود فاضياً بالكامل (100%) فهذا يعني أن القالب المصدر لا
        # يتضمن رقم بطاقة أصلاً (وليس خطأ قراءة) — لا تحذير في هذه الحالة.
        if n * 0.5 < missing_or_short < n:
            warnings.append(
                f"عمود \"{key}\" فاضي أو قصير جداً (أقل من 4 خانات) في أكثر من نصف السجلات "
                f"({missing_or_short} من {n}) — يُحتمل أن قراءة رقم البطاقة من الملف الأصلي غير صحيحة."
            )

    return warnings

# -----------------------------------------------------------------------------
# محرك قراءة وتنظيف البيانات المطور
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj, card_choice, name_length_choice, sort_alphabetically=True):
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
                        cells.append("")
                    elif isinstance(cell, float) and cell.is_integer():
                        cells.append(str(int(cell)))
                    else:
                        cells.append(str(cell).strip().replace('\n', ' '))
                rows_data.append(cells)

    header_records = _extract_records_by_headers(rows_data, card_choice, name_length_choice)
    if header_records:
        warnings = _validate_extracted_records(header_records)
        df = pd.DataFrame(header_records)
        if not df.empty:
            if sort_alphabetically:
                df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
            df.insert(0, "ت", df.index + 1)
        return df, warnings

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

        card_fields = {}
        if card_choice == "رقم البطاقة الحديث":
            card_fields["رقم البطاقة"] = new_card_num
        elif card_choice == "القديم والحديث":
            card_fields["رقم البطاقة القديم"] = old_card_num
            card_fields["رقم البطاقة الحديث"] = new_card_num
        else:
            card_fields["رقم البطاقة"] = old_card_num

        digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
        if len(digit_cells) >= 3:
            withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
        elif len(digit_cells) == 2:
            withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
        else:
            continue
        
        full_name = cells[name_idx]
        name_parts = full_name.split()
        
        if name_length_choice == "الاسم الرباعي (إن وجد)":
            final_name = " ".join(name_parts[:4])
        else:
            final_name = " ".join(name_parts[:3])
            
        raw_records.append({
            "اسم رب الأسرة": final_name,
            **card_fields,
            "الكلي": total,
            "محجوب": withheld,
            "مستحق": eligible
        })
        
    warnings = _validate_extracted_records(raw_records)
    df = pd.DataFrame(raw_records)
    if not df.empty:
        if sort_alphabetically:
            df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        df.insert(0, "ت", df.index + 1)
    return df, warnings

# -----------------------------------------------------------------------------
# دوال إنشاء النماذج (1 إلى 7) بصيغة Word
# -----------------------------------------------------------------------------
def build_professional_word_report(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)

    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    
    is_combined = card_choice == "القديم والحديث"
    orig_headers = ["ت", "اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(0.44), Cm(0.9), Cm(0.9), Cm(0.9), Cm(1.8), Inches(1.0)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    remaining_indices = [1, 2, 3, 4, 5, 7] if is_combined else None

    headers = (["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]) if is_combined else orig_headers
    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)

    if is_combined:
        hdr_cells = table.rows[0].cells
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=14, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 1
        for label in ["القديم", "الحديث"]:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, label, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
        for i in remaining_indices:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            if i in [3, 4, 5]:
                set_cell_vertical_text(cell)
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i==1 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i in [3, 4, 5]:
                set_cell_vertical_text(cell)
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            elif i == 6:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=14, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i==1 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=16, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)
            j = 1
            for card_key in ["رقم البطاقة القديم", "رقم البطاقة الحديث"]:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
                format_cell_advanced(cell, row[card_key], size_pt=16, font_name="Calibri", align="center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                j += 1
            for i in remaining_indices:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
                if i == 1: set_cell_no_wrap(cell)
                val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else "x" if i==2 and is_eligible_zero else "" if i==2 else row["الكلي"] if i==3 else row["مستحق"] if i==4 else row["محجوب"] if i==5 else row["رقم البطاقة"] if i==6 else "محجوب" if i==7 and is_eligible_zero else ""
                font_size = 14 if i==5 else 12 if i==7 and is_eligible_zero else 16
                text_color = RGBColor(203, 67, 53) if i==7 and is_eligible_zero else None
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=16, small_size=10, font_name="Calibri", color_rgb=text_color, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=font_size, font_name="Calibri", color_rgb=text_color, align="left" if i==1 else "center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i==3: set_cell_background(cell, "EBF5FB")
                    elif i==4: set_cell_background(cell, "E8F8F5")
                    elif i==5: set_cell_background(cell, "FADBD8")
                j += 1
        else:
            set_cell_no_wrap(row_cells[1])
            for i in range(8):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else "x" if i==2 and is_eligible_zero else "" if i==2 else row["الكلي"] if i==3 else row["مستحق"] if i==4 else row["محجوب"] if i==5 else row["رقم البطاقة"] if i==6 else "محجوب" if i==7 and is_eligible_zero else ""
                font_size = 14 if i==5 else 12 if i==7 and is_eligible_zero else 16
                text_color = RGBColor(203, 67, 53) if i==7 and is_eligible_zero else None
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=16, small_size=10, font_name="Calibri", color_rgb=text_color, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=font_size, font_name="Calibri", color_rgb=text_color, align="left" if i==1 else "center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i==0: set_cell_background(cell, current_letter_color)
                    elif i==3: set_cell_background(cell, "EBF5FB")
                    elif i==4: set_cell_background(cell, "E8F8F5")
                    elif i==5: set_cell_background(cell, "FADBD8")
    return save_doc_buffer(doc, df)

def build_professional_word_report_v2(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)
    
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    is_combined = card_choice == "القديم والحديث"
    orig_headers = ["ت", "اسم رب الأسرة", "حقل فارغ 1", "حقل فارغ 2", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(0.80), Cm(0.80), Cm(0.9), Cm(0.9), Cm(0.9), Cm(1.8), Cm(1.80)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    remaining_indices = [1, 2, 3, 4, 5, 6, 8] if is_combined else None

    headers = (["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]) if is_combined else orig_headers
    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)

    if is_combined:
        hdr_cells = table.rows[0].cells
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=14, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 1
        for label in ["القديم", "الحديث"]:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, label, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
        for i in remaining_indices:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            if i in [4, 5, 6]: set_cell_vertical_text(cell)
            format_cell_advanced(cell, title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i==1 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i in [4, 5, 6]: set_cell_vertical_text(cell)
            if i == 7:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=14, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=14, font_name="Segoe UI Semibold", align="left" if i==1 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)
            j = 1
            for card_key in ["رقم البطاقة القديم", "رقم البطاقة الحديث"]:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
                format_cell_advanced(cell, row[card_key], size_pt=14, font_name="Calibri", align="center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                j += 1
            for i in remaining_indices:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
                if i == 1: set_cell_no_wrap(cell)
                val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else "x" if i in [2,3] and is_eligible_zero else "" if i in [2,3] else row["الكلي"] if i==4 else row["مستحق"] if i==5 else row["محجوب"] if i==6 else row["رقم البطاقة"] if i==7 else "محجوب" if i==8 and is_eligible_zero else ""
                text_color = RGBColor(203, 67, 53) if i==8 and is_eligible_zero else None
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=text_color, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=text_color, align="left" if i==1 else "center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i==4: set_cell_background(cell, "EBF5FB")
                    elif i==5: set_cell_background(cell, "E8F8F5")
                if i==6: set_cell_background(cell, "E5E7E9")
                j += 1
        else:
            set_cell_no_wrap(row_cells[1])
            for i in range(9):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else "x" if i in [2,3] and is_eligible_zero else "" if i in [2,3] else row["الكلي"] if i==4 else row["مستحق"] if i==5 else row["محجوب"] if i==6 else row["رقم البطاقة"] if i==7 else "محجوب" if i==8 and is_eligible_zero else ""
                text_color = RGBColor(203, 67, 53) if i==8 and is_eligible_zero else None
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=text_color, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=text_color, align="left" if i==1 else "center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i==0: set_cell_background(cell, current_letter_color)
                    elif i==4: set_cell_background(cell, "EBF5FB")
                    elif i==5: set_cell_background(cell, "E8F8F5")
                if i==6: set_cell_background(cell, "E5E7E9")
    return save_doc_buffer(doc, df)

def build_professional_word_report_v3(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)
    
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    is_combined = card_choice == "القديم والحديث"
    orig_headers = ["ت", "اسم رب الأسرة", card_choice, "الكلي", "مستحق", "محجوب", "الشهر الأول", "الشهر الثاني", "الشهر الثالث", "الشهر الرابع"]
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(1.8), Cm(0.9), Cm(0.9), Cm(0.9), Cm(2.3), Cm(2.3), Cm(2.3), Cm(2.3)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    remaining_indices = [1, 3, 4, 5, 6, 7, 8, 9] if is_combined else None

    headers = (["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]) if is_combined else orig_headers
    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)

    if is_combined:
        hdr_cells = table.rows[0].cells
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 1
        for label in ["القديم", "الحديث"]:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, label, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
        for i in remaining_indices:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            if i in [3, 4, 5]: set_cell_vertical_text(cell)
            format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 1 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i in [3, 4, 5]: set_cell_vertical_text(cell)
            if i == 2:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 1 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=16, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)
            j = 1
            for card_key in ["رقم البطاقة القديم", "رقم البطاقة الحديث"]:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
                format_cell_advanced(cell, row[card_key], size_pt=16, font_name="Calibri", align="center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                j += 1
            for i in remaining_indices:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
                if i == 1: set_cell_no_wrap(cell)
                val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else row["رقم البطاقة"] if i==2 else row["الكلي"] if i==3 else row["مستحق"] if i==4 else row["محجوب"] if i==5 else ""
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=16, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=16, font_name="Calibri", color_rgb=None, align="left" if i==1 else "center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                if i == 3: set_cell_background(cell, "E5E7E9")
                if i == 5: set_cell_background(cell, "FCF3CF")
                j += 1
        else:
            set_cell_no_wrap(row_cells[1])
            for i in range(10):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                val = row["ت"] if i==0 else row["اسم رب الأسرة"] if i==1 else row["رقم البطاقة"] if i==2 else row["الكلي"] if i==3 else row["مستحق"] if i==4 else row["محجوب"] if i==5 else ""
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=16, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=16, font_name="Calibri", color_rgb=None, align="left" if i==1 else "center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i == 0: set_cell_background(cell, current_letter_color)
                if i == 3: set_cell_background(cell, "E5E7E9")
                if i == 5: set_cell_background(cell, "FCF3CF")
    return save_doc_buffer(doc, df)

def build_professional_word_report_v4(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)
    
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    is_combined = card_choice == "القديم والحديث"
    orig_headers = ["ت", card_choice, "اسم المواطن", "العدد الكلي"] + [f"سلة {i}" for i in range(1, 13)]
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    col_widths = [Cm(0.9), Cm(1.8), dynamic_name_width, Cm(0.9)] + [Cm(1.05)] * 12
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    remaining_indices = list(range(2, 16)) if is_combined else None

    headers = (["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]) if is_combined else orig_headers
    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)

    hdr_cells = table.rows[0].cells
    if is_combined:
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 1
        for label in ["القديم", "الحديث"]:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, label, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
        for i in remaining_indices:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            if i >= 3: set_cell_vertical_text(cell)
            format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 2 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = hdr_cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i >= 3: set_cell_vertical_text(cell)
            if i == 1:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 2 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)
            j = 1
            for card_key in ["رقم البطاقة القديم", "رقم البطاقة الحديث"]:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
                format_cell_advanced(cell, row[card_key], size_pt=14, font_name="Calibri", align="center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                j += 1
            for i in remaining_indices:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
                if i == 2: set_cell_no_wrap(cell)
                val = row["ت"] if i == 0 else row["رقم البطاقة"] if i == 1 else row["اسم رب الأسرة"] if i == 2 else row["الكلي"] if i == 3 else ""
                cell_align = "left" if i == 2 else "center"
                if i == 2:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i == 3: set_cell_background(cell, "E8F8F5")
                j += 1
        else:
            set_cell_no_wrap(row_cells[2])
            for i in range(16):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                val = row["ت"] if i == 0 else row["رقم البطاقة"] if i == 1 else row["اسم رب الأسرة"] if i == 2 else row["الكلي"] if i == 3 else ""
                cell_align = "left" if i == 2 else "center"
                if i == 2:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i == 0: set_cell_background(cell, current_letter_color)
                    if i == 3: set_cell_background(cell, "E8F8F5")
    return save_doc_buffer(doc, df)

def build_professional_word_report_v5(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base, is_a3=True)
        
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(18), True
    
    headers = ["ت", "اسم رب الأسرة", "عدد الأفراد المستحقة", "حقل كبير فارغ", "حقل كبير فارغ"]
    table = doc.add_table(rows=1, cols=5)
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    table.rows[0].height = Inches(0.75)
    
    col_widths = [Cm(1.5), Cm(8.0), Cm(3.2), Cm(8.0), Cm(8.0)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    COLOR_NAME_BLUE = RGBColor(0, 112, 192)
    COLOR_RED = RGBColor(255, 0, 0)
    
    for i, title in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, title, bold=True, size_pt=18, font_name="Microsoft Uighur", align="center", color_rgb=COLOR_NAVY_BLUE)
            
    prev_letter = None
    current_letter_color = None
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0
        set_cell_no_wrap(row_cells[1])

        for i in range(5):
            cell = row_cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            val = ""
            cell_align = "center"
            text_color = None

            if i == 0: val = row["ت"]
            elif i == 1:
                val = row["اسم رب الأسرة"]
                cell_align = "left"
                text_color = COLOR_RED if is_eligible_zero else COLOR_NAME_BLUE
            elif i == 2: val = "x" if is_eligible_zero else row["مستحق"]
            elif i in [3, 4]: val = "XXXXXXXXXXXX" if is_eligible_zero else ""

            if i == 1:
                format_name_cell_with_small_suffix(cell, val, base_size=16, small_size=10, font_name="Microsoft Uighur", color_rgb=text_color, align=cell_align)
            else:
                format_cell_advanced(cell, val, size_pt=16, font_name="Microsoft Uighur", color_rgb=text_color, align=cell_align)

            if i == 0 and current_letter_color:
                set_cell_background(cell, current_letter_color)

    return save_doc_buffer(doc, df)

def build_professional_word_report_v6(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)
    
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    
    is_combined = card_choice == "القديم والحديث"
    orig_headers = ["ت", card_choice, "اسم المواطن", "العدد المستحق"] + [f"سلة {i}" for i in range(1, 13)]
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    col_widths = [Cm(0.9), Cm(1.8), dynamic_name_width, Cm(1.1)] + [Cm(1.05)] * 12
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    remaining_indices = list(range(2, 16)) if is_combined else None

    headers = (["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]) if is_combined else orig_headers
    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)
    hdr_cells = table.rows[0].cells

    if is_combined:
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 1
        for label in ["القديم", "الحديث"]:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, label, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
        for i in remaining_indices:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            if i >= 3: set_cell_vertical_text(cell)
            format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 2 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = hdr_cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i >= 3: set_cell_vertical_text(cell)
            if i == 1:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 2 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)
            j = 1
            for card_key in ["رقم البطاقة القديم", "رقم البطاقة الحديث"]:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
                format_cell_advanced(cell, row[card_key], size_pt=14, font_name="Calibri", align="center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                j += 1
            for i in remaining_indices:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
                if i == 2: set_cell_no_wrap(cell)
                val = row["ت"] if i == 0 else row["رقم البطاقة"] if i == 1 else row["اسم رب الأسرة"] if i == 2 else row["مستحق"] if i == 3 else ""
                cell_align = "left" if i == 2 else "center"
                if i == 2:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i == 3: set_cell_background(cell, "E8F8F5")
                j += 1
        else:
            set_cell_no_wrap(row_cells[2])
            for i in range(16):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                val = row["ت"] if i == 0 else row["رقم البطاقة"] if i == 1 else row["اسم رب الأسرة"] if i == 2 else row["مستحق"] if i == 3 else ""
                cell_align = "left" if i == 2 else "center"

                if i == 2:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)

                if is_eligible_zero:
                    set_cell_background(cell, "EC7063")
                else:
                    if i == 0: set_cell_background(cell, current_letter_color)
                    if i == 3: set_cell_background(cell, "E8F8F5")

    return save_doc_buffer(doc, df)

# --- النموذج الثامن: مطابق للسادس (العدد المستحق) لكن بـ 8 سلات، وعناوين بلتفاف عادي
# مع حساب عرض كل عمود رياضياً حسب أطول محتوى فيه (بدل عرض ثابت مخمّن) ---
def build_professional_word_report_v8(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)

    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True

    is_combined = card_choice == "القديم والحديث"

    def dynamic_col_width(series, min_chars=1, char_cm=0.22, padding_cm=0.5):
        max_len = max(series.astype(str).str.len().max(), min_chars)
        return Cm(max_len * char_cm + padding_cm)

    ت_width = dynamic_col_width(df["ت"])
    name_width = dynamic_col_width(df["اسم رب الأسرة"], min_chars=15)
    count_width = dynamic_col_width(df["مستحق"])
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)

    orig_headers = ["ت", card_choice, "اسم المواطن", "العدد المستحق"] + [f"سلة {i}" for i in range(1, 9)]
    if is_combined:
        old_card_width = dynamic_col_width(df["رقم البطاقة القديم"])
        new_card_width = dynamic_col_width(df["رقم البطاقة الحديث"])
        col_widths = [ت_width, old_card_width, new_card_width, name_width, count_width] + [Cm(1.05)] * 8
        remaining_indices = list(range(2, 12))
        headers = ["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]
    else:
        card_width = dynamic_col_width(df["رقم البطاقة"])
        col_widths = [ت_width, card_width, name_width, count_width] + [Cm(1.05)] * 8
        remaining_indices = None
        headers = orig_headers

    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)
    hdr_cells = table.rows[0].cells

    if is_combined:
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        cell = hdr_cells[1]
        cell.width, cell.vertical_alignment = old_card_width, WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "القديم", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        cell = hdr_cells[2]
        cell.width, cell.vertical_alignment = new_card_width, WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "الحديث", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 3
        # col_widths here is indexed by NEW layout position, matching remaining_indices' write order
        remaining_widths = [name_width, count_width] + [Cm(1.05)] * 8
        for pos, i in enumerate(remaining_indices):
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = remaining_widths[pos], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 2 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = hdr_cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i == 1:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 2 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)

            cell = row_cells[1]
            cell.width, cell.vertical_alignment = old_card_width, WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["رقم البطاقة القديم"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")

            cell = row_cells[2]
            cell.width, cell.vertical_alignment = new_card_width, WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["رقم البطاقة الحديث"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")

            j = 3
            remaining_widths = [name_width, count_width] + [Cm(1.05)] * 8
            for pos, i in enumerate(remaining_indices):
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = remaining_widths[pos], WD_ALIGN_VERTICAL.CENTER
                if i == 2: set_cell_no_wrap(cell)
                val = row["ت"] if i == 0 else row["رقم البطاقة"] if i == 1 else row["اسم رب الأسرة"] if i == 2 else row["مستحق"] if i == 3 else ""
                cell_align = "left" if i == 2 else "center"
                if i == 2:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                else:
                    if i == 3: set_cell_background(cell, "E8F8F5")
                j += 1
        else:
            set_cell_no_wrap(row_cells[2])
            for i in range(12):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                val = row["ت"] if i == 0 else row["رقم البطاقة"] if i == 1 else row["اسم رب الأسرة"] if i == 2 else row["مستحق"] if i == 3 else ""
                cell_align = "left" if i == 2 else "center"

                if i == 2:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)

                if is_eligible_zero:
                    set_cell_background(cell, "EC7063")
                else:
                    if i == 0: set_cell_background(cell, current_letter_color)
                    if i == 3: set_cell_background(cell, "E8F8F5")

    return save_doc_buffer(doc, df)

# -----------------------------------------------------------------------------
# النموذج التاسع: مطابق 100% لملف Word مرفوع من المستخدم (بدون عمود رقم بطاقة،
# مواد غذائية محددة بألوان عناوين مميزة لكل مادة، عمود تاريخ). لا يستخدم
# bidiVisual عمداً (الملف المصدر لا يستخدمه)، ولا يُطبّق تظليل أحمر للمستحق=صفر
# (الملف المصدر لا يطبّق هذه الميزة على هذا النموذج بالتحديد).
# -----------------------------------------------------------------------------
def build_professional_word_report_v9(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)

    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True

    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    headers = ["ت", "اسم رب الأسرة", "مستحق", "طحين", "سكر", "زيت", "رز", "معجون", "باقوليات", "التاريخ"]
    col_widths = [Cm(1.13), dynamic_name_width, Cm(0.95)] + [Cm(1.52)] * 7

    # لون ونمط كل عنوان مطابق تماماً لتصميم الملف الأصلي المرفوع (لكل مادة لون مميز خاص بها)
    header_styles = {
        "ت": {"color": "2A4B7C", "fill": None, "size": 12, "font": "Segoe UI Semibold"},
        "اسم رب الأسرة": {"color": "2A4B7C", "fill": None, "size": 12, "font": "Segoe UI Semibold"},
        "مستحق": {"color": "00B050", "fill": "DAEEF3", "size": 12, "font": "Calibri"},
        "طحين": {"color": "FF0000", "fill": "E5DFEC", "size": 12, "font": "Segoe UI Semibold"},
        "سكر": {"color": "2A4B7C", "fill": None, "size": 12, "font": "Segoe UI Semibold"},
        "زيت": {"color": "FFC000", "fill": "EAF1DD", "size": 12, "font": "Segoe UI Semibold"},
        "رز": {"color": "2A4B7C", "fill": "DBE5F1", "size": 18, "font": "Segoe UI Semibold"},
        "معجون": {"color": "984806", "fill": "F2DBDB", "size": 12, "font": "Segoe UI Semibold"},
        "باقوليات": {"color": "17365D", "fill": "F2F2F2", "size": 12, "font": "Segoe UI Semibold"},
        "التاريخ": {"color": "17365D", "fill": "F2F2F2", "size": 12, "font": "Segoe UI Semibold"},
    }

    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    table.rows[0].height = Pt(50.4)

    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        cell = hdr_cells[i]
        cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
        style = header_styles[title]
        format_cell_advanced(cell, title, bold=True, size_pt=style["size"], font_name=style["font"], align="center", color_rgb=RGBColor.from_string(style["color"]))
        if style["fill"]:
            set_cell_background(cell, style["fill"])

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter

        new_row = table.add_row()
        new_row.height = Pt(28.8)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        for i in range(len(headers)):
            cell = row_cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i == 0:
                format_cell_advanced(cell, row["ت"], size_pt=14, font_name="Calibri", align="center")
                set_cell_background(cell, current_letter_color)
            elif i == 1:
                set_cell_no_wrap(cell)
                format_cell_advanced(cell, row["اسم رب الأسرة"], size_pt=16, font_name="Calibri", align="right")
            elif i == 2:
                format_cell_advanced(cell, row["مستحق"], size_pt=14, font_name="Calibri", align="center")
                set_cell_background(cell, "E8F8F5")
            else:
                format_cell_advanced(cell, "", size_pt=14, font_name="Calibri", align="center")

    return save_doc_buffer(doc, df)

# --- الدالة الجديدة للنموذج السابع (تفاصيل المواد) ---
def build_professional_word_report_v7(df, filename_base, card_choice, sort_alphabetically=True):
    doc = Document()
    setup_document_layout(doc, filename_base)
    
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"الكشف الإحصائي المنسق للوكيل: {clean_name}")
    title_run.font.name, title_run.font.size, title_run.bold = "Segoe UI Semibold", Pt(14), True
    
    is_combined = card_choice == "القديم والحديث"
    orig_headers = ["ت", "اسم رب الأسرة", card_choice, "الكلي", "المستحق", "المحجوب", "سكر", "زيت", "تمن", "معجون", "فاصوليا", "عدس", "حمص"]
    dynamic_name_width = Cm(max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15) * 0.22 + 0.5)
    # تخصيص مساحات ثابتة للمواد (1.1 سم لكل مادة)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(1.8), Cm(0.8), Cm(0.8), Cm(0.8)] + [Cm(1.1)] * 7
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    remaining_indices = [1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] if is_combined else None

    headers = (["ت", "القديم", "الحديث"] + [orig_headers[i] for i in remaining_indices]) if is_combined else orig_headers
    table = doc.add_table(rows=1, cols=len(headers))
    table.style, table.alignment = 'Table Grid', WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    table.rows[0].height = Inches(0.75)

    if is_combined:
        hdr_cells = table.rows[0].cells
        cell = hdr_cells[0]
        cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
        format_cell_advanced(cell, "ت", bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
        j = 1
        for label in ["القديم", "الحديث"]:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, label, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
        for i in remaining_indices:
            cell = hdr_cells[j]
            cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
            title = orig_headers[i]
            if i >= 3: set_cell_vertical_text(cell)
            format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 1 else "center", color_rgb=COLOR_NAVY_BLUE)
            j += 1
    else:
        for i, title in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # جعل النصوص للأرقام والمواد بشكل عمودي
            if i >= 3: set_cell_vertical_text(cell)
            if i == 2:
                format_header_cell_two_lines(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", color_rgb=COLOR_NAVY_BLUE)
            else:
                format_cell_advanced(cell, title, bold=True, size_pt=12, font_name="Segoe UI Semibold", align="left" if i == 1 else "center", color_rgb=COLOR_NAVY_BLUE)

    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = add_letter_banner_row(table, letter)
                prev_letter = letter
        new_row = table.add_row()
        new_row.height = Inches(0.5)
        row_cells = new_row.cells
        new_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        is_eligible_zero = int(row["مستحق"]) == 0

        if is_combined:
            cell = row_cells[0]
            cell.width, cell.vertical_alignment = col_widths[0], WD_ALIGN_VERTICAL.CENTER
            format_cell_advanced(cell, row["ت"], size_pt=14, font_name="Calibri", align="center")
            if is_eligible_zero: set_cell_background(cell, "EC7063")
            else: set_cell_background(cell, current_letter_color)
            j = 1
            for card_key in ["رقم البطاقة القديم", "رقم البطاقة الحديث"]:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = Cm(1.8), WD_ALIGN_VERTICAL.CENTER
                format_cell_advanced(cell, row[card_key], size_pt=14, font_name="Calibri", align="center")
                if is_eligible_zero: set_cell_background(cell, "EC7063")
                j += 1
            for i in remaining_indices:
                cell = row_cells[j]
                cell.width, cell.vertical_alignment = col_widths[i], WD_ALIGN_VERTICAL.CENTER
                if i == 1: set_cell_no_wrap(cell)
                val = row["ت"] if i == 0 else row["اسم رب الأسرة"] if i == 1 else row["رقم البطاقة"] if i == 2 else row["الكلي"] if i == 3 else row["مستحق"] if i == 4 else row["محجوب"] if i == 5 else ""
                cell_align = "left" if i == 1 else "center"
                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)
                if is_eligible_zero:
                    set_cell_background(cell, "EC7063")
                else:
                    if i == 3: set_cell_background(cell, "EBF5FB")
                    elif i == 4: set_cell_background(cell, "E8F8F5")
                    elif i == 5: set_cell_background(cell, "FADBD8")
                j += 1
        else:
            set_cell_no_wrap(row_cells[1])
            for i in range(13):
                cell = row_cells[i]
                cell.width = col_widths[i]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                val = row["ت"] if i == 0 else row["اسم رب الأسرة"] if i == 1 else row["رقم البطاقة"] if i == 2 else row["الكلي"] if i == 3 else row["مستحق"] if i == 4 else row["محجوب"] if i == 5 else ""
                cell_align = "left" if i == 1 else "center"

                if i == 1:
                    format_name_cell_with_small_suffix(cell, val, base_size=14, small_size=10, font_name="Calibri", color_rgb=None, align="left")
                else:
                    format_cell_advanced(cell, val, size_pt=14, font_name="Calibri", color_rgb=None, align=cell_align)

                if is_eligible_zero:
                    set_cell_background(cell, "EC7063")
                else:
                    if i == 0: set_cell_background(cell, current_letter_color)
                    elif i == 3: set_cell_background(cell, "EBF5FB")
                    elif i == 4: set_cell_background(cell, "E8F8F5")
                    elif i == 5: set_cell_background(cell, "FADBD8")

    return save_doc_buffer(doc, df)

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
# خط Tajawal (أحد أكثر الخطوط العربية استخدامًا في تصميم المواقع والواجهات)
# مضمّن كـ base64 داخل ملف PDF مباشرة لضمان ظهوره بنفس الشكل على أي خادم تشغيل،
# بدل الاعتماد على خطوط النظام المحلي.
# -----------------------------------------------------------------------------
FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")

@lru_cache(maxsize=None)
def _get_embedded_font_base64(filename):
    path = os.path.join(FONTS_DIR, filename)
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except OSError:
        return None

@lru_cache(maxsize=None)
def get_pdf_font_face_css():
    """يبني قواعد @font-face لخط Tajawal (عادي/بولد/إكسترا بولد) إن كانت الملفات متوفرة،
    مع رجوع تلقائي لخطوط النظام إن تعذّر تحميلها."""
    regular = _get_embedded_font_base64("Tajawal-Regular.ttf")
    bold = _get_embedded_font_base64("Tajawal-Bold.ttf")
    extrabold = _get_embedded_font_base64("Tajawal-ExtraBold.ttf")
    if not (regular and bold and extrabold):
        return ""
    return f"""
            @font-face {{
                font-family: 'Tajawal';
                font-weight: 400;
                font-style: normal;
                src: url(data:font/ttf;base64,{regular}) format('truetype');
            }}
            @font-face {{
                font-family: 'Tajawal';
                font-weight: 700;
                font-style: normal;
                src: url(data:font/ttf;base64,{bold}) format('truetype');
            }}
            @font-face {{
                font-family: 'Tajawal';
                font-weight: 800;
                font-style: normal;
                src: url(data:font/ttf;base64,{extrabold}) format('truetype');
            }}
    """

# -----------------------------------------------------------------------------
# المحرك الجديد: إنشاء تقارير PDF
# -----------------------------------------------------------------------------
def _build_report_html_doc(df, filename_base, card_choice, template_choice, sort_alphabetically=True):
    """يبني نص HTML الكامل للكشف (نفس التصميم المستخدم لملف PDF)، بمعزل عن
    خطوة التحويل لـPDF، حتى تُعاد استخدامه أيضاً لتوليد نسخة HTML تفاعلية
    قابلة للتعديل والطباعة مباشرة من المتصفح."""
    clean_name = filename_base
    for w in ["مستكشف", "معدل", "كشف", "منسق", "جاهز", "مدمج"]: clean_name = clean_name.replace(w, "")
    clean_name = " ".join(re.sub(r'[a-zA-Z\-_+_.]', '', clean_name).split())

    total_all = df["الكلي"].astype(int).sum()
    total_eligible = df["مستحق"].astype(int).sum()
    total_withheld = df["محجوب"].astype(int).sum()

    page_orientation = "portrait"
    page_size = "A4"
    is_combined = card_choice == "القديم والحديث"
    card_cols = ["القديم", "الحديث"] if is_combined else [card_choice]

    if template_choice == "النموذج الأول (الأصلي المطور)":
        headers = ["ت"] + card_cols + ["اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", "ملاحظات"] if is_combined else ["ت", "اسم رب الأسرة", "حقل فارغ", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    elif template_choice == "النموذج الثاني (حجم 14 وحقلين فارغين)":
        headers = ["ت"] + card_cols + ["اسم رب الأسرة", "حقل فارغ 1", "حقل فارغ 2", "الكلي", "مستحق", "محجوب", "ملاحظات"] if is_combined else ["ت", "اسم رب الأسرة", "حقل فارغ 1", "حقل فارغ 2", "الكلي", "مستحق", "محجوب", card_choice, "ملاحظات"]
    elif template_choice == "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)":
        headers = ["ت"] + card_cols + ["اسم رب الأسرة", "الكلي", "مستحق", "محجوب", "الشهر الأول", "الشهر الثاني", "الشهر الثالث", "الشهر الرابع"] if is_combined else ["ت", "اسم رب الأسرة", card_choice, "الكلي", "مستحق", "محجوب", "الشهر الأول", "الشهر الثاني", "الشهر الثالث", "الشهر الرابع"]
    elif template_choice == "النموذج الرابع (12 سلة، العدد الكلي)":
        headers = (["ت"] + card_cols + ["اسم المواطن", "العدد الكلي"] if is_combined else ["ت", card_choice, "اسم المواطن", "العدد الكلي"]) + [f"سلة {i}" for i in range(1, 13)]
        page_orientation = "landscape"
    elif template_choice == "النموذج السادس (12 سلة، العدد المستحق)":
        headers = (["ت"] + card_cols + ["اسم المواطن", "العدد المستحق"] if is_combined else ["ت", card_choice, "اسم المواطن", "العدد المستحق"]) + [f"سلة {i}" for i in range(1, 13)]
        page_orientation = "landscape"
    elif template_choice == "النموذج الثامن (8 سلات، العدد المستحق)":
        headers = (["ت"] + card_cols + ["اسم المواطن", "العدد المستحق"] if is_combined else ["ت", card_choice, "اسم المواطن", "العدد المستحق"]) + [f"سلة {i}" for i in range(1, 9)]
        page_orientation = "landscape"
    elif template_choice == "النموذج السابع (تفصيل المواد الغذائية)":
        headers = ["ت"] + card_cols + ["اسم رب الأسرة", "الكلي", "المستحق", "المحجوب", "سكر", "زيت", "تمن", "معجون", "فاصوليا", "عدس", "حمص"] if is_combined else ["ت", "اسم رب الأسرة", card_choice, "الكلي", "المستحق", "المحجوب", "سكر", "زيت", "تمن", "معجون", "فاصوليا", "عدس", "حمص"]
        page_orientation = "landscape"
    elif template_choice == "النموذج التاسع (توزيع مواد غذائية، بدون رقم بطاقة)":
        headers = ["ت", "اسم رب الأسرة", "مستحق", "طحين", "سكر", "زيت", "رز", "معجون", "باقوليات", "التاريخ"]
        page_size = "A4"
        page_orientation = "portrait"
    else:
        headers = ["ت", "اسم رب الأسرة", "عدد الأفراد المستحقة", "حقل كبير فارغ", "حقل كبير فارغ"]
        page_size = "A3"
        page_orientation = "landscape"

    rows_html = ""
    prev_letter = None
    current_letter_color = "D4E6F1"
    for idx, row in df.iterrows():
        if sort_alphabetically:
            letter = get_name_group_letter(row["اسم رب الأسرة"])
            if letter != prev_letter:
                current_letter_color = get_letter_banner_color(letter)
                base_color = get_letter_base_color(letter)
                rows_html += (
                    f'<tr><td colspan="{len(headers)}" '
                    f'style="background-color: #{current_letter_color}; color: #{base_color}; '
                    f'font-weight: bold; font-size: 14pt; text-align: center; padding: 6px 4px;">'
                    f'{letter}</td></tr>'
                )
                prev_letter = letter

        is_eligible_zero = int(row["مستحق"]) == 0
        row_bg = "background-color: #EC7063;" if is_eligible_zero else ""
        ter_bg = f"background-color: #{current_letter_color};" if not is_eligible_zero else ""

        cells_html = ""

        card_vals = [(row["رقم البطاقة القديم"], "font-size: 11.07pt;"), (row["رقم البطاقة الحديث"], "font-size: 11.07pt;")] if is_combined else [(row["رقم البطاقة"], "font-size: 11.07pt;")]

        if template_choice == "النموذج الأول (الأصلي المطور)":
            if is_combined:
                vals = [
                    (row["ت"], ter_bg),
                    *card_vals,
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    ("x" if is_eligible_zero else "", ""),
                    (row["الكلي"], "background-color: #EBF5FB;" if not is_eligible_zero else ""),
                    (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else ""),
                    (row["محجوب"], "background-color: #FADBD8;" if not is_eligible_zero else ""),
                    ("محجوب" if is_eligible_zero else "", "color: #CB4335; font-weight: bold;" if is_eligible_zero else "")
                ]
            else:
                vals = [
                    (row["ت"], ter_bg),
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    ("x" if is_eligible_zero else "", ""),
                    (row["الكلي"], "background-color: #EBF5FB;" if not is_eligible_zero else ""),
                    (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else ""),
                    (row["محجوب"], "background-color: #FADBD8;" if not is_eligible_zero else ""),
                    *card_vals,
                    ("محجوب" if is_eligible_zero else "", "color: #CB4335; font-weight: bold;" if is_eligible_zero else "")
                ]
        elif template_choice == "النموذج الثاني (حجم 14 وحقلين فارغين)":
            if is_combined:
                vals = [
                    (row["ت"], ter_bg),
                    *card_vals,
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    ("x" if is_eligible_zero else "", ""),
                    ("x" if is_eligible_zero else "", ""),
                    (row["الكلي"], "background-color: #EBF5FB;" if not is_eligible_zero else ""),
                    (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else ""),
                    (row["محجوب"], "background-color: #E5E7E9;" if not is_eligible_zero else ""),
                    ("محجوب" if is_eligible_zero else "", "color: #CB4335; font-weight: bold;" if is_eligible_zero else "")
                ]
            else:
                vals = [
                    (row["ت"], ter_bg),
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    ("x" if is_eligible_zero else "", ""),
                    ("x" if is_eligible_zero else "", ""),
                    (row["الكلي"], "background-color: #EBF5FB;" if not is_eligible_zero else ""),
                    (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else ""),
                    (row["محجوب"], "background-color: #E5E7E9;" if not is_eligible_zero else ""),
                    *card_vals,
                    ("محجوب" if is_eligible_zero else "", "color: #CB4335; font-weight: bold;" if is_eligible_zero else "")
                ]
        elif template_choice == "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)":
            if is_combined:
                vals = [
                    (row["ت"], ter_bg),
                    *card_vals,
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    (row["الكلي"], "background-color: #E5E7E9;" if not is_eligible_zero else ""),
                    (row["مستحق"], ""),
                    (row["محجوب"], "background-color: #FCF3CF;" if not is_eligible_zero else ""),
                    ("", ""), ("", ""), ("", ""), ("", "")
                ]
            else:
                vals = [
                    (row["ت"], ter_bg),
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    *card_vals,
                    (row["الكلي"], "background-color: #E5E7E9;" if not is_eligible_zero else ""),
                    (row["مستحق"], ""),
                    (row["محجوب"], "background-color: #FCF3CF;" if not is_eligible_zero else ""),
                    ("", ""), ("", ""), ("", ""), ("", "")
                ]
        elif template_choice == "النموذج الرابع (12 سلة، العدد الكلي)":
            vals = [
                (row["ت"], ter_bg),
                *card_vals,
                (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                (row["الكلي"], "background-color: #E8F8F5;" if not is_eligible_zero else "")
            ] + [("", "")] * 12
        elif template_choice == "النموذج السادس (12 سلة، العدد المستحق)":
            vals = [
                (row["ت"], ter_bg),
                *card_vals,
                (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else "")
            ] + [("", "")] * 12
        elif template_choice == "النموذج الثامن (8 سلات، العدد المستحق)":
            vals = [
                (row["ت"], ter_bg),
                *card_vals,
                (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else "")
            ] + [("", "")] * 8
        elif template_choice == "النموذج السابع (تفصيل المواد الغذائية)":
            if is_combined:
                vals = [
                    (row["ت"], ter_bg),
                    *card_vals,
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    (row["الكلي"], "background-color: #EBF5FB;" if not is_eligible_zero else ""),
                    (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else ""),
                    (row["محجوب"], "background-color: #FADBD8;" if not is_eligible_zero else ""),
                    ("", ""), ("", ""), ("", ""), ("", ""), ("", ""), ("", ""), ("", "")
                ]
            else:
                vals = [
                    (row["ت"], ter_bg),
                    (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                    *card_vals,
                    (row["الكلي"], "background-color: #EBF5FB;" if not is_eligible_zero else ""),
                    (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else ""),
                    (row["محجوب"], "background-color: #FADBD8;" if not is_eligible_zero else ""),
                    ("", ""), ("", ""), ("", ""), ("", ""), ("", ""), ("", ""), ("", "")
                ]
        elif template_choice == "النموذج التاسع (توزيع مواد غذائية، بدون رقم بطاقة)":
            vals = [
                (row["ت"], ter_bg),
                (row["اسم رب الأسرة"], "text-align: right; font-weight: bold; font-size: 14pt;"),
                (row["مستحق"], "background-color: #E8F8F5;" if not is_eligible_zero else "")
            ] + [("", "")] * 7
        else:
            vals = [
                (row["ت"], ter_bg),
                (row["اسم رب الأسرة"], f"text-align: right; font-weight: bold; font-size: 14pt; color: {'#FF0000' if is_eligible_zero else '#0070C0'};"),
                ("x" if is_eligible_zero else row["مستحق"], ""),
                ("XXXXXXXXXXXX" if is_eligible_zero else "", ""),
                ("XXXXXXXXXXXX" if is_eligible_zero else "", "")
            ]

        for val, style in vals:
            cell_style = f"{row_bg} {style}"
            cells_html += f'<td style="{cell_style}">{val}</td>'

        rows_html += f'<tr>{cells_html}</tr>'

    # "كلي/مستحق/محجوب" الثلاثة معاً (النماذج 1، 2، 3، 7) فقط تُعرض بعرض معقول
    # يتناسب مع نصها القصير (بدل ترك العرض التلقائي يوسّعها بلا داع)، مع تضييق
    # عمودي الاسم والملاحظات معها — نص أفقي عادي مألوف، بلا أي دوران؛ محاولة
    # سابقة بتدوير النص 90 درجة أنتجت خطوطاً عربية متكسّرة غير مقروءة فاستُبعدت.
    # لا يُطبَّق أي من هذا على قوالب أخرى (السلال، النموذج التاسع...) تجنباً
    # لأي تغيير غير مطلوب في تصاميمها.
    COMPACT_HEADERS = {"الكلي", "مستحق", "المستحق", "محجوب", "المحجوب"}
    # عرض عمود الاسم يجب أن يتسع لأطول اسم فعلي بالملف دون التفاف (النص مقفل
    # بلا لف)، وإلا انسكب بصرياً على العمود المجاور — عمود مقاس بنص ثابت لا
    # يتحمل نصاً متغير الطول أقصر مما يحتاجه أطول اسم حقيقي بالبيانات.
    name_col_key = "اسم رب الأسرة" if "اسم رب الأسرة" in df.columns else ("اسم المواطن" if "اسم المواطن" in df.columns else None)
    max_name_len = df[name_col_key].astype(str).str.len().max() if name_col_key and not df.empty else 15
    use_compact_columns = any(h in COMPACT_HEADERS for h in headers)

    def _header_cell_html(h):
        pill_extra = " pill-tight" if use_compact_columns else ""
        if "كلي" in h:
            return f'<th><span class="pill{pill_extra} pill-blue">{h}</span></th>'
        if "مستحق" in h:
            return f'<th><span class="pill{pill_extra} pill-green">{h}</span></th>'
        if "محجوب" in h:
            return f'<th><span class="pill{pill_extra} pill-red">{h}</span></th>'
        return f'<th>{h}</th>'

    headers_html = "".join([_header_cell_html(h) for h in headers])

    # أعمدة "حرجة" محتواها كلمة واحدة غير قابلة للف (رقم، اسم عمود، رقم بطاقة)
    # تحصل على نسبة ثابتة مضمونة بصرف النظر عن عدد بقية الأعمدة، وإلا خاطرنا
    # بأن يزاحمها ازدحام الأعمدة (كوضع "القديم والحديث" مع 4 أشهر أو 7 مواد
    # معاً) فتفيض بلا مكان تلف إليه. الأعمدة "المرنة" الباقية (حقول فارغة،
    # أشهر، مواد غذائية) عادة عناوينها قابلة للف على سطرين ولا بيانات حقيقية
    # فيها، فتتقاسم ما تبقى من المساحة بحسب طول عنوانها.
    def _critical_pct(h):
        if h in COMPACT_HEADERS:
            return max(len(h) * 1.3, 11.5) * 0.85
        if h == "ت":
            return 6.0
        if "اسم" in h:
            return min(max(15.0, max_name_len * 1.5), 25.0)
        if h == "ملاحظات":
            return 6.0
        if "بطاق" in h or "تموين" in h or h in ("القديم", "الحديث"):
            return 10.2
        return None

    table_class = ""
    colgroup_html = ""
    if use_compact_columns:
        table_class = "compact"
        critical = [_critical_pct(h) for h in headers]
        critical_total = sum(p for p in critical if p is not None)
        remaining_pct = max(100.0 - critical_total, 15.0)
        flex_weights = [min(max(len(h), 3), 7) * 0.55 if p is None else 0 for h, p in zip(headers, critical)]
        total_flex_weight = sum(flex_weights) or 1
        resolved_widths = [
            p if p is not None else remaining_pct * w / total_flex_weight
            for p, w in zip(critical, flex_weights)
        ]

        # طلب صريح: تقليص عرض كل الأعمدة الأخرى 5% وتحويل المساحة المُحرَّرة
        # لعمود "ملاحظات" تحديداً حتى يظهر بوضوح (كان أضيق عمود سابقاً).
        notes_idx = next((i for i, h in enumerate(headers) if h == "ملاحظات"), None)
        if notes_idx is not None:
            freed = 0.0
            for i in range(len(resolved_widths)):
                if i != notes_idx:
                    cut = resolved_widths[i] * 0.05
                    resolved_widths[i] -= cut
                    freed += cut
            resolved_widths[notes_idx] += freed

        # طلب صريح: تصغير عمود "حقل فارغ" (الحقل بلا بيانات حقيقية) بمقدار
        # 50% من عرضه الطبيعي ثم 30% إضافية فوقها (أي يبقى 35% من الأصل)،
        # وتحويل كل المساحة المُحرَّرة لعمود "ملاحظات" حتى يظهر كاملاً بالورقة.
        blank_indices = [i for i, h in enumerate(headers) if h.startswith("حقل فارغ")]
        if blank_indices and notes_idx is not None:
            freed_blank = 0.0
            for i in blank_indices:
                new_width = max(resolved_widths[i] * 0.35, 4.0)
                freed_blank += resolved_widths[i] - new_width
                resolved_widths[i] = new_width
            resolved_widths[notes_idx] += freed_blank
        colgroup_html = "<colgroup>" + "".join(f'<col style="width:{w:.2f}%">' for w in resolved_widths) + "</colgroup>"
    font_face_css = get_pdf_font_face_css()
    pdf_font_stack = "'Tajawal', 'Segoe UI Semibold', 'Segoe UI', 'Calibri', 'Tahoma', 'Arial', sans-serif"

    html_doc = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="utf-8">
        <style>
            {font_face_css}
            @page {{
                size: {page_size} {page_orientation};
                margin-top: 6mm;
                margin-bottom: 6mm;
                margin-left: 0.4mm;
                margin-right: 1.35mm;
                @bottom-center {{
                    content: "صفحة " counter(page);
                    font-size: 9pt;
                    font-weight: normal;
                    color: #7A8AA3;
                    font-family: {pdf_font_stack};
                }}
            }}
            body {{
                font-family: {pdf_font_stack};
                font-weight: bold;
                direction: rtl;
                margin: 0;
                padding-top: 10mm;
                padding-bottom: 10mm;
                padding-left: 0.5mm;
                padding-right: 1.8mm;
                background-image: radial-gradient(circle, #DCE4F0 1px, transparent 1px);
                background-size: 16px 16px;
            }}
            .invoice-card {{
                background-color: #FFFFFF;
                border: 2.5px solid #1B3A63;
                border-radius: 20px;
                overflow: hidden;
            }}
            .invoice-header {{
                background-color: #1B3A63;
                color: #FFFFFF;
                padding: 14px 20px 12px;
                border-bottom: 4px solid #D4AC0D;
            }}
            .invoice-title {{
                font-size: 19pt;
                font-weight: 800;
                letter-spacing: 0.3px;
                white-space: nowrap;
            }}
            .pill {{
                display: inline-block;
                padding: 3px 12px;
                border-radius: 999px;
                font-weight: 800;
            }}
            .pill-blue {{ background-color: #D6E9FA; color: #1F618D; }}
            .pill-green {{ background-color: #D3F3E8; color: #117864; }}
            .pill-red {{ background-color: #FBD9D3; color: #C0392B; }}
            .pill-tight {{
                padding: 3px 1px;
                font-size: 12.5pt;
            }}
            table.compact {{
                table-layout: fixed;
            }}
            .invoice-subtitle {{
                font-size: 9pt;
                font-weight: normal;
                opacity: 0.85;
                margin-top: 4px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 15.81pt;
            }}
            th {{
                background-color: #EAF0F8;
                color: #1B3A63;
                border: 1px solid #C3D0E3;
                padding: 7px 4px;
                text-align: center;
                font-weight: bold;
            }}
            td {{
                border: 1px solid #C3D0E3;
                padding: 6.44px 3px;
                text-align: center;
                vertical-align: middle;
                white-space: nowrap;
                font-weight: bold;
            }}
            tbody tr:nth-child(even) td {{
                background-color: #F6F9FC;
            }}
            .invoice-footer {{
                background-color: #F4F7FB;
                border-top: 2px solid #1B3A63;
                padding: 12px 20px;
            }}
            .stats {{
                text-align: right;
                font-weight: bold;
                font-size: 12pt;
                color: #1B3A63;
                line-height: 1.6;
            }}
        </style>
    </head>
    <body>
        <div class="invoice-card">
            <div class="invoice-header">
                <div class="invoice-title">الكشف الإحصائي المنسق للوكيل: {clean_name}</div>
                <div class="invoice-subtitle">سجل إلكتروني رسمي — نظام تنسيق كشوفات الوكلاء</div>
            </div>
            <table class="{table_class}">
                {colgroup_html}
                <thead>
                    <tr>{headers_html}</tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
            <div class="invoice-footer">
                <div class="stats">
                    العدد الكلي للافراد = {total_all}<br>
                    العدد الكلي للمستحقين = {total_eligible}<br>
                    العدد الكلي للمحجوبين = {total_withheld}
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    return html_doc, page_size, page_orientation, headers


def build_pdf_report(df, filename_base, card_choice, template_choice, sort_alphabetically=True):
    html_doc, page_size, page_orientation, _headers = _build_report_html_doc(
        df, filename_base, card_choice, template_choice, sort_alphabetically
    )

    if WEASYPRINT_AVAILABLE:
        pdf_buffer = BytesIO()
        HTML(string=html_doc).write_pdf(pdf_buffer)
        pdf_buffer.seek(0)
        return pdf_buffer
    elif PDFKIT_AVAILABLE:
        options = {
            'page-size': page_size,
            'orientation': 'Landscape' if page_orientation == 'landscape' else 'Portrait',
            'encoding': 'UTF-8',
            'margin-top': '8mm',
            'margin-right': '6mm',
            'margin-bottom': '8mm',
            'margin-left': '6mm',
            'custom-header': [('Accept-Encoding', 'gzip')],
            'no-outline': None
        }
        pdf_bytes = pdfkit.from_string(html_doc, False, options=options)
        pdf_buffer = BytesIO(pdf_bytes)
        return pdf_buffer
    else:
        raise Exception("لا توجد مكتبة PDF مثبتة (weasyprint أو pdfkit).")

# -----------------------------------------------------------------------------
# واجهة استخدام التطبيق (Streamlit Interface)
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 رفع الكشف المراد تدقيقه وتنسيقه للمطبعة</h3>", unsafe_allow_html=True)
uploaded_files = st.file_uploader("ارفع كشف الوكلاء (يمكنك رفع أكثر من ملف)", type=['docx', 'xlsx'], accept_multiple_files=True, key="doc_input_v8", label_visibility="collapsed")

merge_choice = st.radio(
    "📎 عند رفع أكثر من ملف:",
    ["معالجة كل ملف بشكل منفرد", "دمج كل الملفات في ملف واحد وترتيبها"],
    index=0,
    horizontal=True
)
merge_files = (merge_choice == "دمج كل الملفات في ملف واحد وترتيبها")

st.markdown("<br>", unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 1, 2])

with col1:
    selected_card = st.radio(
        "📄 اختر نوع رقم البطاقة:",
        ["رقم البطاقة القديم", "رقم البطاقة الحديث", "القديم والحديث"],
        index=0,
        horizontal=False
    )

with col2:
    name_length_choice = st.radio(
        "👤 طول اسم رب الأسرة:",
        ["الاسم الثلاثي فقط", "الاسم الرباعي (إن وجد)"],
        index=0,
        horizontal=False
    )

with col3:
    template_choice = st.radio(
        "🎨 اختر نموذج قالب الـ Word المطلوب:",
        [
            "النموذج الأول (الأصلي المطور)", 
            "النموذج الثاني (حجم 14 وحقلين فارغين)",
            "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)",
            "النموذج الرابع (12 سلة، العدد الكلي)",
            "النموذج الخامس (ورقة A3، حقول كبيرة، خط Uighur)",
            "النموذج السادس (12 سلة، العدد المستحق)",
            "النموذج السابع (تفصيل المواد الغذائية)",
            "النموذج الثامن (8 سلات، العدد المستحق)",
            "النموذج التاسع (توزيع مواد غذائية، بدون رقم بطاقة)"
        ],
        index=0,
        horizontal=False
    )

sort_choice = st.radio(
    "🔤 ترتيب بيانات الجدول:",
    ["ترتيب أبجدي بحسب الاسم", "الحفاظ على ترتيب الملف الأصلي (بدون ترتيب أبجدي)"],
    index=0,
    horizontal=True
)
sort_alphabetically = (sort_choice == "ترتيب أبجدي بحسب الاسم")

st.markdown("<br>", unsafe_allow_html=True)

if uploaded_files:
    current_filenames = [f.name for f in uploaded_files]
    if (st.session_state.uploaded_filenames != current_filenames or
        st.session_state.selected_card != selected_card or
        st.session_state.template_choice != template_choice or
        st.session_state.name_choice != name_length_choice or
        st.session_state.sort_choice != sort_choice or
        st.session_state.merge_choice != merge_choice):
        st.session_state.processing_done = False

if st.button("⚙️ تشغيل محرك التنظيم والتنسيق المتقدم الكلي"):
    if uploaded_files:
        order_desc = 'وترتيب القيود أبجدياً ' if sort_alphabetically else '(بترتيب الملف الأصلي) '
        mode_desc = 'ودمجها بملف واحد' if merge_files else 'لكل ملف على حدة'
        with st.spinner(f'جاري معالجة {order_desc}{mode_desc} وإعداد التنسيق الشرطي والمقاييس...'):
            try:
                extracted = []
                file_warnings = {}
                for f in uploaded_files:
                    df_res, extraction_warnings = extract_and_clean_data(f, selected_card, name_length_choice, sort_alphabetically)
                    if not df_res.empty:
                        extracted.append({"filename": f.name.rsplit('.', 1)[0], "df": df_res})
                        log_processed_file(f.name, len(df_res), selected_card, name_length_choice, template_choice, sort_choice)
                        if extraction_warnings:
                            file_warnings[f.name] = extraction_warnings

                if extracted:
                    if merge_files:
                        merged_df = pd.concat([e["df"] for e in extracted], ignore_index=True)
                        if sort_alphabetically:
                            merged_df = merged_df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
                        merged_df["ت"] = merged_df.index + 1
                        if len(extracted) > 1:
                            merged_filename = "مدمج_" + "_".join([e["filename"][:10] for e in extracted])
                        else:
                            merged_filename = extracted[0]["filename"]
                        results = [{"filename": merged_filename, "df": merged_df}]
                    else:
                        results = extracted

                    st.session_state.results = results
                    st.session_state.uploaded_filenames = [f.name for f in uploaded_files]
                    st.session_state.selected_card = selected_card
                    st.session_state.template_choice = template_choice
                    st.session_state.name_choice = name_length_choice
                    st.session_state.sort_choice = sort_choice
                    st.session_state.merge_choice = merge_choice
                    st.session_state.extraction_warnings = file_warnings
                    st.session_state.processing_done = True
                else:
                    st.error("لم يتم العثور على بيانات جداول متوافقة في الملفات المرفوعة.")
            except Exception as e:
                st.error(f"خطأ غير متوقع: {e}")
    else:
        st.warning("الرجاء رفع ملف docx أو xlsx أولاً.")

if st.session_state.processing_done:
    used_card_type = st.session_state.selected_card
    used_template = st.session_state.template_choice
    used_sort_alphabetically = (st.session_state.sort_choice == "ترتيب أبجدي بحسب الاسم")
    order_note = "أبجدياً" if used_sort_alphabetically else "بترتيب الملف الأصلي"
    results = st.session_state.results
    mode_note = "تم دمج الملفات المرفوعة بملف واحد" if st.session_state.merge_choice == "دمج كل الملفات في ملف واحد وترتيبها" else f"تمت معالجة {len(results)} ملف بشكل منفصل"

    st.success(f"✅ {mode_note} بنجاح ({order_note}).")

    for fname, warns in st.session_state.extraction_warnings.items():
        for w in warns:
            st.warning(f"⚠️ تنبيه فحص بيانات — ملف \"{fname}\": {w}")

    def build_word_for_template(df_final, output_filename, used_card_type, used_template, used_sort_alphabetically=True):
        if used_template == "النموذج الأول (الأصلي المطور)":
            return build_professional_word_report(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج الثاني (حجم 14 وحقلين فارغين)":
            return build_professional_word_report_v2(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)":
            return build_professional_word_report_v3(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج الرابع (12 سلة، العدد الكلي)":
            return build_professional_word_report_v4(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج السادس (12 سلة، العدد المستحق)":
            return build_professional_word_report_v6(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج السابع (تفصيل المواد الغذائية)":
            return build_professional_word_report_v7(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج الثامن (8 سلات، العدد المستحق)":
            return build_professional_word_report_v8(df_final, output_filename, used_card_type, used_sort_alphabetically)
        elif used_template == "النموذج التاسع (توزيع مواد غذائية، بدون رقم بطاقة)":
            return build_professional_word_report_v9(df_final, output_filename, used_card_type, used_sort_alphabetically)
        else:
            return build_professional_word_report_v5(df_final, output_filename, used_card_type, used_sort_alphabetically)

    for idx, item in enumerate(results):
        df_final = item["df"]
        output_filename = item["filename"]

        st.markdown(f"---\n#### 📄 {output_filename} — ({len(df_final)}) قيد اسم")

        with st.spinner(f'جاري صياغة وهيكلة مستندات Word و PDF لملف "{output_filename}"...'):
            word_output = build_word_for_template(df_final, output_filename, used_card_type, used_template, used_sort_alphabetically)

        dl_col1, dl_col2 = st.columns(2)

        with dl_col1:
            st.download_button(
                label="📄 تحميل الكشف المنسق (Word)",
                data=word_output,
                file_name=f"كشف_منسق_جاهز_{output_filename}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"word_dl_{idx}",
            )

        with dl_col2:
            if PDFKIT_AVAILABLE or WEASYPRINT_AVAILABLE:
                try:
                    pdf_output = build_pdf_report(df_final, output_filename, used_card_type, used_template, used_sort_alphabetically)
                    st.download_button(
                        label="📕 تحميل الكشف المنسق (PDF جاهز للطباعة)",
                        data=pdf_output,
                        file_name=f"كشف_منسق_جاهز_{output_filename}.pdf",
                        mime="application/pdf",
                        key=f"pdf_dl_{idx}",
                    )
                except Exception as e:
                    st.error(f"حدث خطأ أثناء إعداد PDF لملف \"{output_filename}\": {e}")
            else:
                st.warning("⚠️ يرجى تثبيت مكتبة `pdfkit` أو `weasyprint` لتفعيل خاصية تحميل PDF.")
