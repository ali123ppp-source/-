import streamlit as st
import pandas as pd
import re
import io
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 1. إعدادات الصفحة والواجهة
# ==========================================
st.set_page_config(page_title="نظام الوكلاء - استخراج الكشوفات", page_icon="📄", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. الدوال المساعدة للتعامل مع ملفات Word (Helpers)
# ==========================================
def set_cell_background(cell, color_hex):
    """تلوين خلفية الخلية"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_cell_advanced(cell, text, bold=False, size_pt=12, font_name="Calibri", align="center", color_rgb=None):
    """تنسيق النص داخل الخلية"""
    cell.text = str(text)
    for paragraph in cell.paragraphs:
        if align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "left":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(size_pt)
            run.bold = bold
            if color_rgb:
                run.font.color.rgb = color_rgb

def set_table_borders(table, color_hex="000000"):
    """رسم حدود الجدول"""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color_hex)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def set_cell_no_wrap(cell):
    """منع التفاف النص داخل الخلية"""
    tcPr = cell._tc.get_or_add_tcPr()
    noWrap = parse_xml(f'<w:noWrap {nsdecls("w")}/>')
    tcPr.append(noWrap)

def save_doc_buffer(doc, df):
    """حفظ المستند في الذاكرة لتصديره"""
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 3. محرك استخراج وتنظيف البيانات
# ==========================================
def extract_and_clean_data(df):
    """
    (قم بوضع الكود الخاص بك هنا لتنظيف البيانات، 
    حالياً نعتبر أن البيانات تأتي بأسماء الأعمدة المطلوبة)
    """
    # مثال مبسط للتأكد من وجود الأعمدة
    required_cols = ["ت", "اسم رب الأسرة", "الكلي", "مستحق", "محجوب"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = "0"
            
    if "رقم البطاقة" not in df.columns and "الرقم التمويني" not in df.columns:
         df["رقم البطاقة"] = "---"
         
    return df

# ==========================================
# 4. محركات بناء تقارير Word (النماذج V1 إلى V4 - مبسطة كمثال)
# ==========================================
def build_professional_word_report(df, filename_base, card_choice):
    doc = Document()
    doc.add_heading(f'النموذج الأول: {filename_base}', level=1)
    return save_doc_buffer(doc, df)

def build_professional_word_report_v2(df, filename_base, card_choice):
    doc = Document()
    doc.add_heading(f'النموذج الثاني: {filename_base}', level=1)
    return save_doc_buffer(doc, df)

def build_professional_word_report_v3(df, filename_base, card_choice):
    doc = Document()
    doc.add_heading(f'النموذج الثالث: {filename_base}', level=1)
    return save_doc_buffer(doc, df)

def build_professional_word_report_v4(df, filename_base, card_choice):
    doc = Document()
    doc.add_heading(f'النموذج الرابع: {filename_base}', level=1)
    return save_doc_buffer(doc, df)

# ==========================================
# 5. النموذج الخامس الجديــــد (نموذج التسليم مع البصمة)
# ==========================================
def build_professional_word_report_v5(df, filename_base, card_choice):
    doc = Document()
    
    # تصغير الهوامش لاستغلال مساحة الورقة
    for section in doc.sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(0.3)
        section.right_margin = Cm(0.3)
        
    # تنظيف اسم الوكيل للعنوان
    clean_name = filename_base
    words_to_remove = ["مستكشف", "معدل", "كشف", "منسق", "جاهز", ".xlsx", ".csv"]
    for w in words_to_remove:
        clean_name = clean_name.replace(w, "")
    clean_name = re.sub(r'[a-zA-Z]', '', clean_name)
    clean_name = re.sub(r'[\-_+_.]', '', clean_name)
    clean_name = " ".join(clean_name.split())
    
    # إضافة العنوان
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    title_run = title_p.add_run(f"كشف تسليم المواد للوكيل: {clean_name}")
    title_run.font.name = "Segoe UI Semibold"
    title_run.font.size = Pt(14)
    title_run.bold = True
    
    # إعداد الجدول الأساسي
    headers = ["ت", "اسم رب الأسرة", card_choice, "الكلي", "مستحق", "محجوب", "توقيع المستلم / البصمة"]
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    set_table_borders(table, color_hex="2A4B7C")
    table._tbl.tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    table.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    # حساب المسافات والأعمدة
    max_name_len = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max_name_len * 0.22 + 0.5)
    
    # حقل البصمة واسع (4 سم)
    col_widths = [Cm(0.9), dynamic_name_width, Cm(3.0), Cm(1.2), Cm(1.2), Cm(1.2), Cm(4.0)]
    COLOR_NAVY_BLUE = RGBColor(42, 75, 124)
    
    # تنسيق عناوين الجدول
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        cell_align = "left" if i == 1 else "center"
        format_cell_advanced(hdr_cells[i], title, bold=True, size_pt=13, font_name="Segoe UI Semibold", align=cell_align, color_rgb=COLOR_NAVY_BLUE)
            
    HEX_ELEGANT_BLUE = "D4E6F1"
    HEX_ALERT_RED = "EC7063"
    
    # تعبئة الجدول بالبيانات
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        r_trPr = table.rows[idx+1]._tr.get_or_add_trPr()
        r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        
        # حماية من الأخطاء في حال كانت القيمة فارغة
        try:
            is_eligible_zero = int(float(row.get("مستحق", 0))) == 0
        except ValueError:
            is_eligible_zero = False
            
        set_cell_no_wrap(row_cells[1])
        
        for i in range(7):
            row_cells[i].width = col_widths[i]
            
            val = ""
            if i == 0: val = row.get("ت", str(idx+1))
            elif i == 1: val = row.get("اسم رب الأسرة", "")
            elif i == 2: val = row.get(card_choice, "")
            elif i == 3: val = row.get("الكلي", "")
            elif i == 4: val = row.get("مستحق", "")
            elif i == 5: val = row.get("محجوب", "")
            elif i == 6: val = "" # حقل التوقيع/البصمة يترك فارغاً
                    
            format_cell_advanced(row_cells[i], val, size_pt=14, font_name="Calibri", color_rgb=None, align="left" if i == 1 else "center")
            
            # تلوين المحجوب / غير المستحق بالأحمر
            if is_eligible_zero and i != 6:
                set_cell_background(row_cells[i], HEX_ALERT_RED)
            else:
                if i == 0: set_cell_background(row_cells[i], HEX_ELEGANT_BLUE)

    return save_doc_buffer(doc, df)


# ==========================================
# 6. واجهة المستخدم (Streamlit UI)
# ==========================================
st.title("📄 نظام الوكلاء - استخراج الكشوفات المطور")
st.write("قم برفع ملف الإكسل لاختيار القالب المناسب وتصدير ملف وورد جاهز للطباعة.")

uploaded_file = st.file_uploader("📂 ارفع ملف الإكسل هنا", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    # قراءة البيانات
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        
    st.success("تم قراءة الملف بنجاح!")
    
    col1, col2 = st.columns(2)
    
    with col1:
         used_card_type = st.selectbox("💳 اختر نوع البطاقة:", ["رقم البطاقة", "الرقم التمويني"])
         
    with col2:
        # القائمة المحدثة تحتوي على النموذج الخامس
        used_template = st.radio(
            "🎨 اختر نموذج قالب الـ Word المطلوب:",
            [
                "النموذج الأول (الأصلي المطور)", 
                "النموذج الثاني (حجم 14 وحقلين فارغين)",
                "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)",
                "النموذج الرابع (12 سلة، العدد الكلي)",
                "النموذج الخامس (نموذج تسليم مع حقل التوقيع/البصمة)" # <== النموذج الجديد هنا
            ],
            index=4, # تم جعل النموذج الخامس هو الافتراضي (يمكنك تغييره لـ 0)
            horizontal=False
        )

    if st.button("🚀 بناء وتصدير ملف الوورد"):
        with st.spinner('جاري معالجة البيانات وبناء مستند Word...'):
            # تنظيف البيانات
            df_final = extract_and_clean_data(df)
            output_filename = uploaded_file.name.split('.')[0]
            
            # الشروط للتشغيل بناءً على القالب المختار
            if used_template == "النموذج الأول (الأصلي المطور)":
                word_output = build_professional_word_report(df_final, output_filename, used_card_type)
            elif used_template == "النموذج الثاني (حجم 14 وحقلين فارغين)":
                word_output = build_professional_word_report_v2(df_final, output_filename, used_card_type)
            elif used_template == "النموذج الثالث (خط 16، عناوين 12، 4 أشهر)":
                word_output = build_professional_word_report_v3(df_final, output_filename, used_card_type)
            elif used_template == "النموذج الرابع (12 سلة، العدد الكلي)":
                word_output = build_professional_word_report_v4(df_final, output_filename, used_card_type)
            else: 
                # استدعاء النموذج الخامس هنا
                word_output = build_professional_word_report_v5(df_final, output_filename, used_card_type)
                
            st.success("✅ تم الانتهاء من بناء الملف بنجاح!")
            
            st.download_button(
                label="📥 تحميل ملف الوورد الجاهز",
                data=word_output,
                file_name=f"كشف_{output_filename}_مطور.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
