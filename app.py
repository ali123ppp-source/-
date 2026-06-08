import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# إعدادات واجهة المستخدم للـ نظام الجديد
st.set_page_config(page_title="نظام تنسيق وتدقيق كشوفات الوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #1ABC9C; color: white; width: 100%; font-weight: bold; border-radius: 8px; font-size: 18px;}
    .report-box { background-color: #F8F9F9; padding: 15px; border-radius: 8px; border-right: 5px solid #1ABC9C; text-align: right; margin-bottom: 10px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام تنسيق وتدقيق كشوفات الوكلاء الاحترافي 📄✨</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>يقوم النظام بترتيب الأسماء أبجدياً، عزل بطاقات المحجوبين (الذين أفرادهم صفر)، وتوليد كشف Word منسق وملون بالكامل طبقاً للمواصفات القياسية.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# إدارة حالة الجلسة (Session State) لمنع اختفاء البيانات عند التحميل
# -----------------------------------------------------------------------------
if "processing_done" not in st.session_state:
    st.session_state.processing_done = False
    st.session_state.df_final = None
    st.session_state.output_filename = ""

# -----------------------------------------------------------------------------
# مساعدات تنسيق ملف الـ Word (التلوين والخطوط ودعم اللغة العربية)
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    """تلوين خلفية الخلايا بدقة"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_cell_arabic(cell, text, bold=False, color_rgb=None, size_pt=16, font_name="Calibri"):
    """تنسيق النصوص العربية داخل الخلايا وإجبار الوورد على قراءة الخط والحجم المخصص للغة العربية"""
    cell.text = str(text)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT  # محاذاة لليمين
    
    # تفعيل خاصية الكتابة من اليمين لليسار (RTL) للفقرة
    pPr = p.paragraph_format.element.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    
    for run in p.runs:
        run.bold = bold
        if color_rgb:
            run.font.color.rgb = color_rgb
            
        # إجبار نظام وورد على تطبيق الخط على النصوص العربية (Complex Scripts)
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)  # cs تعني Complex Scripts (العربية)
        rPr.append(rFonts)
        
        run.font.size = Pt(size_pt)

# -----------------------------------------------------------------------------
# محرك استخراج وتنظيف البيانات من الملف المرفوع
# -----------------------------------------------------------------------------
def extract_and_clean_data(file_obj):
    doc = Document(file_obj)
    raw_records = []
    
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            
            # تخطي الترويسات والعناوين
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
                continue
            
            # تحديد حقل الاسم العربي
            name_idx = -1
            max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len:
                        max_len = len(c)
                        name_idx = i
            
            if name_idx == -1: continue
            
            # رصد أرقام البطاقات (أرقام >= 5 خانات)
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices: continue
            
            # الاعتماد على رقم البطاقة القديم فقط (تجاهل الحديث إذا وُجد)
            old_card_num = cells[card_indices[0]]
            
            # استخراج الأرقام الحسابية قبل الاسم
            digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
            if len(digit_cells) >= 3:
                withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
            elif len(digit_cells) == 2:
                withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
            else:
                continue
                
            raw_records.append({
                "اسم رب الأسرة": cells[name_idx],
                "رقم البطاقة القديم": old_card_num,
                "الكلي": total,
                "محجوب": withheld,
                "مستحق": eligible
            })
            
    # تحويل البيانات إلى DataFrame للترتيب الأبجدي والمعالجة الحسابية
    df = pd.DataFrame(raw_records)
    if not df.empty:
        # 1- ترتيب الأسماء أبجدياً مع الحفاظ على البيانات المرتبطة
        df = df.sort_values(by="اسم رب الأسرة").reset_index(drop=True)
        # توليد التسلسل التلقائي الجديد بعد الترتيب الأبجدي
        df.insert(0, "ت", df.index + 1)
    return df

# -----------------------------------------------------------------------------
# محرك توليد التقرير النهائي المنسق والمطابق لشروط التلوين والقياسات
# -----------------------------------------------------------------------------
def build_professional_word_report(df):
    doc = Document()
    
    # إعداد هوامش الصفحة لتستوعب الجدول العريض
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        
    # حقول العناوين بالترتيب المطلوب
    headers = ["ت", "اسم رب الأسرة", "حقل فارغ", "رقم البطاقة القديم", "الكلي", "محجوب", "مستحق", "ملاحظات"]
    
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # حساب عرض عمود الاسم ديناميكياً بناءً على أطول اسم متواجد بالبيانات
    max_name_chars = max(df["اسم رب الأسرة"].astype(str).str.len().max(), 15)
    dynamic_name_width = Cm(max(4.5, max_name_chars * 0.22))
    
    # 3- تنسيق صف العناوين الرئيسي وحجم الخط 16 Calibri
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        format_cell_arabic(hdr_cells[i], title, bold=True, size_pt=16)
        
        # تعيين القياسات المحددة بدقة للسنتيمتر لعناوين الجدول
        if i == 0: hdr_cells[i].width = Cm(1.0)
        elif i == 1: hdr_cells[i].width = dynamic_name_width
        elif i == 2: hdr_cells[i].width = Cm(0.25)   # حقل فارغ ربع سم
        elif i == 3: hdr_cells[i].width = Cm(3.0)    # رقم البطاقة تلقائي متناسق مع الحجم
        elif i == 4: hdr_cells[i].width = Cm(1.5)
        elif i == 5: hdr_cells[i].width = Cm(1.5)
        elif i == 6: hdr_cells[i].width = Cm(1.5)
        elif i == 7: hdr_cells[i].width = Cm(0.5)    # ملاحظات نصف سم
        
    # الأكواد الستة عشرية (Hex) للألوان المطلوبة
    HEX_YELLOW = "FFFF00"       # أصفر فاقع/كامل لعمود الترقيم
    HEX_LIGHT_BLUE = "DDEBF7"   # أزرق خافت لعمود الكلي
    HEX_LIGHT_GREEN = "E2EFDA"  # أخضر فاتح لعمود المستحق
    HEX_LIGHT_RED = "FCE4D6"    # أحمر/برتقالي فاتح لعمود المحجوبين الكلي
    HEX_ROW_RED = "FADBD8"      # تظليل السجل بالكامل باللون الأحمر للحالات التي مجموع أفرادها صفر
    
    # تعبئة البيانات وتطبيق التنسيق الشرطي
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        
        # 2- فحص عمود الأفراد الكلية إذا كان صفر
        is_total_zero = int(row["الكلي"]) == 0
        
        # تطبيق قياسات الأعمدة على صفوف البيانات لضمان ثبات الهيكل
        row_cells[0].width = Cm(1.0)
        row_cells[1].width = dynamic_name_width
        row_cells[2].width = Cm(0.25)
        row_cells[3].width = Cm(3.0)
        row_cells[4].width = Cm(1.5)
        row_cells[5].width = Cm(1.5)
        row_cells[6].width = Cm(1.5)
        row_cells[7].width = Cm(0.5)
        
        for i in range(8):
            val = ""
            text_color = None
            
            if i == 0: val = row["ت"]
            elif i == 1: val = row["اسم رب الأسرة"]
            elif i == 2: val = "x" if is_total_zero else ""  # وضع علامة x في الحقل الفارغ للمحجوب
            elif i == 3: val = row["رقم البطاقة القديم"]
            elif i == 4: val = row["الكلي"]
            elif i == 5: val = row["محجوب"]
            elif i == 6: val = row["مستحق"]
            elif i == 7: 
                val = "محجوب" if is_total_zero else ""
                if is_total_zero:
                    text_color = RGBColor(255, 0, 0)  # كلمة محجوب باللون الأحمر
            
            format_cell_arabic(row_cells[i], val, size_pt=16, color_rgb=text_color)
            
            # قواعد التظليل والتلوين الشرطي للجدول
            if is_total_zero:
                # تظليل السجل بالكامل باللون الأحمر للمحجوبين
                set_cell_background(row_cells[i], HEX_ROW_RED)
            else:
                # التلوين القياسي للأعمدة في الحالات الطبيعية
                if i == 0: set_cell_background(row_cells[i], HEX_YELLOW)        # عمود الترقيم أصفر
                elif i == 4: set_cell_background(row_cells[i], HEX_LIGHT_BLUE)   # عمود الكلي أزرق خافت
                elif i == 5: set_cell_background(row_cells[i], HEX_LIGHT_RED)    # عمود المحجوب أحمر فاتح
                elif i == 6: set_cell_background(row_cells[i], HEX_LIGHT_GREEN)  # عمود المستحق أخضر فاتح

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة الاستخدام (Streamlit Layout)
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 ارفع الكشف المطلوب معالجته وتنسيقه</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("ارفع ملف الوورد المُراد تنظيم بياناته", type=['docx'], key="doc_input", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

# في حال تغيير الملف، نقوم بتصفير النتائج السابقة تلقائياً
if uploaded_file:
    current_filename = uploaded_file.name.rsplit('.', 1)[0]
    if st.session_state.output_filename != current_filename:
        st.session_state.processing_done = False

if st.button("🚀 بدء المعالجة والتنسيق الاحترافي الشامل"):
    if uploaded_file:
        with st.spinner('جاري قراءة الكشف، فرز الأسماء أبجدياً وتطبيق التنسيق الشرطي...'):
            try:
                df_res = extract_and_clean_data(uploaded_file)
                
                if not df_res.empty:
                    st.session_state.df_final = df_res
                    st.session_state.output_filename = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state.processing_done = True
                else:
                    st.error("عذراً، لم يتم العثور على جداول بيانات متوافقة داخل الملف المرفوع.")
                    st.session_state.processing_done = False
            except Exception as e:
                st.error(f"حدث خطأ غير متوقع أثناء معالجة المستند: {e}")
    else:
        st.warning("الرجاء رفع ملف الكشف بامتداد docx أولاً لتشغيل النظام.")

# -----------------------------------------------------------------------------
# عرض النتائج وتحميل التقرير النهائي (مستقل لضمان ثبات الواجهة)
# -----------------------------------------------------------------------------
if st.session_state.processing_done:
    df_final = st.session_state.df_final
    output_filename = st.session_state.output_filename
    
    st.success(f"📊 تم الانتهاء من المعالجة بنجاح! إجمالي العوائل المدرجة: {len(df_final)} عائلة مرتبة أبجدياً.")
    
    # عرض عينة سريعة تفاعلية للمستخدم على المتصفح
    st.markdown("<h3 style='text-align: right; color: #1ABC9C;'>📋 عينة من البيانات المرتبة والمنقحة</h3>", unsafe_allow_html=True)
    st.dataframe(df_final, use_container_width=True, hide_index=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # توليد ملف الوورد النهائي للتحميل بضغطة زر واحدة بدون مشاكل الـ Rerun
    with st.spinner('جاري بناء وتلوين ملف Word النهائي...'):
        word_output = build_professional_word_report(df_final)
        
    st.download_button(
        label="📥 تحميل كشف الوكلاء المنسق والمعدل بالكامل (Word)",
        data=word_output,
        file_name=f"كشف_منسق_{output_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )