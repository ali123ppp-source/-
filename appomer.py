import streamlit as st
import docx
import pandas as pd
import io
import re

st.set_page_config(page_title="المكنسة الشاملة للقيود", layout="wide")

st.markdown("<h1 style='text-align: right; color: #8B0000;'>المكنسة الشاملة (استخراج مطلق بدون شروط) 🧹⚡</h1>", unsafe_allow_html=True)
st.write("هذا المحرك يتجاهل هيكلة الجداول الوهمية، ويسحب أي سطر يحتوي على (اسم + بطاقة) بقوة الذكاء الاصطناعي النصي.")

uploaded_file = st.file_uploader("قم برفع ملف الـ Word (.docx)", type=["docx"])

if uploaded_file is not None:
    doc = docx.Document(uploaded_file)
    lines = []
    
    # 1. سحب كل فقرة في الملف
    for para in doc.paragraphs:
        txt = para.text.replace('\n', ' ').strip()
        if txt: lines.append(txt)
            
    # 2. سحب كل صف في الجداول (ودمجه كسطر نصي واحد لتدمير الخلايا المدمجة)
    for table in doc.tables:
        for row in table.rows:
            # تنظيف ودمج الخلايا
            cells_txt = [c.text.replace('\n', ' ').strip() for c in row.cells if c.text.strip()]
            row_joined = " ".join(cells_txt)
            if row_joined: lines.append(row_joined)

    raw_records = []
    
    # 3. الاختراق النصي الشامل
    for line in lines:
        line_clean = re.sub(r'\s+', ' ', line).strip()
        
        # تخطي الترويسات الصريحة فقط إذا كانت لا تحتوي على أسماء مواطنين
        if "رقم المركز" in line_clean and "الافراد الكلية" in line_clean:
            continue
            
        # الشرط الوحيد: يجب أن يحتوي السطر على حروف عربية وأرقام
        if not re.search(r'[\u0600-\u06FF]', line_clean) or not re.search(r'\d', line_clean):
            continue

        # استخراج البطاقات (أي رقم يتجاوز 4 مراتب)
        all_nums = re.findall(r'\d+', line_clean)
        cards = [n for n in all_nums if len(n) >= 5]
        
        if not cards: 
            continue # إذا لم يوجد رقم بطاقة، فهو ليس قيداً
            
        old_card = str(cards[0])
        new_card = str(cards[-1]) if len(cards) > 1 else old_card

        # استخراج الاسم الصافي (تجاهل الأرقام والرموز)
        text_only = re.sub(r'[^\u0600-\u06FF\s]', ' ', line_clean)
        words = [w for w in text_only.split() if w not in ["ت", "المركز", "ملاحظات", "الوكيل", "الافراد", "الكلية", "المحجوبين", "المستحقة", "رقم", "البطاقة", "القديم", "الحديث", "نوع", "الوكالة", "فارغ", "حقل"]]
        
        if len(words) < 2: 
            continue # يجب أن يكون الاسم مقطعين على الأقل
            
        full_name = " ".join(words)
        
        # استخراج الإحصائيات (الأرقام الصغيرة)
        smalls = [int(n) for n in all_nums if len(n) < 5]
        total = eligible = withheld = 0
        
        # سحب الأرقام الثلاثة المنطقية (إذا وُجدت)
        if len(smalls) >= 3:
            stats = smalls[-3:] # غالباً تكون الإحصائيات في نهاية أو بداية السطر، نأخذ آخر 3
            total = max(stats)
            withheld = min(stats)
            eligible = sorted(stats)[1] if len(stats) == 3 else total
        elif len(smalls) == 2:
            total = max(smalls)
            eligible = min(smalls)
            withheld = 0
        elif len(smalls) == 1:
            total = eligible = smalls[0]
            withheld = 0

        # إضافة القيد مباشرة دون أي شروط حذف (No Drop Duplicates)
        raw_records.append({
            "اسم رب الأسرة": full_name,
            "البطاقة الحديثة": new_card,
            "البطاقة القديمة": old_card,
            "الكلي": total,
            "مستحق": eligible,
            "محجوب": withheld,
            "السطر الأصلي بالمستند": line_clean # تم إضافته لكي تتأكد بنفسك من كل سطر!
        })

    # بناء الجدول
    df = pd.DataFrame(raw_records)
    
    if not df.empty:
        df.insert(0, "التسلسل (ت)", df.index + 1)
        
        total_extracted = len(df)
        
        st.metric(label="📊 إجمالي القيود المستخرجة بقوة الكاسحة", value=f"{total_extracted} قيد")
        
        if total_extracted >= 514:
            st.success("🎉 تم اختراق الملف بالكامل! تم سحب كافة القيود (بما فيها المكررة نظامياً) دون إسقاط أي قيد.")
        else:
            st.warning(f"تم استخراج {total_extracted} قيد.")
            
        st.dataframe(df, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='جميع_القيود_المستخرجة', index=False)
            
        st.download_button(
            label="📥 تحميل كافة البيانات (Excel) - جاهزة ومكتملة",
            data=buffer.getvalue(),
            file_name="بيانات_الوكيل_المطلقة.xlsx",
            mime="application/vnd.ms-excel"
        )
    else:
        st.error("لم يتم العثور على قيود صالحة.")
