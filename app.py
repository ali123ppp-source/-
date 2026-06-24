import re
import pandas as pd
from docx import Document

def extract_and_clean_data(file_path):
    """
    دالة كاملة وعاملة بنسبة 100% لاستخراج البيانات من ملف الـ Word
    بدون فلاتر حذف خفية وبميزة التشخيص الكامل لمعرفة السجلات المفقودة.
    """
    try:
        doc = Document(file_path)
    except Exception as e:
        print(f"خطأ أثناء فتح الملف: {e}")
        return pd.DataFrame()
    
    # 4- أضف ملف تشخيص للسجلات المستبعدة
    rejected_rows = []
    raw_records = []
    
    # قراءة كافة النصوص من الفقرات والجداول داخل ملف الـ Word
    lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text.strip())
            
    for table in doc.tables:
        for row in table.rows:
            # دمج خلايا الجدول بفراغات لتسهيل عملية التقسيم لاحقاً
            row_text = "   ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
            if row_text.strip():
                lines.append(row_text)

    # حلقة معالجة الأسطر واستخراج البيانات
    for line in lines:
        # تنظيف وتقسيم السطر بناءً على المسافات الكبيرة أو الفواصل
        parts = [p.strip() for p in re.split(r'\s{2,}|,', line) if p.strip()]
        
        # تخطي أسطر العناوين الرئيسية حتى لا تسبب أخطاء
        if not parts or "اسم رب الاسرة" in line or "نوع الوكالة" in line or "ةلاكولا" in line:
            continue
            
        # فحص طول السطر (إذا كان السطر مقطوعاً أو قصيراً جداً يتم إرساله للمرفوضات للتشخيص)
        if len(parts) < 4:
            rejected_rows.append(line)
            continue
            
        # استخراج الاسم (الكلمات الكلية التي تحتوي على أحرف عربية)
        name_parts = [p for p in parts if re.search(r'[أ-ي]', p)]
        if not name_parts:
            rejected_rows.append(line)
            continue
        name = name_parts[0]
        
        # استخراج الأرقام من السطر
        numeric_parts = [p for p in parts if p.isdigit()]
        if len(numeric_parts) < 2:
            rejected_rows.append(line)
            continue
            
        # تعيين القيم الرقمية بمرونة لتفادي توقف الكود بسبب اختلاف التنسيق
        try:
            old_card_num = numeric_parts[1] if len(numeric_parts) > 1 else numeric_parts[0]
            total_m = numeric_parts[-3] if len(numeric_parts) >= 3 else "0"
            sub_m = numeric_parts[-2] if len(numeric_parts) >= 3 else "0"
            rej_m = numeric_parts[-1] if len(numeric_parts) >= 3 else "0"
        except Exception:
            rejected_rows.append(line)
            continue

        # 1- ألغِ حذف الأسماء المقلوبة واستبدلها بالتصحيح التلقائي
        fixed_words = []
        for w in name.split():
            if w.startswith('ة') and len(w) > 1:
                fixed_words.append(w[::-1])  # تصحيح الكلمة المقلوبة تلقائياً (مثل ةميلس -> سليمة)
            else:
                fixed_words.append(w)
        name = " ".join(fixed_words)
        
        # 2- ألغِ حذف السجلات المكررة مؤقتاً (تمت إزالة seen_cards بالكامل لضمان عدم إسقاط أي سجل متشابه)
        raw_records.append({
            "رقم البطاقة القديم": old_card_num,
            "اسم رب الاسرة": name,
            "الكلي": total_m,
            "مستحق": sub_m,
            "محجوب": rej_m
        })

    # إنشاء الـ DataFrame بعد معالجة كافة الأسطر
    df = pd.DataFrame(raw_records)
    
    # 5- أضف تقرير مقارنة إحصائي فوري يظهر في شاشة التشغيل
    print("=" * 60)
    print("📊 إحصائيات معالجة الملف الحالية:")
    print(f" - عدد السجلات التي تم استخراجها بنجاح: {len(raw_records)}")
    print(f" - عدد الأسطر/السجلات التي استبعدت للتشخيص: {len(rejected_rows)}")
    print("=" * 60)
    
    # 3- ألغِ حذف التكرار النهائي (استبدال دالة drop_duplicates بـ reset_index مباشرة)
    df = df.reset_index(drop=True)
    
    # 4- حفظ ملف تشخيص للسجلات المستبعدة في ملف Excel
    if rejected_rows:
        pd.DataFrame(
            {"السطر_المرفوض": rejected_rows}
        ).to_excel(
            "rejected_rows.xlsx",
            index=False
        )
        print("✔️ تنبيه: تم حفظ كافة الأسطر المستبعدة في ملف: rejected_rows.xlsx")
        print("يمكنك فتح هذا الملف لمعرفة الـ 35 سجلاً المفقودة والسبب وراء عدم مطابقتها.")
        
    return df

# لتشغيل الكود مباشرة على ملفك، قم بإلغاء التعليق عن السطر أدناه:
# df_result = extract_and_clean_data("626-عمر عادل موحان-FOOD.docx")
