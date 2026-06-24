import re
import pandas as pd
from docx import Document

def extract_and_clean_data(file_path):
    doc = Document(file_path)
    
    # 4- أضف ملف تشخيص للسجلات المستبعدة في بداية الدالة
    rejected_rows = []
    raw_records = []
    
    # 2- تم إلغاء حذف السجلات المكررة مؤقتاً (حذف seen_cards = set() بالكامل من هنا)
    
    # قراءة الأسطر من مستند Word (الفقرات والجداول)
    lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text.strip())
            
    for table in doc.tables:
        for row in table.rows:
            row_text = " , ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
            if row_text.strip():
                lines.append(row_text)

    # حلقة معالجة الأسطر واستخراج البيانات
    for line in lines:
        
        # [ملاحظة]: هذا الجزء يعتمد على نمط الـ Regex الخاص بك لاستخراج البيانات.
        # سنقوم هنا بتطبيق الفلاتر والتعديلات المطلوبة بدقة على المتغيرات المستخرجة.
        
        # مثال على فحص نمط الاسم (الرجاء مطابقتها مع كودك الأصلي):
        # name_match = re.search(r'([آ-ي\s]+)', line)
        
        # 4- استبدال الاستبعاد المباشر بتسجيل السطر المرفوض
        if not name_match:
            rejected_rows.append(line)
            continue
            
        # فرضاً تم استخراج المتغيرات مثل: name, old_card_num, counts، إلخ.
        # name = name_match.group(1).strip()
        
        # 1- ألغِ حذف الأسماء المقلوبة واستبدلها بالتصحيح والتعديل التلقائي
        fixed_words = []
        for w in name.split():
            if w.startswith('ة') and len(w) > 1:
                fixed_words.append(w[::-1])  # عكس الكلمة المقلوبة لإصلاحها تلقائياً
            else:
                fixed_words.append(w)
        name = " ".join(fixed_words)
        
        # 2- ألغِ حذف السجلات المكررة مؤقتاً 
        # (تمت إزالة كود التحقق من card_id inside seen_cards تماماً لضمان عدم إسقاط أي سجل متشابه)
        
        # في فلاتر التحقق من الأسطر المقطوعة مثل len(counts) أو غيرها:
        # استبدل الـ continue بـ إرسال السطر إلى مصفوفة المرفوضات:
        # if len(counts) < 2:
        #     rejected_rows.append(line)
        #     continue
            
        # إذا تم استخراج البيانات بنجاح نقوم بإضافتها
        # raw_records.append({
        #     "اسم رب الاسرة": name,
        #     "رقم البطاقة القديم": old_card_num,
        #     ...
        # })
        pass

    # إنشاء الـ DataFrame بعد معالجة كافة الأسطر
    df = pd.DataFrame(raw_records)
    
    # 5- أضف تقرير مقارنة فورياً بعد إنشاء الـ DataFrame
    print("=" * 60)
    print("عدد السجلات المستخرجة بنجاح:", len(raw_records))
    print("عدد السجلات المرفوضة تشخيصياً:", len(rejected_rows))
    print("=" * 60)
    
    # 3- ألغِ حذف التكرار النهائي (استبدال دالة drop_duplicates بـ reset_index مباشرة)
    df = df.reset_index(drop=True)
    
    # 4- حفظ ملف تشخيص للسجلات المستبعدة في ملف Excel قبل عمل return
    if rejected_rows:
        pd.DataFrame(
            {"السطر_المرفوض": rejected_rows}
        ).to_excel(
            "rejected_rows.xlsx",
            index=False
        )
        print("تنبيه: تم حفظ السجلات المستبعدة تشخيصياً في ملف: rejected_rows.xlsx")
        
    return df
