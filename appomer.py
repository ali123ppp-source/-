import re
import pandas as pd
from docx import Document

def extract_and_clean_data_v2(file_obj):
    doc = Document(file_obj)
    raw_records = []
    
    # 1. استخراج كل النصوص من الجداول والفقرات ودمجها في قائمة واحدة لتجاهل التمزق
    all_text = []
    for p in doc.paragraphs:
        if p.text.strip(): all_text.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip().replace('\n', ' ')
                if text and text not in all_text: # منع التكرار الأولي
                    all_text.append(text)
                    
    # دمج النص بالكامل للبحث الذكي
    full_text = " ".join(all_text)
    
    # 2. التخلص من القيود المكررة المتطابقة باستخدام الـ Regex
    # نمط يبحث عن (رقم بطاقة حديث) مسافة (رقم بطاقة قديم) مسافة (اسم عربي) مسافة (3 أرقام)
    pattern = re.compile(r'(\d{6,7})\s+(\d{5,7})\s+([أ-ي\s]+?)\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})')
    matches = pattern.findall(full_text)
    
    seen_cards = set()
    
    for match in matches:
        new_card, old_card, name, total, eligible, withheld = match
        
        # تنظيف الاسم من المسافات الزائدة
        clean_name = " ".join(name.split())
        
        # تخطي العناوين أو القيود المكررة
        if "اسم رب" in clean_name or old_card in seen_cards:
            continue
            
        seen_cards.add(old_card)
        
        raw_records.append({
            "اسم رب الأسرة": clean_name,
            "رقم البطاقة القديم": old_card,
            "الكلي": int(total),
            "محجوب": int(withheld),
            "مستحق": int(eligible)
        })
        
    df = pd.DataFrame(raw_records)
    
    # معالجة التسلسل الأصلي
    if not df.empty:
        df = df.reset_index(drop=True)
        df.insert(0, "ت", df.index + 1) # سيتم إعادة الترقيم بناءً على القيود السليمة المكتشفة
        
    return df
