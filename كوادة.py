import streamlit as st
import pandas as pd
from io import BytesIO

# إعدادات واجهة المستخدم
st.set_page_config(page_title="نظام مطابقة كشوفات الوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #28B463; color: white; width: 100%; font-weight: bold; border-radius: 8px; font-size: 18px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الذكي لكشوفات الوكلاء 🔍⚖️</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>تعتمد هذه الأداة على مطابقة <b>رقم البطاقة التموينية القديم</b> حصراً لتجاهل أي أخطاء في الأسماء.</p>", unsafe_allow_html=True)
st.markdown("---")

# إنشاء عمودين لرفع الملفات
col1, col2 = st.columns(2)

with col1:
    st.markdown("<h3 style='text-align: right;'>📂 الملف الأول (الملف الأصلي/التفاصيل)</h3>", unsafe_allow_html=True)
    file1 = st.file_uploader("ارفع الملف الأول بصيغة Excel", type=['xlsx', 'xls'], key="file1")

with col2:
    st.markdown("<h3 style='text-align: right;'>📂 الملف الثاني (ملف الوكيل/المحدث)</h3>", unsafe_allow_html=True)
    file2 = st.file_uploader("ارفع الملف الثاني بصيغة Excel", type=['xlsx', 'xls'], key="file2")

st.markdown("<br>", unsafe_allow_html=True)

# 🧠 دالة البحث الذكي عن الأعمدة لتفادي أخطاء تغير الأسماء
def find_column(df, keywords):
    for kw in keywords:
        for col in df.columns:
            if kw in str(col):
                return col
    return df.columns[0] # خيار أخير في حال لم يجد الكلمة

if st.button("🚀 تشغيل محرك المقارنة واستخراج التقرير"):
    if file1 and file2:
        with st.spinner("جاري تحليل البيانات ومطابقة القيود..."):
            try:
                # 1. قراءة وتنظيف الملف الأول
                df1 = pd.read_excel(file1)
                
                # استخدام البحث الذكي
                col_name_1 = find_column(df1, ['اسم', 'الاسم'])
                col_card_1 = find_column(df1, ['قديم', 'بطاق'])
                col_count_1 = find_column(df1, ['سلة', 'افراد', 'عدد', 'مستحق'])

                df1.rename(columns={col_count_1: 'افراد_الملف_الأول', col_name_1: 'الاسم_الملف_الأول', col_card_1: 'رقم_البطاقة'}, inplace=True)
                df1['رقم_البطاقة'] = pd.to_numeric(df1['رقم_البطاقة'], errors='coerce')
                df1['افراد_الملف_الأول'] = pd.to_numeric(df1['افراد_الملف_الأول'].replace(['x', 'X', 'أ', 'محجوب'], 0), errors='coerce').fillna(0)
                df1 = df1.dropna(subset=['رقم_البطاقة'])

                # 2. قراءة وتنظيف الملف الثاني
                df2_test = pd.read_excel(file2, nrows=3)
                if df2_test.iloc[0].isna().sum() > len(df2_test.columns) / 2:
                    df2 = pd.read_excel(file2, header=1)
                else:
                    df2 = pd.read_excel(file2)

                # استخدام البحث الذكي للملف الثاني
                col_name_2 = find_column(df2, ['اسم', 'الاسم'])
                col_card_2 = find_column(df2, ['قديم', 'بطاق'])
                col_count_2 = find_column(df2, ['افراد', 'مستحق', 'عدد'])

                df2.rename(columns={col_count_2: 'افراد_الملف_الثاني', col_name_2: 'الاسم_الملف_الثاني', col_card_2: 'رقم_البطاقة'}, inplace=True)
                df2['رقم_البطاقة'] = pd.to_numeric(df2['رقم_البطاقة'], errors='coerce')
                df2['افراد_الملف_الثاني'] = pd.to_numeric(df2['افراد_الملف_الثاني'], errors='coerce').fillna(0)
                df2 = df2.dropna(subset=['رقم_البطاقة'])

                # 3. الدمج والمقارنة
                merged = pd.merge(df1, df2, on='رقم_البطاقة', how='outer', indicator=True)

                missing_in_file2 = merged[merged['_merge'] == 'left_only'][['رقم_البطاقة', 'الاسم_الملف_الأول', 'افراد_الملف_الأول']]
                missing_in_file1 = merged[merged['_merge'] == 'right_only'][['رقم_البطاقة', 'الاسم_الملف_الثاني', 'افراد_الملف_الثاني']]

                both_present = merged[merged['_merge'] == 'both']
                diff_counts = both_present[both_present['افراد_الملف_الأول'] != both_present['افراد_الملف_الثاني']][['رقم_البطاقة', 'الاسم_الملف_الأول', 'الاسم_الملف_الثاني', 'افراد_الملف_الأول', 'افراد_الملف_الثاني']]

                # 4. حفظ النتائج في الذاكرة (Buffer)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    missing_in_file2.to_excel(writer, sheet_name="مفقود في الملف الثاني", index=False)
                    missing_in_file1.to_excel(writer, sheet_name="مضاف جديد (مفقود بالأول)", index=False)
                    diff_counts.to_excel(writer, sheet_name="اختلاف عدد الأفراد", index=False)
                
                output.seek(0)
                
                # عرض النتائج كإحصائيات
                st.success("✅ تمت المقارنة بنجاح!")
                st.info(f"🔻 عدد القيود المفقودة في الملف الثاني: **{len(missing_in_file2)}**")
                st.info(f"🔺 عدد القيود الجديدة/المفقودة في الملف الأول: **{len(missing_in_file1)}**")
                st.warning(f"⚠️ عدد القيود التي اختلف فيها عدد الأفراد: **{len(diff_counts)}**")

                # زر تحميل التقرير
                st.download_button(
                    label="📥 تحميل تقرير المقارنة الشامل (Excel)",
                    data=output,
                    file_name="تقرير_المطابقة_النهائي.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة الملفات: {e}")
    else:
        st.warning("يرجى رفع الملفين أولاً قبل بدء المقارنة.")
