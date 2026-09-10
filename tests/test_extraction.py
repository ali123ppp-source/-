"""
اختبارات تراجعية لمحرك استخراج البيانات (قراءة الملفات وتحديد مكان الأعمدة).

الهدف: حماية دائمة ضد رجوع أخطاء تحديد الأعمدة (مثل انزلاق رقم البطاقة إلى
عمود "الكلي") التي اكتُشفت وتم تشخيصها على ملف حقيقي. لا تعتمد على أي مكتبة
اختبار خارجية (لا pytest) — تشغيل مباشر:

    python3 tests/test_extraction.py
"""
import os
import sys
import types

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)


def _install_streamlit_stub():
    st_stub = types.ModuleType("streamlit")

    class _Spinner:
        def __enter__(self): return self
        def __exit__(self, *a): return False

    class _Ctx:
        def __enter__(self): return self
        def __exit__(self, *a): return False

    class _SessionState(dict):
        def __getattr__(self, name):
            try:
                return self[name]
            except KeyError:
                raise AttributeError(name)
        def __setattr__(self, name, value):
            self[name] = value

    def _noop(*a, **k):
        return None

    st_stub.session_state = _SessionState()
    for fn in ["markdown", "file_uploader", "radio", "button", "success", "error",
               "warning", "download_button", "dataframe", "write", "set_page_config",
               "checkbox", "text_input", "selectbox", "info", "stop", "title",
               "header", "subheader"]:
        setattr(st_stub, fn, _noop)
    st_stub.columns = lambda n, **k: [_Ctx() for _ in range(n if isinstance(n, int) else len(n))]
    st_stub.expander = lambda *a, **k: _Ctx()
    st_stub.spinner = lambda *a, **k: _Spinner()
    sys.modules["streamlit"] = st_stub


_install_streamlit_stub()

import importlib.util
_spec = importlib.util.spec_from_file_location("app", os.path.join(REPO_ROOT, "app.py"))
app = importlib.util.module_from_spec(_spec)
sys.modules["app"] = app
_spec.loader.exec_module(app)


PASSED = []
FAILED = []


def check(name, condition, detail=""):
    if condition:
        PASSED.append(name)
    else:
        FAILED.append((name, detail))


# ---------------------------------------------------------------------------
# 1) البنية الحقيقية المشخَّصة هذه الجلسة: عمود فاضي بين الاسم والكلي،
#    وعمود البطاقة الجديدة باسم "البطاقة الجديدة" (لا يحتوي كلمة "حديث").
# ---------------------------------------------------------------------------
def test_real_world_layout_with_blank_column():
    rows_data = [
        ["ت", "رقم البطاقة", "البطاقة الجديدة", "اسم العائلة", "", "كلي", "مستحق", "محجوب"],
        ["1", "223615", "4401375", "أحمد كريم عبار ال عبد ويس", "", "5", "5", "0"],
        ["2", "245630", "3997044", "أحمد مهدى صالح العجيلي", "", "6", "6", "0"],
    ]
    records = app._extract_records_by_headers(rows_data, "رقم البطاقة الحديث", "الاسم الثلاثي فقط")
    check("real_world_layout: records extracted", records is not None and len(records) == 2)
    if records:
        r = records[0]
        check("real_world_layout: card is new-card value (not leaked into كلي)",
              r["رقم البطاقة"] == "4401375", detail=str(r))
        check("real_world_layout: الكلي is the small total, not the card number",
              r["الكلي"] == 5, detail=str(r))
        check("real_world_layout: مستحق correct", r["مستحق"] == 5, detail=str(r))
        check("real_world_layout: محجوب correct", r["محجوب"] == 0, detail=str(r))


# ---------------------------------------------------------------------------
# 2) مرادف جديد لعمود "الكلي": "مجموع" بدل "كلي"/"اجمالي".
# ---------------------------------------------------------------------------
def test_total_column_synonym_majmoo():
    rows_data = [
        ["ت", "اسم رب الأسرة", "رقم البطاقة", "مجموع الافراد", "مستحق", "محجوب"],
        ["1", "زينب عباس كاظم", "1234567", "4", "4", "0"],
    ]
    records = app._extract_records_by_headers(rows_data, "رقم البطاقة القديم", "الاسم الثلاثي فقط")
    check("majmoo_synonym: header detected via 'مجموع'", records is not None and len(records) == 1,
          detail=str(records))
    if records:
        check("majmoo_synonym: الكلي correct", records[0]["الكلي"] == 4, detail=str(records[0]))


# ---------------------------------------------------------------------------
# 2ب) فخ: عمود بعنوان "مجموع المستحقين" يجب أن يُطابَق كـ"مستحق"، لا كـ"الكلي"
#     (كلمة "مجموع" لا يجب أن تطغى على "مستحق"/"محجوب" لو ظهرت بالخلية نفسها).
# ---------------------------------------------------------------------------
def test_majmoo_does_not_shadow_mustahiq_or_mahjoob():
    rows_data = [
        ["ت", "اسم رب الأسرة", "رقم البطاقة", "مجموع الافراد", "مجموع المستحقين", "محجوب"],
        ["1", "زينب عباس كاظم", "1234567", "6", "4", "2"],
    ]
    header_idx, idx_map = app._locate_header_row(rows_data)
    check("majmoo_shadow: header found despite two 'مجموع' cells", header_idx == 0, detail=str(idx_map))
    check("majmoo_shadow: 'مجموع المستحقين' mapped to مستحق, not الكلي",
          idx_map["مستحق"] == 4, detail=str(idx_map))
    check("majmoo_shadow: 'مجموع الافراد' (no مستحق/محجوب) still mapped to الكلي",
          idx_map["كلي"] == 3, detail=str(idx_map))
    records = app._extract_records_by_headers(rows_data, "رقم البطاقة القديم", "الاسم الثلاثي فقط")
    check("majmoo_shadow: record extracted with correct values",
          records is not None and records[0]["الكلي"] == 6 and records[0]["مستحق"] == 4,
          detail=str(records))


# ---------------------------------------------------------------------------
# 3) عمود "الرقم التسلسلي" يجب ألا يُعامَل كعمود بطاقة (يحتوي "رقم" كنص فرعي).
# ---------------------------------------------------------------------------
def test_sequence_number_column_not_treated_as_card():
    rows_data = [
        ["الرقم التسلسلي", "اسم رب الأسرة", "رقم البطاقة", "الكلي", "مستحق", "محجوب"],
        ["1", "خالد حسن فرج", "9988776", "3", "3", "0"],
    ]
    header_idx, idx_map = app._locate_header_row(rows_data)
    check("sequence_column: header found", header_idx == 0, detail=str(idx_map))
    check("sequence_column: 'الرقم التسلسلي' not mistaken for a card column",
          idx_map["بطاقة_قديم"] != 0, detail=str(idx_map))
    check("sequence_column: real card column detected at index 2",
          idx_map["بطاقة_قديم"] == 2, detail=str(idx_map))


# ---------------------------------------------------------------------------
# 4) التشكيل (diacritics) بعنوان العمود لا يجب أن يكسر المطابقة.
# ---------------------------------------------------------------------------
def test_diacritics_do_not_break_matching():
    rows_data = [
        ["ت", "اِسْمُ رَبِّ الأُسْرَةِ", "رقم البطاقة", "كُلِّي", "مُسْتَحَقّ", "مَحْجُوب"],
        ["1", "علي حسين جبار", "5551234", "5", "5", "0"],
    ]
    header_idx, idx_map = app._locate_header_row(rows_data)
    check("diacritics: header row found despite tashkeel", header_idx == 0, detail=str(idx_map))
    required = ["اسم", "كلي", "مستحق", "محجوب"]
    check("diacritics: all required fields mapped", all(idx_map[k] != -1 for k in required),
          detail=str(idx_map))


# ---------------------------------------------------------------------------
# 5) مدقق البيانات (_validate_extracted_records) يلتقط رقم بطاقة انزلق لعمود الكلي.
# ---------------------------------------------------------------------------
def test_validator_flags_leaked_card_number_in_total():
    bad_records = [
        {"اسم رب الأسرة": "فلان الفلاني", "رقم البطاقة": "4401375", "الكلي": 4401375, "مستحق": 5, "محجوب": 0},
        {"اسم رب الأسرة": "فلان الثاني", "رقم البطاقة": "3997044", "الكلي": 3997044, "مستحق": 6, "محجوب": 0},
    ]
    warnings = app._validate_extracted_records(bad_records)
    check("validator: flags huge الكلي values", len(warnings) >= 1, detail=str(warnings))
    check("validator: warning mentions الكلي", any("الكلي" in w for w in warnings), detail=str(warnings))


def test_validator_flags_mustahiq_exceeding_total():
    bad_records = [
        {"اسم رب الأسرة": "فلان الفلاني", "رقم البطاقة": "1234567", "الكلي": 3, "مستحق": 9, "محجوب": 0},
    ]
    warnings = app._validate_extracted_records(bad_records)
    check("validator: flags مستحق > الكلي", len(warnings) >= 1, detail=str(warnings))


def test_validator_silent_on_clean_data():
    good_records = [
        {"اسم رب الأسرة": "فلان الفلاني", "رقم البطاقة": "4401375", "الكلي": 5, "مستحق": 5, "محجوب": 0},
        {"اسم رب الأسرة": "فلان الثاني", "رقم البطاقة": "3997044", "الكلي": 6, "مستحق": 6, "محجوب": 0},
    ]
    warnings = app._validate_extracted_records(good_records)
    check("validator: no warnings on clean data", warnings == [], detail=str(warnings))


# ---------------------------------------------------------------------------
# 6) عدم وجود صف ترويسة مطابق يجب أن يُرجع None بدل استخراج بيانات خاطئة.
# ---------------------------------------------------------------------------
def test_no_matching_header_returns_none():
    rows_data = [
        ["عنوان تقرير عشوائي لا صلة له بالبيانات"],
        ["نص آخر غير مرتبط بالجدول المطلوب"],
    ]
    records = app._extract_records_by_headers(rows_data, "رقم البطاقة القديم", "الاسم الثلاثي فقط")
    check("no_header: returns None when nothing matches", records is None)


# ---------------------------------------------------------------------------
# 7) extract_and_clean_data الآن يرجع (df, warnings) كـ tuple.
# ---------------------------------------------------------------------------
def test_extract_and_clean_data_returns_tuple():
    class FakeUpload:
        def __init__(self, rows):
            import io
            from docx import Document
            doc = Document()
            table = doc.add_table(rows=1, cols=len(rows[0]))
            for i, text in enumerate(rows[0]):
                table.rows[0].cells[i].text = text
            for row in rows[1:]:
                cells = table.add_row().cells
                for i, text in enumerate(row):
                    cells[i].text = text
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            self._buf = buf
            self.name = "test.docx"
        def read(self, *a, **k):
            return self._buf.read(*a, **k)
        def seek(self, *a, **k):
            return self._buf.seek(*a, **k)
        def __getattr__(self, item):
            return getattr(self._buf, item)

    rows = [
        ["ت", "اسم رب الأسرة", "رقم البطاقة", "الكلي", "مستحق", "محجوب"],
        ["1", "زينب عباس كاظم", "1234567", "4", "4", "0"],
    ]
    upload = FakeUpload(rows)
    result = app.extract_and_clean_data(upload, "رقم البطاقة القديم", "الاسم الثلاثي فقط", True)
    check("extract_and_clean_data: returns a 2-tuple", isinstance(result, tuple) and len(result) == 2,
          detail=str(type(result)))
    df, warnings = result
    check("extract_and_clean_data: df has one row", len(df) == 1, detail=str(df))
    check("extract_and_clean_data: warnings is a list", isinstance(warnings, list))


def main():
    tests = [
        test_real_world_layout_with_blank_column,
        test_total_column_synonym_majmoo,
        test_majmoo_does_not_shadow_mustahiq_or_mahjoob,
        test_sequence_number_column_not_treated_as_card,
        test_diacritics_do_not_break_matching,
        test_validator_flags_leaked_card_number_in_total,
        test_validator_flags_mustahiq_exceeding_total,
        test_validator_silent_on_clean_data,
        test_no_matching_header_returns_none,
        test_extract_and_clean_data_returns_tuple,
    ]
    for t in tests:
        try:
            t()
        except Exception as e:
            FAILED.append((t.__name__, f"EXCEPTION: {e!r}"))

    print(f"PASSED: {len(PASSED)}")
    for name in PASSED:
        print(f"  ok  {name}")
    if FAILED:
        print(f"FAILED: {len(FAILED)}")
        for name, detail in FAILED:
            print(f"  FAIL  {name}  -- {detail}")
        sys.exit(1)
    print("ALL TESTS PASSED")


if __name__ == "__main__":
    main()
