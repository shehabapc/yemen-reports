import os
import streamlit as st
from datetime import date, datetime
import pandas as pd
from utils.data_loader import load_uploaded_file, extract_branch_name, parse_rep_data
from utils.processor import process_rep_data
from utils.excel_builder import build_excel_report
from utils.charts import create_summary_charts

# إعداد الصفحة
st.set_page_config(
    page_title="الشركة العربية للأدوية | نظام تقارير المناديب",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS (نفس الكود السابق مع إضافة تحسينات بسيطة)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
    html, body, [class*="css"], div, span, label {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
    }
    .stApp { background-color: #f8fafc; }
    .card-box {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
    }
    .section-title {
        color: #0d5c75;
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 8px;
        border-bottom: 2px solid #f1f5f9;
        padding-bottom: 8px;
    }
    div.stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #0d5c75 0%, #053b4c 100%) !important;
        color: white !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        padding: 12px 24px !important;
        border-radius: 10px !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(13, 92, 117, 0.25) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #d4a373 0%, #b88656 100%) !important;
        box-shadow: 0 6px 18px rgba(212, 163, 115, 0.35) !important;
        transform: translateY(-2px);
    }
    div.stDownloadButton > button {
        width: 100%;
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        padding: 12px 24px !important;
        border-radius: 10px !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.25) !important;
        transition: all 0.3s ease !important;
    }
    div.stDownloadButton > button:hover {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# الهيدر
logo_filename = "APC logo.png"
col_head_title, col_head_logo = st.columns([3, 1])
with col_head_title:
    st.markdown("""
    <div style="padding-top: 10px;">
        <h1 style="color: #0d5c75; font-weight: 800; font-size: 30px; margin: 0;">الشركة العربية للأدوية المحدودة</h1>
        <h3 style="color: #d4a373; font-weight: 700; font-size: 20px; margin-top: 4px; margin-bottom: 0;">Arab Pharmaceuticals Co. Ltd.</h3>
        <p style="color: #64748b; font-size: 14px; margin-top: 6px;">نظام توليد وتقارير حركة ومتابعة زيارات المناديب المنسقة تلقائياً</p>
    </div>
    """, unsafe_allow_html=True)
with col_head_logo:
    if os.path.exists(logo_filename):
        st.image(logo_filename, use_container_width=True)
    else:
        st.info("📊 APC Logo")
st.markdown("<hr style='margin-top: 10px; margin-bottom: 25px; border-color: #cbd5e1;'>", unsafe_allow_html=True)

# --- إعدادات التقرير ---
st.markdown('<div class="section-title">📅 إعدادات التقرير والملف الخام</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    start_date = st.date_input("تاريخ بداية التقرير", date(2026, 7, 21))
with col2:
    end_date = st.date_input("تاريخ نهاية التقرير", date(2026, 8, 20))
with col3:
    uploaded_file = st.file_uploader("رفع ملف الإكسل الخام (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])
    if uploaded_file:
        # عرض اسم الملف ونوعه
        st.caption(f"📄 {uploaded_file.name}")

# خيارات متقدمة (قابلة للتوسع)
with st.expander("⚙️ خيارات متقدمة"):
    # اختيار ورقة العمل (لـ Excel)
    sheet_name = None
    if uploaded_file and not uploaded_file.name.endswith('.csv'):
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = st.selectbox("اختر ورقة العمل", xls.sheet_names, index=0)
        except:
            pass
    # اختيار أيام العطلة
    weekend_days = st.multiselect(
        "أيام العطلة الأسبوعية",
        options={'الإثنين':0, 'الثلاثاء':1, 'الأربعاء':2, 'الخميس':3, 'الجمعة':4, 'السبت':5, 'الأحد':6},
        default=['الجمعة'],
        format_func=lambda x: x
    )
    weekend_indices = [weekend_days[day] for day in weekend_days]
    # تخصيص ألوان التقرير
    header_color = st.color_picker("لون خلفية رأس الجدول", "#D9E1F2")
    summary_color = st.color_picker("لون خلفية قسم الملخص", "#FCE4D6")

st.markdown("<br>", unsafe_allow_html=True)

# --- زر التشغيل ---
if st.button("🚀 إصدار التقرير النهائي") and uploaded_file is not None:
    with st.spinner("جاري معالجة البيانات وبناء ملف الإكسل المنسق..."):
        try:
            # 1. تحميل البيانات
            df_raw, sheet_used = load_uploaded_file(uploaded_file)
            branch_name = extract_branch_name(df_raw)
            # 2. استخراج بيانات المناديب
            rep_data = parse_rep_data(df_raw)
            if not rep_data:
                st.error("لم يتم العثور على بيانات للمناديب في الملف المرفق.")
                st.stop()
            # 3. معالجة البيانات
            processed, total_visits, total_absence = process_rep_data(
                rep_data, start_date, end_date, weekend_indices
            )
            # 4. بناء ملف الإكسل
            excel_output = build_excel_report(
                processed, branch_name, start_date, end_date,
                header_color=header_color, summary_color=summary_color
            )
            # 5. عرض النتائج
            st.markdown('<div class="section-title">📊 ملخص نتائج الفرع والإنتاجية</div>', unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("اسم الفرع", branch_name)
            m2.metric("إجمالي المناديب", len(processed))
            m3.metric("مجموع الزيارات", total_visits)
            m4.metric("مجموع أيام الغياب", f"{total_absence:.1f}")

            st.success(f"✅ تم الانتهاء بنجاح من معالجة بيانات فرع: {branch_name}")

            # تحميل الملف
            clean_branch = branch_name.replace(" ", "_")
            st.download_button(
                label="📥 تحميل التقرير النهائي المنسق (Excel)",
                data=excel_output,
                file_name=f"التقرير_النهائي_فرع_{clean_branch}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # عرض الرسوم البيانية
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-title">📈 رسوم بيانية تحليلية</div>', unsafe_allow_html=True)
            create_summary_charts(processed)

            # معاينة الجداول
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-title">🔍 معاينة جداول حركة المناديب</div>', unsafe_allow_html=True)
            tabs = st.tabs([f"👤 {rep}" for rep in processed.keys()])
            for tab, (rep_name, rep_info) in zip(tabs, processed.items()):
                with tab:
                    df_preview = pd.DataFrame(rep_info['rows'])
                    df_preview.columns = ['تسلسل', 'التاريخ', 'اليوم', 'زيارات صباحية', 'زيارات مسائية',
                                          'إجمالي الزيارات', 'بداية الصباحية', 'نهاية الصباحية',
                                          'بداية المسائية', 'نهاية المسائية', 'إجمالي الوقت', 'ملاحظات']
                    st.dataframe(df_preview, use_container_width=True)

        except Exception as e:
            st.error(f"⚠️ حدث خطأ أثناء معالجة الملف: {e}")
            # عرض تفاصيل الخطأ للمطور (في وضع التصحيح)
            # st.exception(e)  # يمكن تفعيله للتصحيح

elif uploaded_file is None:
    st.info("💡 يرجى رفع ملف الإكسل الخام لتمكين زر إصدار التقرير.")

# حفظ الحالة (اختياري) - يمكن إضافة session_state لتخزين البيانات المعالجة
