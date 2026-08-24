import os
import streamlit as st
import pandas as pd
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from io import BytesIO

# --- 1. إعدادات الصفحة والواجهة ---
st.set_page_config(
    page_title="الشركة العربية للأدوية | نظام تقارير المناديب",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 2. التنسيقات البصرية الهيكلية (CSS) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');

    /* ضبط الخط والاتجاه العام */
    html, body, [class*="css"], div, span, label {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
    }

    /* خلفية التطبيق */
    .stApp {
        background-color: #f8fafc;
    }

    /* بطاقات الأقسام */
    .card-box {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
    }

    /* عناوين الأقسام */
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

    /* تخصيص زر التشغيل الرئيسي */
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

    /* زر التحميل الأخضر */
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

    /* تحسين طريقة عرض المقاييس Metrics */
    [data-testid="stMetricValue"] {
        font-size: 24px !important;
        font-weight: 800 !important;
        color: #0d5c75 !important;
    }

    [data-testid="stMetricBackground"] {
        background-color: #f8fafc !important;
        border-radius: 10px !important;
        padding: 12px !important;
        border: 1px solid #e2e8f0 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. الهيدر ورأس الصفحة مع شعار الشركة ---
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

# --- 4. مدخلات المستخدم والإعدادات ---
st.markdown('<div class="section-title">📅 إعدادات التقرير والملف الخام</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 1, 2])

with col1:
    start_date = st.date_input("تاريخ بداية التقرير", date(2026, 7, 21))

with col2:
    end_date = st.date_input("تاريخ نهاية التقرير", date(2026, 8, 20))

with col3:
    uploaded_file = st.file_uploader("رفع ملف الإكسل الخام (.xlsx)", type=["xlsx"])

st.markdown("<br>", unsafe_allow_html=True)

# --- 5. زر التشغيل ومعالجة البيانات ---
if st.button("🚀 إصدار التقرير النهائي") and uploaded_file is not None:
    with st.spinner("جاري معالجة البيانات وبناء ملف الإكسل المنسق..."):
        try:
            # قراءة الملف واستخراج اسم الفرع
            df_full = pd.read_excel(uploaded_file, sheet_name='Table 1', skiprows=5)
            branch_header = str(df_full.columns[0])
            branch_name = branch_header.split(':')[-1].strip() if "الفرع" in branch_header else "غير محدد"
            
            # استخراج بيانات المناديب
            rep_data = {}
            current_rep = None
            for idx, row in df_full.iterrows():
                rep_val = row['Unnamed: 10'] if 'Unnamed: 10' in df_full.columns else row.iloc[-1]
                if pd.notna(rep_val) and "المندوب" not in str(rep_val):
                    current_rep = str(rep_val).strip()
                    rep_data[current_rep] = []
                    continue
                date_col = 'Unnamed: 9' if 'Unnamed: 9' in df_full.columns else df_full.columns[-2]
                if pd.isna(row[date_col]) or str(row[date_col]).strip() == 'التاريخ':
                    continue
                if current_rep:
                    rep_data[current_rep].append({
                        'total_time': row.iloc[0], 'eve_end': row.iloc[1], 'eve_start': row.iloc[2],
                        'mor_end': row.iloc[3], 'mor_start': row.iloc[4], 'tot_reports': row.iloc[5],
                        'eve_reports': row.iloc[6], 'mor_reports': row.iloc[7],
                        'day_str': row.iloc[8], 'date': row[date_col]
                    })

            # وظائف تنسيق الوقت
            def clean_time(t_val):
                if pd.isna(t_val): return ""
                t_str = str(t_val).strip().split('.')[0]
                dt = pd.to_datetime(t_str, format='%H:%M:%S', errors='coerce')
                return t_str if pd.isna(dt) else dt.strftime('%I:%M:%S %p')

            def clean_total_time(t_val):
                if pd.isna(t_val): return ""
                t_str = str(t_val).strip().split('.')[0]
                try:
                    parts = t_str.split(':')
                    if len(parts) == 3: return f"{int(parts[0]):02d}:{int(parts[1]):02d}:{int(parts[2]):02d}"
                except: pass
                return t_str

            # بناء الجدول الزمني
            date_range = pd.date_range(pd.to_datetime(start_date), pd.to_datetime(end_date))
            days_ar = {0: 'الإثنين', 1: 'الثلاثاء', 2: 'الأربعاء', 3: 'الخميس', 4: 'الجمعة', 5: 'السبت', 6: 'الأحد'}
            processed_reps = {}

            total_all_visits = 0
            total_all_absence = 0

            for rep, rows in rep_data.items():
                df_rep = pd.DataFrame(rows)
                df_rep['date'] = pd.to_datetime(df_rep['date'], errors='coerce')
                df_rep = df_rep.dropna(subset=['date']).set_index('date')
                final_rows = []
                seq, total_working_days, absence_days, total_visits = 1, 0, 0, 0
                
                for d in date_range:
                    if d.weekday() == 4: continue  # استبعاد الجمعة
                    total_working_days += 1
                    day_name = days_ar[d.weekday()]
                    date_str = d.strftime('%Y-%m-%d')
                    
                    if d in df_rep.index:
                        row_df = df_rep.loc[d]
                        if isinstance(row_df, pd.DataFrame): row_df = row_df.iloc[0]
                        m_reps = row_df.get('mor_reports', 0) if pd.notna(row_df.get('mor_reports')) else 0
                        e_reps = row_df.get('eve_reports', 0) if pd.notna(row_df.get('eve_reports')) else 0
                        t_reps = row_df.get('tot_reports', 0) if pd.notna(row_df.get('tot_reports')) else 0
                        
                        m_start = clean_time(row_df.get('mor_start'))
                        m_end = clean_time(row_df.get('mor_end'))
                        e_start = clean_time(row_df.get('eve_start'))
                        e_end = clean_time(row_df.get('eve_end'))
                        
                        if m_reps == 0: m_start = 'لا يوجد زيارات'
                        if e_reps == 0: e_start = 'لا يوجد زيارات'
                        
                        total_visits += t_reps
                        final_rows.append({
                            'seq': seq, 'date': date_str, 'day': day_name,
                            'mor_reports': m_reps, 'eve_reports': e_reps, 'tot_reports': t_reps,
                            'mor_start': m_start, 'mor_end': m_end, 'eve_start': e_start, 'eve_end': e_end,
                            'total_time': clean_total_time(row_df.get('total_time')), 'notes': ''
                        })
                    else:
                        penalty = 0.5 if d.weekday() == 3 else 1.0
                        absence_days += penalty
                        final_rows.append({
                            'seq': seq, 'date': date_str, 'day': day_name,
                            'mor_reports': 0, 'eve_reports': 0, 'tot_reports': 0,
                            'mor_start': 'لا يوجد زيارات', 'mor_end': '', 'eve_start': 'لا يوجد زيارات', 'eve_end': '',
                            'total_time': '', 'notes': 'غياب'
                        })
                    seq += 1
                    
                net_working_days = total_working_days - absence_days
                average_visits = (total_visits / net_working_days) if net_working_days > 0 else 0
                processed_reps[rep] = {
                    'rows': final_rows,
                    'summary': {'total_visits': total_visits, 'absence_days': absence_days, 'net_working_days': net_working_days, 'average_visits': average_visits}
                }
                total_all_visits += total_visits
                total_all_absence += absence_days

            # إنشاء وتنسيق ملف الإكسل
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            header_fill_main = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            summary_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            bold_font = Font(name="Arial", size=11, bold=True, color="000000")
            regular_font = Font(name="Arial", size=10, color="000000")
            absent_font = Font(name="Arial", size=10, bold=True, color="FF0000")
            border = Border(top=Side(border_style="thin", color="000000"), left=Side(border_style="thin", color="000000"), right=Side(border_style="thin", color="000000"), bottom=Side(border_style="thin", color="000000"))
            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for rep, data in processed_reps.items():
                ws = wb.create_sheet(title=rep[:31])
                ws.sheet_view.rightToLeft = True
                
                ws.merge_cells('A2:M2')
                ws['A2'].value = f"مـــلخص الزيارات  اجمـــالي  بحسب الفترات\nفرع : {branch_name} | المندوب : {rep}    |   خلال الفترة من: {start_date.strftime('%Y/%m/%d')}  إلى: {end_date.strftime('%Y/%m/%d')}   |"
                ws['A2'].font, ws['A2'].alignment = bold_font, center_align
                ws.row_dimensions[2].height = 40
                
                headers_r4 = {'M': 'المندوب', 'L': 'تسلسل', 'K': 'التاريخ', 'J': 'اليوم', 'G': 'عدد التقارير', 'E': 'الفترة الصباحية', 'C': 'الفترة المسائية', 'B': 'إجمالي عدد ساعات\nودقائق العمل', 'A': 'ملاحظات'}
                for col, text in headers_r4.items():
                    cell = ws[f"{col}4"]
                    cell.value, cell.font, cell.fill, cell.alignment, cell.border = text, bold_font, header_fill_main, center_align, border
                for m in ['G4:I4', 'E4:F4', 'C4:D4', 'M4:M5', 'L4:L5', 'K4:K5', 'J4:J5', 'B4:B5', 'A4:A5']: ws.merge_cells(m)
                
                headers_r5 = {'I': 'الصباحية', 'H': 'المسائية', 'G': 'الاجمالي', 'F': 'وقت الجلسة الأولى', 'E': 'وقت الجلسة الأخيرة', 'D': 'وقت الجلسة الأولى', 'C': 'وقت الجلسة الأخيرة'}
                for col, text in headers_r5.items():
                    cell = ws[f"{col}5"]
                    cell.value, cell.font, cell.fill, cell.alignment, cell.border = text, bold_font, header_fill_main, center_align, border
                for r in ws['A4:M5']: 
                    for cell in r: cell.border = border
                ws.row_dimensions[4].height, ws.row_dimensions[5].height = 30, 30
                
                rows = data['rows']
                last_r = 5
                for i, row_data in enumerate(rows):
                    r = i + 6
                    last_r = r
                    data_map = {'M': rep if i == 0 else "", 'L': row_data['seq'], 'K': row_data['date'], 'J': row_data['day'], 'I': row_data['mor_reports'], 'H': row_data['eve_reports'], 'G': row_data['tot_reports'], 'F': row_data['mor_start'], 'E': row_data['mor_end'], 'D': row_data['eve_start'], 'C': row_data['eve_end'], 'B': row_data['total_time'], 'A': row_data['notes']}
                    for col, val in data_map.items():
                        cell = ws[f"{col}{r}"]
                        cell.value, cell.border, cell.alignment = val, border, center_align
                        cell.font = absent_font if val == 'غياب' else regular_font

                sum_start_r = last_r + 2
                summary = data['summary']
                sum_data = [
                    ("إجمالي عدد زيارات الأطباء:", summary['total_visits']),
                    ("إجمالي أيام الغياب (الخميس نصف يوم):", summary['absence_days']),
                    ("صافي أيام العمل:", summary['net_working_days']),
                    ("المتوسط اليومي للزيارات:", round(summary['average_visits'], 2))
                ]
                
                for i, (label, val) in enumerate(sum_data):
                    curr_r = sum_start_r + i
                    ws.merge_cells(f'K{curr_r}:M{curr_r}')
                    label_cell = ws[f'K{curr_r}']
                    val_cell = ws[f'J{curr_r}']
                    label_cell.value, val_cell.value = label, val
                    for cell in [label_cell, val_cell]:
                        cell.font, cell.alignment, cell.fill, cell.border = bold_font, center_align, summary_fill, border
                    ws.merge_cells(f'I{curr_r}:J{curr_r}')
                    ws[f'I{curr_r}'].border = border
                        
                widths = {'M': 25, 'L': 6, 'K': 12, 'J': 10, 'I': 10, 'H': 10, 'G': 10, 'F': 15, 'E': 15, 'D': 15, 'C': 15, 'B': 20, 'A': 15}
                for col, w in widths.items(): ws.column_dimensions[col].width = w

            # تحضير الملف للتحميل
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            # --- 6. عرض المؤشرات السريعة وزر التحميل ---
            st.markdown('<div class="section-title">📊 ملخص نتائج الفرع والإنتاجية</div>', unsafe_allow_html=True)
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("اسم الفرع", branch_name)
            m2.metric("إجمالي المناديب", f"{len(processed_reps)} مناديب")
            m3.metric("مجموع الزيارات", f"{total_all_visits} زيارة")
            m4.metric("مجموع أيام الغياب", f"{total_all_absence} يوم")

            st.success(f"✅ تم الانتهاء بنجاح من معالجة بيانات فرع: {branch_name}")

            clean_branch_name = branch_name.replace(" ", "_")
            st.download_button(
                label="📥 تحميل التقرير النهائي المنسق (Excel)",
                data=output,
                file_name=f"التقرير_النهائي_فرع_{clean_branch_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # معاينة التقرير داخل التبويبات
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-title">🔍 معاينة جداول حركة المناديب</div>', unsafe_allow_html=True)
            tabs = st.tabs([f"👤 {rep}" for rep in processed_reps.keys()])
            for tab, (rep_name, rep_info) in zip(tabs, processed_reps.items()):
                with tab:
                    df_preview = pd.DataFrame(rep_info['rows'])
                    df_preview.columns = ['تسلسل', 'التاريخ', 'اليوم', 'زيارات صباحية', 'زيارات مسائية', 'إجمالي الزيارات', 'بداية الصباحية', 'نهاية الصباحية', 'بداية المسائية', 'نهاية المسائية', 'إجمالي الوقت', 'ملاحظات']
                    st.dataframe(df_preview, use_container_width=True)

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
elif uploaded_file is None:
    st.info("💡 يرجى رفع ملف الإكسل الخام لتمكين زر إصدار التقرير.")
