import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io

# ============================================================
# إعدادات الصفحة - RTL
# ============================================================
st.set_page_config(
    page_title="نظام دمج تقييمات المعلمين - الذكي",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS لتوجيه RTL وتحسين المظهر
# ============================================================
st.markdown("""
<style>
    /* توجيه الصفحة بالكامل لـ RTL */
    .main > div {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* توجيه الجداول */
    .stDataFrame {
        direction: rtl !important;
    }
    .stDataFrame table {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* توجيه العناوين */
    h1, h2, h3, h4, h5, h6 {
        text-align: right !important;
        direction: rtl !important;
    }
    
    /* توجيه التبويبات */
    .stTabs [data-baseweb="tab-list"] {
        direction: rtl !important;
        gap: 2px !important;
    }
    .stTabs [data-baseweb="tab"] {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* توجيه الأعمدة */
    .stColumns {
        direction: rtl !important;
    }
    
    /* توجيه المقاييس (metrics) */
    .stMetric {
        direction: rtl !important;
        text-align: right !important;
    }
    .stMetric label {
        text-align: right !important;
    }
    
    /* توجيه المعلومات (info, warning, success) */
    .stAlert {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* توجيه الأزرار */
    .stButton button {
        direction: rtl !important;
    }
    
    /* توجيه مربعات الاختيار والقوائم */
    .stSelectbox, .stMultiSelect {
        direction: rtl !important;
    }
    
    /* توجيه التوسيع (expander) */
    .streamlit-expanderHeader {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* توجيه التحميلات */
    .stDownloadButton {
        direction: rtl !important;
    }
    
    /* تعديل عرض الجداول */
    .dataframe {
        direction: rtl !important;
    }
    .dataframe th, .dataframe td {
        text-align: right !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 نظام دمج تقييمات المعلمين - المطابقة الذكية")
st.markdown("---")

# ============================================================
# رفع الملفات
# ============================================================
col1, col2 = st.columns(2)
with col1:
    mushrif_file = st.file_uploader("📂 ارفع ملف المشرفين المجمّع", type=["xlsx"])
with col2:
    admin_file = st.file_uploader("📂 ارفع ملف الشؤون الإدارية (HR)", type=["xlsx"])

if mushrif_file and admin_file:
    
    with st.spinner("🔄 جاري قراءة وتحليل الملفات..."):
        
        # قراءة الملفات
        mushrif_df = pd.read_excel(mushrif_file, dtype={'التقييم': str})
        admin_df = pd.read_excel(admin_file)
        
        # توحيد أسماء الأعمدة
        mushrif_df = mushrif_df.rename(columns={
            "الاسم": "اسم المعلم",
            "رقم الهوية": "رقم الهوية",
            "التقييم": "التقييم",
            "المشرف": "اسم المشرف"
        })
        admin_df = admin_df.rename(columns={
            "الاسم": "اسم المعلم",
            "رقم الهوية": "رقم الهوية"
        })
        
        # ============================================================
        # دوال المعالجة
        # ============================================================
        
        def normalize_id(id_value):
            if pd.isna(id_value):
                return ""
            id_str = str(id_value).strip()
            if '.' in id_str:
                id_str = id_str.split('.')[0]
            return id_str
        
        def pad_to_9(id_str):
            if not id_str:
                return id_str
            id_str = str(id_str).strip()
            if '.' in id_str:
                id_str = id_str.split('.')[0]
            if len(id_str) == 8 and id_str.isdigit():
                return '0' + id_str
            return id_str
        
        def clean_text(text):
            if pd.isna(text):
                return ""
            return str(text).strip()
        
        def similarity(a, b):
            if not a or not b:
                return 0
            return SequenceMatcher(None, a, b).ratio()
        
        # تطبيق التطبيع
        mushrif_df["رقم الهوية"] = mushrif_df["رقم الهوية"].apply(normalize_id)
        mushrif_df["اسم المعلم"] = mushrif_df["اسم المعلم"].apply(clean_text)
        mushrif_df["رقم الهوية_standard"] = mushrif_df["رقم الهوية"].apply(pad_to_9)
        
        admin_df["رقم الهوية"] = admin_df["رقم الهوية"].apply(normalize_id)
        admin_df["اسم المعلم"] = admin_df["اسم المعلم"].apply(clean_text)
        admin_df["رقم الهوية_standard"] = admin_df["رقم الهوية"].apply(pad_to_9)
        
        # ============================================================
        # بناء قاموس HR (المرجع)
        # ============================================================
        hr_map = {}
        for _, row in admin_df.iterrows():
            id_num = row["رقم الهوية_standard"]
            name = row["اسم المعلم"]
            if id_num and name:
                hr_map[id_num] = name
        
        # ============================================================
        # بناء قاموس المشرفين
        # ============================================================
        supervisor_map = {}
        for _, row in mushrif_df.iterrows():
            id_num = row["رقم الهوية_standard"]
            if id_num:
                if id_num not in supervisor_map:
                    supervisor_map[id_num] = {
                        "التقييم": row["التقييم"],
                        "المشرف": row["اسم المشرف"],
                        "الاسم_المدخل": row["اسم المعلم"]
                    }
                else:
                    existing = supervisor_map[id_num]
                    if isinstance(existing["المشرف"], list):
                        existing["المشرف"].append(row["اسم المشرف"])
                        existing["الاسم_المدخل"].append(row["اسم المعلم"])
                    else:
                        supervisor_map[id_num] = {
                            "التقييم": existing["التقييم"],
                            "المشرف": [existing["المشرف"], row["اسم المشرف"]],
                            "الاسم_المدخل": [existing["الاسم_المدخل"], row["اسم المعلم"]]
                        }
        
        # ============================================================
        # المطابقة الذكية
        # ============================================================
        
        results = []
        not_found = []
        suggestions = []
        error_details = []
        
        for _, admin_row in admin_df.iterrows():
            admin_id = admin_row["رقم الهوية_standard"]
            admin_name = admin_row["اسم المعلم"]
            
            if admin_id in supervisor_map:
                data = supervisor_map[admin_id]
                score = data["التقييم"]
                supervisor = data["المشرف"]
                entered_name = data["الاسم_المدخل"]
                
                if isinstance(entered_name, list):
                    statuses = []
                    for sup, ent_name in zip(supervisor, entered_name):
                        if ent_name == admin_name:
                            statuses.append(f"✅ صحيح عند المشرف {sup}")
                        else:
                            statuses.append(f"⚠️ خطأ في الاسم عند المشرف {sup}")
                            error_details.append({
                                "المشرف": sup,
                                "المعلم": admin_name,
                                "الرقم": admin_id,
                                "الاسم المدخل": ent_name,
                                "الاسم الصحيح": admin_name,
                                "نوع الخطأ": "خطأ في الاسم"
                            })
                    status = " | ".join(statuses)
                    display_supervisor = " / ".join(supervisor)
                    display_name = " / ".join(entered_name)
                else:
                    if entered_name == admin_name:
                        status = "✅ صحيح"
                    else:
                        status = f"⚠️ خطأ في الاسم عند المشرف {supervisor}"
                        error_details.append({
                            "المشرف": supervisor,
                            "المعلم": admin_name,
                            "الرقم": admin_id,
                            "الاسم المدخل": entered_name,
                            "الاسم الصحيح": admin_name,
                            "نوع الخطأ": "خطأ في الاسم"
                        })
                    display_supervisor = supervisor
                    display_name = entered_name
                
                results.append({
                    "👤 الاسم (HR)": admin_name,
                    "🆔 الهوية": admin_id,
                    "⭐ التقييم": score,
                    "👨‍🏫 المشرف": display_supervisor,
                    "📝 الاسم المدخل": display_name,
                    "📌 الحالة": status
                })
            
            else:
                found_suggestion = False
                best_match_id = None
                best_match_name = None
                best_id_sim = 0
                best_name_sim = 0
                
                for hr_id, hr_name in hr_map.items():
                    if hr_id != admin_id:
                        id_sim = similarity(admin_id, hr_id)
                        name_sim = similarity(admin_name, hr_name)
                        
                        if name_sim > 0.85 and id_sim > 0.6:
                            if id_sim > best_id_sim:
                                best_id_sim = id_sim
                                best_name_sim = name_sim
                                best_match_id = hr_id
                                best_match_name = hr_name
                                found_suggestion = True
                
                if found_suggestion:
                    suggestions.append({
                        "👤 الاسم (المدخل)": admin_name,
                        "🆔 الرقم (المدخل)": admin_id,
                        "🔍 الرقم المقترح": best_match_id,
                        "✅ الاسم الصحيح في HR": best_match_name,
                        "📊 تشابه الرقم": f"{best_id_sim:.0%}",
                        "📊 تشابه الاسم": f"{best_name_sim:.0%}"
                    })
                else:
                    not_found.append({
                        "👤 الاسم (المدخل)": admin_name,
                        "🆔 الرقم (المدخل)": admin_id,
                        "📌 الملاحظة": "لم يتم العثور على هذا المعلم في HR ولا يوجد اسم مشابه"
                    })
        
        # ============================================================
        # إنشاء DataFrames
        # ============================================================
        results_df = pd.DataFrame(results) if results else pd.DataFrame()
        suggestions_df = pd.DataFrame(suggestions) if suggestions else pd.DataFrame()
        not_found_df = pd.DataFrame(not_found) if not_found else pd.DataFrame()
        error_details_df = pd.DataFrame(error_details) if error_details else pd.DataFrame()
        
        if not error_details_df.empty:
            supervisor_error_summary = error_details_df.groupby("المشرف").size().reset_index()
            supervisor_error_summary.columns = ["المشرف", "عدد الأخطاء"]
        else:
            supervisor_error_summary = pd.DataFrame()
        
        # ============================================================
        # عرض النتائج
        # ============================================================
        
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("📊 إجمالي المعلمين في HR", len(admin_df))
        col2.metric("✅ موجودون", len(results_df))
        col3.metric("🔍 اقتراحات تصحيح", len(suggestions_df))
        col4.metric("❌ غير موجودين", len(not_found_df))
        col5.metric("⚠️ أخطاء في الاسم", len(error_details_df))
        
        st.markdown("---")
        
        # ============================================================
        # تبويبات العرض (RTL)
        # ============================================================
        
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "📊 الموجودون",
            "🔍 اقتراحات التصحيح",
            "❌ غير الموجودين",
            "⚠️ أخطاء المشرفين",
            "📈 المخططات",
            "📥 تحميل التقارير"
        ])
        
        with tab1:
            st.subheader("📊 المعلمين الموجودين في الإدارة")
            if not results_df.empty:
                st.dataframe(results_df, use_container_width=True, height=400)
                
                col1, col2, col3 = st.columns(3)
                col1.metric("✅ صحيح", len(results_df[results_df['📌 الحالة'] == "✅ صحيح"]))
                col2.metric("⚠️ خطأ في الاسم", len(results_df[results_df['📌 الحالة'].str.contains("خطأ", na=False)]))
                col3.metric("⭐ متوسط التقييم", 
                           f"{pd.to_numeric(results_df['⭐ التقييم'], errors='coerce').mean():.1f}" 
                           if not results_df.empty else "0")
            else:
                st.info("لا توجد بيانات")
        
        with tab2:
            st.subheader("🔍 اقتراحات تصحيح رقم الهوية")
            st.info("📌 هذه الحالات قد يكون فيها خطأ إملائي في رقم واحد أو اثنين، والاسم متطابق تقريباً")
            if not suggestions_df.empty:
                st.dataframe(suggestions_df, use_container_width=True)
                st.warning(f"⚠️ {len(suggestions_df)} حالة تحتاج إلى مراجعة وتصحيح الرقم")
            else:
                st.success("✅ لا توجد اقتراحات للتصحيح")
        
        with tab3:
            st.subheader("❌ المعلمين غير الموجودين في الإدارة")
            if not not_found_df.empty:
                st.warning(f"⚠️ {len(not_found_df)} معلم غير موجود في قاعدة البيانات")
                st.dataframe(not_found_df, use_container_width=True)
                st.info("📌 هذه المعلمين يحتاجون إلى إضافتهم في نظام الإدارة")
            else:
                st.success("✅ جميع المعلمين موجودون في قاعدة البيانات")
        
        with tab4:
            st.subheader("⚠️ تفاصيل أخطاء المشرفين")
            if not error_details_df.empty:
                st.dataframe(error_details_df, use_container_width=True)
                
                st.subheader("📊 ملخص الأخطاء لكل مشرف")
                st.dataframe(supervisor_error_summary, use_container_width=True)
                
                if not supervisor_error_summary.empty:
                    st.subheader("📊 رسم بياني لأخطاء المشرفين")
                    st.bar_chart(supervisor_error_summary.set_index("المشرف"))
            else:
                st.success("✅ لا توجد أخطاء من المشرفين")
        
        with tab5:
            st.subheader("📈 المخططات البيانية")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if not results_df.empty:
                    status_counts = results_df['📌 الحالة'].value_counts().reset_index()
                    status_counts.columns = ['الحالة', 'العدد']
                    st.subheader("📊 توزيع حالات الموجودين")
                    st.dataframe(status_counts, use_container_width=True)
                    st.bar_chart(status_counts.set_index('الحالة'))
            
            with col2:
                if not results_df.empty:
                    numeric_scores = pd.to_numeric(results_df['⭐ التقييم'], errors='coerce')
                    if not numeric_scores.dropna().empty:
                        st.subheader("📊 توزيع التقييمات")
                        score_dist = numeric_scores.value_counts().sort_index().reset_index()
                        score_dist.columns = ['التقييم', 'العدد']
                        st.dataframe(score_dist.head(10), use_container_width=True)
                        st.bar_chart(score_dist.set_index('التقييم').head(10))
            
            st.subheader("📊 إحصائيات عامة")
            stats_data = {
                "الفئة": ["الموجودون", "اقتراحات تصحيح", "غير موجودين", "أخطاء في الاسم"],
                "العدد": [len(results_df), len(suggestions_df), len(not_found_df), len(error_details_df)]
            }
            stats_df = pd.DataFrame(stats_data)
            st.dataframe(stats_df, use_container_width=True)
            st.bar_chart(stats_df.set_index("الفئة"))
        
        with tab6:
            st.subheader("📥 تحميل التقارير")
            
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                if not results_df.empty:
                    results_df.to_excel(writer, sheet_name="الموجودين", index=False)
                if not suggestions_df.empty:
                    suggestions_df.to_excel(writer, sheet_name="اقتراحات_تصحيح", index=False)
                if not not_found_df.empty:
                    not_found_df.to_excel(writer, sheet_name="غير_الموجودين", index=False)
                if not error_details_df.empty:
                    error_details_df.to_excel(writer, sheet_name="أخطاء_المشرفين", index=False)
                if not supervisor_error_summary.empty:
                    supervisor_error_summary.to_excel(writer, sheet_name="ملخص_أخطاء_المشرفين", index=False)
            output_buffer.seek(0)
            
            st.download_button(
                label="📥 تحميل التقرير الكامل (Excel)",
                data=output_buffer,
                file_name="تقرير_المعلمين_الكامل.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.subheader("📥 تحميل ملفات منفصلة")
            
            if not results_df.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    results_df.to_excel(writer, sheet_name="الموجودين", index=False)
                buffer.seek(0)
                st.download_button(
                    label="📥 تحميل الموجودين",
                    data=buffer,
                    file_name="الموجودين.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            if not suggestions_df.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    suggestions_df.to_excel(writer, sheet_name="اقتراحات_تصحيح", index=False)
                buffer.seek(0)
                st.download_button(
                    label="📥 تحميل اقتراحات التصحيح",
                    data=buffer,
                    file_name="اقتراحات_تصحيح.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            if not not_found_df.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    not_found_df.to_excel(writer, sheet_name="غير_الموجودين", index=False)
                buffer.seek(0)
                st.download_button(
                    label="📥 تحميل غير الموجودين",
                    data=buffer,
                    file_name="غير_الموجودين.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            if not error_details_df.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    error_details_df.to_excel(writer, sheet_name="أخطاء_المشرفين", index=False)
                buffer.seek(0)
                st.download_button(
                    label="📥 تحميل أخطاء المشرفين",
                    data=buffer,
                    file_name="أخطاء_المشرفين.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        # ============================================================
        # خيارات التصفية
        # ============================================================
        
        with st.expander("🎯 خيارات التصفية والبحث"):
            col1, col2 = st.columns(2)
            
            with col1:
                if not results_df.empty:
                    if st.button("عرض المعلمين الذين لديهم مشاكل فقط"):
                        filtered = results_df[results_df['📌 الحالة'].str.contains("خطأ", na=False)].copy()
                        if filtered.empty:
                            st.info("✅ لا توجد مشاكل")
                        else:
                            st.dataframe(filtered, use_container_width=True)
            
            with col2:
                supervisors_list = mushrif_df["اسم المشرف"].dropna().unique().tolist()
                if supervisors_list:
                    selected_supervisor = st.selectbox("اختر اسم المشرف:", supervisors_list)
                    if selected_supervisor:
                        filtered_by_supervisor = results_df[results_df['👨‍🏫 المشرف'].str.contains(selected_supervisor, na=False)]
                        if filtered_by_supervisor.empty:
                            st.info(f"لا توجد بيانات للمدرسين تحت إشراف {selected_supervisor}")
                        else:
                            st.dataframe(filtered_by_supervisor, use_container_width=True)

else:
    st.info("👈 يرجى رفع ملف المشرفين وملف الشؤون الإدارية للبدء")
    st.markdown("""
    ### 📋 تنسيق الملفات المطلوب:
    
    **ملف المشرفين:**
    - يحتوي على أعمدة: `الاسم`, `رقم الهوية`, `التقييم`, `المشرف`
    
    **ملف الشؤون الإدارية (HR):**
    - يحتوي على أعمدة: `الاسم`, `رقم الهوية`
    """)

st.markdown("---")
st.caption("📌 نظام دمج تقييمات المعلمين - الإصدار الذكي v3.0 (RTL)")
