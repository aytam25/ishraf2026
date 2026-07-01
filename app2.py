import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io
import time
import urllib.parse

# ============================================================
# 1. إعدادات الصفحة المتقدمة
# ============================================================
st.set_page_config(
    page_title="نظام دمج تقييمات المعلمين - النسخة الاحترافية",
    page_icon="📊",
    layout="wide"
)
    
# ============================================================
# 2. حقن دلالات التنسيق الكامل لتأمين الواجهة العربي (RTL)
# ============================================================
st.markdown("""
    <style>
    /* تطبيق اتجاه القراءة والكتابة من اليمين إلى اليسار على كامل التطبيق */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        direction: rtl;
        text-align: right;
    }
    /* ضبط محاذاة التبويبات (Tabs) لتناسب الاتجاه العربي */
    div[data-testid="stTabs"] {
        direction: rtl;
    }
    div[data-testid="stTabs"] button {
        direction: rtl;
        text-align: right;
    }
    /* تحسين مظهر العناوين */
    .main-title { 
        font-size: 2.2rem !important; 
        font-weight: bold; 
        color: #1E3A8A; 
        text-align: right; 
    }
    .section-title {
        font-size: 1.5rem !important;
        font-weight: bold;
        color: #1E3A8A;
        margin-top: 15px;
        margin-bottom: 10px;
    }
    /* تنسيق كروت المؤشرات الحيوية (Metrics) */
    div[data-testid="stMetric"] { 
        background-color: #f8fafc; 
        padding: 10px 15px; 
        border-radius: 8px; 
        border: 1px solid #e2e8f0; 
        text-align: right;
    }
    /* محاذاة نصوص التنبيهات وصناديق رفع الملفات لليد اليمنى */
    .stAlert, div[data-testid="stFileUploader"] {
        direction: rtl;
        text-align: right;
    }
    /* ضبط حاويات الجداول الذكية داخلياً وعكس أشرطة التمرير للـ RTL */
    div[data-testid="stDataFrame"], div[data-testid="stDataFrame"] > div {
        direction: rtl !important;
        text-align: right !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📊 نظام دمج تقييمات المعلمين - المطابقة الذكية ورادار المخاطر</p>', unsafe_allow_html=True)
st.markdown("---")

# ============================================================
# 3. رفع الملفات
# ============================================================
col1, col2 = st.columns(2)
with col1:
    mushrif_file = st.file_uploader("📂 ارفع ملف المشرفين المجمّع", type=["xlsx"])
with col2:
    admin_file = st.file_uploader("📂 ارفع ملف الشؤون الإدارية (HR)", type=["xlsx"])

if mushrif_file and admin_file:
    
    with st.status("🔄 جاري معالجة البيانات وتطبيق خوارزميات رادار المخاطر...", expanded=True) as status:
        
        start_time = time.time()
        
        st.write("⏳ قراءة وتحليل ملفات Excel داخل الذاكرة...")
        mushrif_df = pd.read_excel(mushrif_file, dtype={'التقييم': str})
        admin_df = pd.read_excel(admin_file)
        
        # توحيد أسماء الأعمدة المتوقعة
        mushrif_df = mushrif_df.rename(columns={
            "الاسم": "اسم المعلم", "رقم الهوية": "رقم الهوية",
            "التقييم": "التقييم", "المشرف": "اسم المشرف"
        })
        admin_df = admin_df.rename(columns={
            "الاسم": "اسم المعلم", "رقم الهوية": "رقم الهوية"
        })
        
        # ------------------------------------------------------------
        # 🛡️ صمام الأمان: فحص بنية الملفات والتحقق من وجود الأعمدة المطلوبة
        # ------------------------------------------------------------
        required_mushrif = ["اسم المعلم", "رقم الهوية", "التقييم", "اسم المشرف"]
        required_admin = ["اسم المعلم", "رقم الهوية"]
        
        missing_mushrif = [col for col in required_mushrif if col not in mushrif_df.columns]
        missing_admin = [col for col in required_admin if col not in admin_df.columns]
        
        if missing_mushrif or missing_admin:
            status.update(label="❌ فشل التحقق من بنية الملفات المرفوعة!", state="error", expanded=True)
            
            if missing_mushrif:
                st.error(f"⚠️ **خطأ في ملف المشرفين:** لم نجد الأعمدة التالية أو مرادفات لها: {missing_mushrif}")
                st.info("💡 يرجى التأكد من أن الملف يحتوي على أعمدة واضحة للأسماء، أرقام الهويات، التقييمات، وأسماء المشرفين.")
                
            if missing_admin:
                st.error(f"⚠️ **خطأ في ملف الشؤون الإدارية (HR):** لم نجد الأعمدة التالية أو مرادفات لها: {missing_admin}")
                st.info("💡 يرجى التأكد من أن ملف الـ HR يحتوي على عمود للأسماء وعمود لأرقام الهويات.")
                
            st.stop()
            
        # ------------------------------------------------------------
        # 4. تنظيف وتطبيع البيانات
        # ------------------------------------------------------------
        st.write("🧼 تنظيف وتطبيع أرقام الهوية والأسماء برمجياً (Vectorized)...")
        for df in [mushrif_df, admin_df]:
            df["رقم الهوية_standard"] = df["رقم الهوية"].astype(str).str.strip().str.split('.').str[0]
            df["رقم الهوية_standard"] = df["رقم الهوية_standard"].replace(['nan', 'None', '<NA>'], '')
            
            # حشو الأرقام المكونة من 8 خانات بصفر جهة اليسار
            mask = (df["رقم الهوية_standard"].str.len() == 8) & (df["رقم الهوية_standard"].str.isdigit())
            df.loc[mask, "رقم الهوية_standard"] = '0' + df.loc[mask, "رقم الهوية_standard"]
            
            df["اسم المعلم"] = df["اسم المعلم"].fillna("").astype(str).str.strip()

        # ------------------------------------------------------------
        # 🛡️ دالة قياس نسبة التشابه الذكي 
        # ------------------------------------------------------------
        def similarity(a, b):
            a_str = str(a).strip() if not pd.isna(a) else ""
            b_str = str(b).strip() if not pd.isna(b) else ""
            
            if a_str in ["nan", "None", ""] or b_str in ["nan", "None", ""]:
                return 0
                
            if a_str == b_str: 
                return 1.0
                
            return SequenceMatcher(None, a_str, b_str).ratio()
        
        st.write("🏗️ بناء كشافات البحث السريع لملفات المشرفين...")
        supervisor_map = {}
        for row in mushrif_df.to_dict(orient='records'):
            id_num = row["رقم الهوية_standard"]
            if not id_num: continue
            
            if id_num not in supervisor_map:
                supervisor_map[id_num] = {
                    "التقييم": row["التقييم"],
                    "المشرف": [row["اسم المشرف"]],
                    "الاسم_المدخل": [row["اسم المعلم"]]
                }
            else:
                supervisor_map[id_num]["المشرف"].append(row["اسم المشرف"])
                supervisor_map[id_num]["الاسم_المدخل"].append(row["اسم المعلم"])

        st.write("🧠 تشغيل خوارزمية المطابقة الذكية وتحليل درجات الخطورة...")
        
        results, suggestions, not_found, error_details = [], [], [], []
        teacher_all_errors = []
        
        admin_ids_set = set(admin_df["رقم الهوية_standard"].unique())
        unmatched_supervisor_ids = [sid for sid in supervisor_map if sid not in admin_ids_set]
        
        for admin_row in admin_df.to_dict(orient='records'):
            admin_id = admin_row["رقم الهوية_standard"]
            admin_name = admin_row["اسم المعلم"]
            
            if admin_id in supervisor_map:
                data = supervisor_map[admin_id]
                score = data["التقييم"]
                
                statuses = []
                for sup, ent_name in zip(data["المشرف"], data["الاسم_المدخل"]):
                    if ent_name == admin_name:
                        statuses.append(f"✅ صحيح عند المشرف {sup}")
                    else:
                        statuses.append(f"⚠️ خطأ في الاسم عند المشرف {sup}")
                        name_sim = similarity(admin_name, ent_name)
                        
                        teacher_all_errors.append({
                            "المشرف المسؤول": sup,
                            "اسم المعلم (من الإدارة)": admin_name,
                            "الاسم المدخل (من المشرف)": ent_name,
                            "رقم هوية (HR)": admin_id,
                            "رقم هوية (المشرف)": admin_id,  # متطابقة برقم الهوية ولكن الاسم مختلف
                            "التقييم": score,
                            "📊 تشابه الاسم": f"{name_sim:.0%}",
                            "🚨 طبيعة الخطأ": "خطأ مطبعي في الاسم فقط",
                            "🔥 درجة الخطورة": "ℹ️ منخفضة (خطأ في الاسم فقط)"
                        })
                        
                        error_details.append({
                            "المشرف": sup, "المعلم": admin_name, "الرقم": admin_id,
                            "الاسم المدخل": ent_name, "الاسم الصحيح": admin_name, "نوع الخطأ": "خطأ في الاسم"
                        })
                
                results.append({
                    "المشرف المسؤول": " / ".join(data["المشرف"]),
                    "اسم المعلم (من الإدارة)": admin_name,
                    "الاسم المدخل (من المشرف)": " / ".join(data["الاسم_المدخل"]),
                    "رقم هوية (HR)": admin_id,
                    "رقم هوية (المشرف)": admin_id,
                    "التقييم": score,
                    "📌 الحالة": " | ".join(statuses)
                })
            
            else:
                found_suggestion = False
                best_match_id, best_match_name = None, None
                best_id_sim, best_name_sim = 0, 0
                
                for sup_id in unmatched_supervisor_ids:
                    id_sim = similarity(admin_id, sup_id)
                    if id_sim <= 0.6: continue
                    
                    for sup_name in supervisor_map[sup_id]["الاسم_المدخل"]:
                        name_sim = similarity(admin_name, sup_name)
                        
                        if name_sim > 0.85 and id_sim > 0.6:
                            if id_sim > best_id_sim:
                                best_id_sim, best_name_sim = id_sim, name_sim
                                best_match_id, best_match_name = sup_id, sup_name
                                found_suggestion = True
                
                if found_suggestion:
                    idx = supervisor_map[best_match_id]["الاسم_المدخل"].index(best_match_name)
                    sup_person = supervisor_map[best_match_id]["المشرف"][idx]
                    score = supervisor_map[best_match_id]["التقييم"]
                    
                    if best_name_sim == 1.0:
                        severity_level = "⚠️ متوسطة (خطأ في رقم الهوية)"
                        error_nature = "خطأ في رقم الهوية فقط"
                    else:
                        severity_level = "🚨 عالية جداً (خطأ في الاسم والهوية)"
                        error_nature = "خطأ مركب (الاسم والهوية معاً)"
                    
                    teacher_all_errors.append({
                        "المشرف المسؤول": sup_person,
                        "اسم المعلم (من الإدارة)": admin_name,
                        "الاسم المدخل (من المشرف)": best_match_name,
                        "رقم هوية (HR)": admin_id,
                        "رقم هوية (المشرف)": best_match_id, # هنا الهوية تختلف عن الـ HR
                        "التقييم": score,
                        "📊 تشابه الاسم": f"{best_name_sim:.0%}",
                        "🚨 طبيعة الخطأ": error_nature,
                        "🔥 درجة الخطورة": severity_level
                    })
                    
                    suggestions.append({
                        "المشرف المسؤول": sup_person,
                        "اسم المعلم (من الإدارة)": admin_name,
                        "الاسم المدخل (من المشرف)": best_match_name,
                        "رقم هوية (HR)": admin_id,
                        "رقم الهوية عند المشرف": best_match_id,
                        "التقييم": score
                    })
                else:
                    not_found.append({
                        "اسم المعلم (من الإدارة)": admin_name,
                        "رقم الهوية": admin_id,
                        "📌 الملاحظة": "لم يتم رصد أي تقييم أو اسم مشابه من قبل المشرفين"
                    })
        
        results_df = pd.DataFrame(results) if results else pd.DataFrame()
        suggestions_df = pd.DataFrame(suggestions) if suggestions else pd.DataFrame()
        not_found_df = pd.DataFrame(not_found) if not_found else pd.DataFrame()
        error_details_df = pd.DataFrame(error_details) if error_details else pd.DataFrame()
        teacher_all_errors_df = pd.DataFrame(teacher_all_errors) if teacher_all_errors else pd.DataFrame()
        
        if not error_details_df.empty:
            supervisor_error_summary = error_details_df.groupby("المشرف").size().reset_index()
            supervisor_error_summary.columns = ["المشرف", "عدد الأخطاء"]
        else:
            supervisor_error_summary = pd.DataFrame()
            
        end_time = time.time()
        execution_time = end_time - start_time
        
        status.update(label=f"⚡ تمت معالجة المخاطر والمطابقة المركبة بنجاح في {execution_time:.2f} ثانية!", state="complete", expanded=False)

    # دالة تلوين خلايا درجة الخطورة ديناميكياً
    def style_severity(val):
        if "عالية جداً" in str(val):
            return "background-color: #fee2e2; color: #991b1b; font-weight: bold; border: 1px solid #fca5a5;"
        elif "متوسطة" in str(val):
            return "background-color: #fef3c7; color: #92400e; font-weight: bold; border: 1px solid #fde68a;"
        elif "منخفضة" in str(val):
            return "background-color: #e0f2fe; color: #075985; border: 1px solid #bae6fd;"
        return ""

    # ============================================================
    # 5. عرض الإحصائيات الذكية والمؤشرات الحيوية
    # ============================================================
    st.markdown("### 📊 نظرة عامة على البيانات الكلية")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📋 إجمالي معلمين HR", len(admin_df))
    col2.metric("✅ تم دمجهم بنجاح", len(results_df))
    col3.metric("🔍 إجمالي المعلمين المخطأ بحقهم", len(teacher_all_errors_df))
    col4.metric("❌ معلمين بلا تقييم", len(not_found_df))
    col5.metric("⚠️ أخطاء المشرفين الكلية", len(error_details_df) + len(suggestions_df))
    
    st.markdown("---")

    # ============================================================
    # 6. تبويبات العرض المطورة (التقارير التفصيلية المجمعة)
    # ============================================================
    tab1, tab_errors, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 الموجودون والدمج", "🔍 رادار الأخطاء الشاملة للمعلم 🎯", 
        "🔍 مقترحات تصحيح الهوية", "❌ معلمين بلا تقييم", 
        "⚠️ رادار أخطاء المشرفين", "📈 التحليلات والمخططات", "📥 مركز تحميل التقارير"
    ])
    
    with tab1:
        st.subheader("📊 بيان المعلمين المدمجة تقييماتهم كاملاً")
        if not results_df.empty:
            # إعادة الترتيب هنا أيضاً للاتساق
            cols_order_base = ["المشرف المسؤول", "اسم المعلم (من الإدارة)", "الاسم المدخل (من المشرف)", "رقم هوية (HR)", "رقم هوية (المشرف)", "التقييم", "📌 الحالة"]
            results_df = results_df[[c for c in cols_order_base if c in results_df.columns]]
            col_order = list(results_df.columns)[::-1]
            col_config = {col: st.column_config.Column(alignment="right") for col in results_df.columns}
            st.dataframe(results_df, column_order=col_order, column_config=col_config, use_container_width=True, height=400, hide_index=True)
        else:
            st.info("لا توجد بيانات متاحة للعرض")
            
    with tab_errors:
        st.subheader("🔍 كشف مجمع وبؤرة تحليل أخطاء المعلمين")
        if not teacher_all_errors_df.empty:
            whatsapp_links = []
            for row in teacher_all_errors_df.to_dict(orient='records'):
                sup = row["المشرف المسؤول"]
                teacher = row["اسم المعلم (من الإدارة)"]
                err_type = row["🚨 طبيعة الخطأ"]
                msg = f"السلام عليكم أستاذ {sup}، يرجى التكرم بتعديل بيانات المعلم(ة) ({teacher}) في ملف التقييم الخاص بك، حيث تبين وجود ({err_type}). شكراً لتعاونك."
                encoded_msg = urllib.parse.quote(msg)
                whatsapp_links.append(f"https://wa.me/?text={encoded_msg}")
            
            teacher_all_errors_df["💬 تنبيه الواتساب"] = whatsapp_links
            
            cols_order_err = ["المشرف المسؤول", "اسم المعلم (من الإدارة)", "الاسم المدخل (من المشرف)", "رقم هوية (HR)", "رقم هوية (المشرف)", "التقييم", "📊 تشابه الاسم", "🚨 طبيعة الخطأ", "🔥 درجة الخطورة", "💬 تنبيه الواتساب"]
            teacher_all_errors_df = teacher_all_errors_df[[c for c in cols_order_err if c in teacher_all_errors_df.columns]]
            
            col_order = list(teacher_all_errors_df.columns)[::-1]
            col_config = {col: st.column_config.Column(alignment="right") for col in teacher_all_errors_df.columns}
            col_config["💬 تنبيه الواتساب"] = st.column_config.LinkColumn("💬 تنبيه الواتساب", display_text="📱 إرسال التنبيه للمشرف", alignment="right")
            
            styled_errors_df = teacher_all_errors_df.style.map(style_severity, subset=["🔥 درجة الخطورة"])
            st.dataframe(styled_errors_df, column_order=col_order, column_config=col_config, use_container_width=True, height=400, hide_index=True)
        else:
            st.success("✅ سجلات نظيفة تماماً! لا يوجد أي معلمين لديهم أخطاء في الأسماء أو الهويات.")
            
    with tab2:
        st.subheader("🔍 نظام المقترحات الذكي لتصحيح الهويات")
        if not suggestions_df.empty:
            col_order = list(suggestions_df.columns)[::-1]
            col_config = {col: st.column_config.Column(alignment="right") for col in suggestions_df.columns}
            st.dataframe(suggestions_df, column_order=col_order, column_config=col_config, use_container_width=True, hide_index=True)
        else:
            st.success("✅ نظيف! لا توجد فروقات أو أخطاء في كتابة الأرقام.")
            
    with tab3:
        st.subheader("❌ معلمين لم يرفع المشرفون تقييمات لهم")
        if not not_found_df.empty:
            col_order = list(not_found_df.columns)[::-1]
            col_config = {col: st.column_config.Column(alignment="right") for col in not_found_df.columns}
            st.dataframe(not_found_df, column_order=col_order, column_config=col_config, use_container_width=True, hide_index=True)
        else:
            st.success("✅ ممتاز! تم إدخال تقييمات لجميع المعلمين المقيدين في HR.")
            
    with tab4:
        st.subheader("⚠️ سجل أخطاء المشرفين بالتفصيل")
        if not error_details_df.empty:
            col_left, col_right = st.columns([2, 1])
            with col_left:
                st.markdown("#### 📝 جدول المخالفات والأخطاء المدخلة")
                col_order_det = list(error_details_df.columns)[::-1]
                col_config_det = {col: st.column_config.Column(alignment="right") for col in error_details_df.columns}
                st.dataframe(error_details_df, column_order=col_order_det, column_config=col_config_det, use_container_width=True, hide_index=True)
            with col_right:
                st.markdown("#### 📊 ترتيب المشرفين حسب عدد الأخطاء")
                col_order_sum = list(supervisor_error_summary.columns)[::-1]
                col_config_sum = {col: st.column_config.Column(alignment="right") for col in supervisor_error_summary.columns}
                st.dataframe(supervisor_error_summary, column_order=col_order_sum, column_config=col_config_sum, use_container_width=True, hide_index=True)
        else:
            st.success("✅ مذهل! جميع الأسماء المدخلة متطابقة.")
            
    with tab5:
        st.subheader("📈 رادار المخططات البيانية والتحليل الإحصائي")
        if not results_df.empty:
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                if not teacher_all_errors_df.empty:
                    st.markdown("##### 📌 توزيع مستويات خطورة البيانات المدخلة")
                    err_counts = teacher_all_errors_df['🔥 درجة الخطورة'].value_counts().reset_index()
                    err_counts.columns = ['درجة الخطورة', 'العدد']
                    st.bar_chart(err_counts.set_index('درجة الخطورة'))
            with col_chart2:
                st.markdown("##### 📊 الكثافة العددية للأخطاء لكل مشرف")
                if not supervisor_error_summary.empty:
                    st.bar_chart(supervisor_error_summary.set_index("المشرف"))
        else:
            st.info("قم برفع ملفات تحتوي على بيانات صالحة لتوليد المخططات.")
            
    with tab6:
        st.subheader("📥 مركز الاستخراج وتحميل التقارير المجمعة والمنفصلة")
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            if not results_df.empty: results_df.to_excel(writer, sheet_name="الموجودين_والمدمجين", index=False)
            if not teacher_all_errors_df.empty: teacher_all_errors_df.to_excel(writer, sheet_name="سجل_مخاطر_أخطاء_المعلمين", index=False)
            if not suggestions_df.empty: suggestions_df.to_excel(writer, sheet_name="اقتراحات_تعديل_الهويات", index=False)
            if not not_found_df.empty: not_found_df.to_excel(writer, sheet_name="معلمين_بدون_تقييم", index=False)
            if not error_details_df.empty: error_details_df.to_excel(writer, sheet_name="تفاصيل_أخطاء_المشرفين", index=False)
        output_buffer.seek(0)
        
        st.download_button(
            label="📥 تحميل التقرير الشامل والمدمج (ملف Excel واحد بتبويبات تشمل سجل المخاطر)",
            data=output_buffer,
            file_name="التقرير_الشامل_لدمج_المعلمين_ورادار_المخاطر.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ============================================================
    # 🎛️ [منصة التصفية بالأسفل] منصة التصفية التفاعلية الذكية المحدثة بالأعمدة الجديدة والواتساب المجمع
    # ============================================================
    st.markdown("---")
    st.markdown('<p class="section-title">⚙️ خيارات العرض والتصفية الحية (حسب المشرف والمشاكل)</p>', unsafe_allow_html=True)
    
    col_opt1, col_opt2 = st.columns([1, 2])
    with col_opt1:
        st.markdown("**خيارات العرض**")
        show_problems_only = st.checkbox("⚠️ عرض المعلمين الذين لديهم مشاكل فقط", value=False, key="filter_prob_key")
        
    with col_opt2:
        st.markdown("**تصفية حسب المشرف**")
        supervisor_options = ["الكل"]
        if not mushrif_df["اسم المشرف"].dropna().empty:
            supervisor_options.extend(mushrif_df["اسم المشرف"].unique().tolist())
            
        selected_sup = st.selectbox("اختر اسم المشرف:", options=supervisor_options, key="filter_sup_key")

    # 📱 [ميزة مضافة] ميزة التنبيه المجمع الذكي للمشرف المحدد
    if selected_sup != "الكل" and not teacher_all_errors_df.empty:
        sup_errors = teacher_all_errors_df[teacher_all_errors_df["المشرف المسؤول"] == selected_sup]
        if not sup_errors.empty:
            st.markdown(f"##### 📱 التنبيه المجمع والذكي للمشرف: ({selected_sup})")
            msg_lines = [f"السلام عليكم أستاذ {selected_sup}، يرجى التكرم بتعديل بيانات المعلمين التالية في ملف التقييم الخاص بك لوجود بعض الملاحظات:"]
            for idx, row in enumerate(sup_errors.to_dict(orient='records'), 1):
                msg_lines.append(f"{idx}- {row['اسم المعلم (من الإدارة)']} ⬅️ ({row['🚨 طبيعة الخطأ']})")
            msg_lines.append("شاكرين ومقدرين حسن تعاونكم وجهودكم المبذولة.")
            
            full_msg = "\n".join(msg_lines)
            encoded_msg = urllib.parse.quote(full_msg)
            whatsapp_url = f"https://wa.me/?text={encoded_msg}"
            
            st.link_button(f"📱 إرسال كشف الأخطاء المجمع عبر واتساب ({len(sup_errors)} ملاحظات)", whatsapp_url, use_container_width=True)
            st.markdown("---")

    # بناء الجدول المصفّح ديناميكياً
    if show_problems_only:
        filtered_df = teacher_all_errors_df.copy() if not teacher_all_errors_df.empty else pd.DataFrame()
        if selected_sup != "الكل" and not filtered_df.empty:
            filtered_df = filtered_df[filtered_df["المشرف المسؤول"] == selected_sup]
    else:
        filtered_df = results_df.copy() if not results_df.empty else pd.DataFrame()
        if selected_sup != "الكل" and not filtered_df.empty:
            filtered_df = filtered_df[filtered_df["المشرف المسؤول"].str.contains(selected_sup, na=False)]

    if not filtered_df.empty:
        # التراتبية الصارمة: المشرف -> اسم الإدارة -> اسم المشرف -> هوية HR -> هوية المشرف
        desired_cols = ["المشرف المسؤول", "اسم المعلم (من الإدارة)", "الاسم المدخل (من المشرف)", "رقم هوية (HR)", "رقم هوية (المشرف)", "التقييم"]
        existing_cols = [c for c in desired_cols if c in filtered_df.columns]
        remaining_cols = [c for c in filtered_df.columns if c not in existing_cols]
        filtered_df = filtered_df[existing_cols + remaining_cols]

        filtered_buffer = io.BytesIO()
        with pd.ExcelWriter(filtered_buffer, engine='openpyxl') as writer:
            export_df = filtered_df.copy()
            if "💬 تنبيه الواتساب" in export_df.columns:
                export_df = export_df.drop(columns=["💬 تنبيه الواتساب"])
            export_df.to_excel(writer, sheet_name="النتائج_المصفاة", index=False)
        filtered_buffer.seek(0)

        
        if show_problems_only:
            btn_label = f"📥 (Excel) تحميل ملف مشاكل المشرف: {selected_sup}"
            st.link_button(f"📱 إرسال كشف الأخطاء المجمع عبر واتساب ({len(sup_errors)} ملاحظات)", whatsapp_url, use_container_width=True)
            st.markdown("---")

        else:
            btn_label = f"📥 (Excel) تحميل بيان المشرف: {selected_sup}"
        if selected_sup == "الكل" and not show_problems_only:
            btn_label = "📥 (Excel) تحميل القائمة المصفاة المكتملة"
            
        st.download_button(
            label=btn_label,
            data=filtered_buffer,
            file_name=f"تقرير_{selected_sup}_{'مشاكل' if show_problems_only else 'عام'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_filtered_btn"
        )
        
        col_order_filt = list(filtered_df.columns)[::-1]
        col_config_filt = {col: st.column_config.Column(alignment="right") for col in filtered_df.columns}
        
        if "💬 تنبيه الواتساب" in filtered_df.columns:
            col_config_filt["💬 تنبيه الواتساب"] = st.column_config.LinkColumn("💬 تنبيه الواتساب", display_text="📱 إرسال التنبيه للمشرف", alignment="right")
            
        display_df = filtered_df.style.map(style_severity, subset=["🔥 درجة الخطورة"]) if "🔥 درجة الخطورة" in filtered_df.columns else filtered_df
        
        st.dataframe(
            display_df,
            column_order=col_order_filt,
            column_config=col_config_filt,
            use_container_width=True,
            height=300,
            hide_index=False
        )
    else:
        st.info("لا توجد بيانات مطابقة لخيارات التصفية المحددة حالياً.")

else:
    st.info("👈 يرجى رفع ملف المشرفين وملف الشؤون الإدارية (HR) من القائمة العلوية للبدء فورا.")

st.markdown("---")
st.caption("📌 نظام دمج تقييمات المعلمين الفائق - الإصدار الذكي v5.3")
