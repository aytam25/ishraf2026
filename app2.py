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
    /* تحسين مظهر العنوان الرئيسي */
    .main-title { 
        font-size: 2.2rem !important; 
        font-weight: bold; 
        color: #1E3A8A; 
        text-align: right; 
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
                
            st.stop() # إيقاف البرنامج هنا بأمان وحماية المستخدم من الأخطاء البرمجية الهيكلية
            
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
        # 🛡️ دالة قياس نسبة التشابه الذكي (المحصنة ضد القيم الغائبة والخلايا الفارغة)
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
        teacher_all_errors = [] # مصفوفة تجمع كافة أخطاء المعلمين لتغذية تبويب رادار المعلم
        
        admin_ids_set = set(admin_df["رقم الهوية_standard"].unique())
        unmatched_supervisor_ids = [sid for sid in supervisor_map if sid not in admin_ids_set]
        
        for admin_row in admin_df.to_dict(orient='records'):
            admin_id = admin_row["رقم الهوية_standard"]
            admin_name = admin_row["اسم المعلم"]
            
            # ------------------------------------------------------------
            # الحالة 1: الرقم متطابق تماماً في السيستم
            # ------------------------------------------------------------
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
                        
                        # إضافة خطأ (منخفض الخطورة) لأن الرقم صحيح والاسم به مشكلة مطبعية
                        teacher_all_errors.append({
                            "👤 اسم المعلم (HR)": admin_name,
                            "🆔 الهوية الصحيحة (HR)": admin_id,
                            "📝 الاسم عند المشرف": ent_name,
                            "🔍 الهوية عند المشرف": admin_id,
                            "👨‍🏫 المشرف المسؤول": sup,
                            "📊 تشابه الاسم": f"{name_sim:.0%}",
                            "📊 تشابه الهوية": "100%",
                            "🚨 طبيعة الخطأ": "خطأ مطبعي في الاسم فقط",
                            "🔥 درجة الخطورة": "ℹ️ منخفضة (خطأ في الاسم فقط)"
                        })
                        
                        error_details.append({
                            "المشرف": sup, "المعلم": admin_name, "الرقم": admin_id,
                            "الاسم المدخل": ent_name, "الاسم الصحيح": admin_name, "نوع الخطأ": "خطأ في الاسم"
                        })
                
                results.append({
                    "👤 الاسم (HR)": admin_name,
                    "🆔 الهوية": admin_id,
                    "⭐ التقييم": score,
                    "👨‍🏫 المشرف": " / ".join(data["المشرف"]),
                    "📝 الاسم المدخل": " / ".join(data["الاسم_المدخل"]),
                    "📌 الحالة": " | ".join(statuses)
                })
            
            # ------------------------------------------------------------
            # الحالة 2: الرقم غير موجود (البحث الذكي والمركب عن الأخطاء)
            # ------------------------------------------------------------
            else:
                found_suggestion = False
                best_match_id, best_match_name = None, None
                best_id_sim, best_name_sim = 0, 0
                
                for sup_id in unmatched_supervisor_ids:
                    id_sim = similarity(admin_id, sup_id)
                    if id_sim <= 0.6: continue
                    
                    for sup_name in supervisor_map[sup_id]["الاسم_المدخل"]:
                        name_sim = similarity(admin_name, sup_name)
                        
                        # التقاط التطابق الذكي (سواء خطأ في الهوية فقط أو خطأ مركب اسم+هوية)
                        if name_sim > 0.85 and id_sim > 0.6:
                            if id_sim > best_id_sim:
                                best_id_sim, best_name_sim = id_sim, name_sim
                                best_match_id, best_match_name = sup_id, sup_name
                                found_suggestion = True
                
                if found_suggestion:
                    # تحديد درجة الخطورة بناءً على نوع الخطأ المدخل
                    idx = supervisor_map[best_match_id]["الاسم_المدخل"].index(best_match_name)
                    sup_person = supervisor_map[best_match_id]["المشرف"][idx]
                    
                    if best_name_sim == 1.0:
                        severity_level = "⚠️ متوسطة (خطأ في رقم الهوية)"
                        error_nature = "خطأ في رقم الهوية فقط"
                    else:
                        severity_level = "🚨 عالية جداً (خطأ في الاسم والهوية)"
                        error_nature = "خطأ مركب (الاسم والهوية معاً)"
                    
                    # تسجيل المخالفة بالتفصيل في جدول رادار الأخطاء الشامل
                    teacher_all_errors.append({
                        "👤 اسم المعلم (HR)": admin_name,
                        "🆔 الهوية الصحيحة (HR)": admin_id,
                        "📝 الاسم عند المشرف": best_match_name,
                        "🔍 الهوية عند المشرف": best_match_id,
                        "👨‍🏫 المشرف المسؤول": sup_person,
                        "📊 تشابه الاسم": f"{best_name_sim:.0%}",
                        "📊 تشابه الهوية": f"{best_id_sim:.0%}",
                        "🚨 طبيعة الخطأ": error_nature,
                        "🔥 درجة الخطورة": severity_level
                    })
                    
                    suggestions.append({
                        "👤 الاسم في (HR)": admin_name,
                        "🆔 الرقم الصحيح (HR)": admin_id,
                        "🔍 الرقم الخاطئ عند المشرف": best_match_id,
                        "📝 الاسم المدخل بواسطة المشرف": best_match_name,
                        "👨‍🏫 المشرف المسؤول": sup_person,
                        "📊 تشابه الرقم": f"{best_id_sim:.0%}",
                        "📊 تشابه الاسم": f"{best_name_sim:.0%}"
                    })
                else:
                    not_found.append({
                        "👤 الاسم (HR)": admin_name,
                        "🆔 الرقم (HR)": admin_id,
                        "📌 الملاحظة": "لم يتم رصد أي تقييم أو اسم مشابه من قبل المشرفين"
                    })
        
        # تحويل المخرجات لـ DataFrames
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

    # ============================================================
    # 5. عرض الإحصائيات الذكية والمؤشرات الحيوية
    # ============================================================
    st.markdown("### 📊 نظرة عامة على البيانات")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📋 إجمالي معلمين HR", len(admin_df))
    col2.metric("✅ تم دمجهم بنجاح", len(results_df))
    col3.metric("🔍 إجمالي المعلمين المخطأ بحقهم", len(teacher_all_errors_df))
    col4.metric("❌ معلمين بلا تقييم", len(not_found_df))
    col5.metric("⚠️ أخطاء المشرفين الكلية", len(error_details_df) + len(suggestions_df))
    
    st.markdown("---")
    
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
    # 6. تبويبات العرض المطورة ومحاذاتها بدقة بالغة
    # ============================================================
    tab1, tab_errors, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 الموجودون والدمج", "🔍 رادار الأخطاء الشاملة للمعلم 🎯", 
        "🔍 مقترحات تصحيح الهوية", "❌ معلمين بلا تقييم", 
        "⚠️ رادار أخطاء المشرفين", "📈 التحليلات والمخططات", "📥 مركز تحميل التقارير"
    ])
    
    with tab1:
        st.subheader("📊 بيان المعلمين المدمجة تقييماتهم")
        if not results_df.empty:
            st.dataframe(results_df, use_container_width=True, height=400)
        else:
            st.info("لا توجد بيانات متاحة للعرض")
            
    with tab_errors:
        st.subheader("🔍 كشف مجمع وبؤرة تحليل أخطاء المعلمين")
        st.info("💡 هذا التبويب يجمع لك كل معلم واجه مشكلة في البيانات المدخلة، ويحتوي على زر لإرسال تنبيه جاهز للمشرف عبر الواتساب.")
        
        if not teacher_all_errors_df.empty:
            # توليد روابط الواتساب الذكية تلقائياً
            whatsapp_links = []
            for row in teacher_all_errors_df.to_dict(orient='records'):
                sup = row["👨‍🏫 المشرف المسؤول"]
                teacher = row["👤 اسم المعلم (HR)"]
                err_type = row["🚨 طبيعة الخطأ"]
                
                msg = f"السلام عليكم أستاذ {sup}، يرجى التكرم بتعديل بيانات المعلم(ة) ({teacher}) في ملف التقييم الخاص بك، حيث تبين وجود ({err_type}). شكراً لتعاونك."
                encoded_msg = urllib.parse.quote(msg)
                whatsapp_links.append(f"https://wa.me/?text={encoded_msg}")
            
            teacher_all_errors_df["💬 تنبيه الواتساب"] = whatsapp_links
            styled_errors_df = teacher_all_errors_df.style.map(style_severity, subset=["🔥 درجة الخطورة"])
            
            st.dataframe(
                styled_errors_df, 
                use_container_width=True, 
                height=400,
                column_config={
                    "💬 تنبيه الواتساب": st.column_config.LinkColumn("💬 تنبيه الواتساب", display_text="📱 إرسال التنبيه للمشرف")
                }
            )
            
            c1, c2, c3 = st.columns(3)
            high_count = len(teacher_all_errors_df[teacher_all_errors_df["🔥 درجة الخطورة"].str.contains("عالية")])
            med_count = len(teacher_all_errors_df[teacher_all_errors_df["🔥 درجة الخطورة"].str.contains("متوسطة")])
            low_count = len(teacher_all_errors_df[teacher_all_errors_df["🔥 درجة الخطورة"].str.contains("منخفضة")])
            
            c1.metric("🚨 أخطاء مركبة (عالية جداً)", high_count)
            c2.metric("⚠️ أخطاء هوية (متوسطة)", med_count)
            c3.metric("ℹ️ أخطاء أسماء (منخفضة)", low_count)
        else:
            st.success("✅ سجلات نظيفة تماماً! لا يوجد أي معلمين لديهم أخطاء في الأسماء أو الهويات.")
            
    with tab2:
        st.subheader("🔍 نظام المقترحات الذكي لتصحيح الهويات")
        if not suggestions_df.empty:
            st.dataframe(suggestions_df, use_container_width=True)
        else:
            st.success("✅ نظيف! لا توجد فروقات أو أخطاء في كتابة الأرقام.")
            
    with tab3:
        st.subheader("❌ معلمين لم يرفع المشرفون تقييمات لهم")
        if not not_found_df.empty:
            st.dataframe(not_found_df, use_container_width=True)
        else:
            st.success("✅ ممتاز! تم إدخال تقييمات لجميع المعلمين المقيدين في HR.")
            
    with tab4:
        st.subheader("⚠️ سجل أخطاء المشرفين بالتفصيل")
        if not error_details_df.empty:
            col_left, col_right = st.columns([2, 1])
            with col_left:
                st.markdown("#### 📝 جدول المخالفات والأخطاء المدخلة")
                st.dataframe(error_details_df, use_container_width=True)
            with col_right:
                st.markdown("#### 📊 ترتيب المشرفين حسب عدد الأخطاء")
                st.dataframe(supervisor_error_summary, use_container_width=True)
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

else:
    st.info("👈 يرجى رفع ملف المشرفين وملف الشؤون الإدارية (HR) من القائمة العلوية للبدء فورا.")

st.markdown("---")
st.caption("📌 نظام دمج تقييمات المعلمين الفائق - الإصدار الذكي v5.0 (مدمج برادار المخاطر والأخطاء المركبة والتلوين الديناميكي)")
