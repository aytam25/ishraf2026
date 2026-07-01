import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io
import time

# ============================================================
# إعدادات الصفحة المتقدمة
# ============================================================
st.set_page_config(
    page_title="نظام دمج تقييمات المعلمين - الذكي الفائق",
    page_icon="📊",
    layout="wide"
)

# تحسين المظهر العام باستخدام CSS مخصص بسيط ومتوافق
st.markdown("""
    <style>
    .main-title { font-size: 2.2rem !important; font-weight: bold; color: #1E3A8A; text-align: right; }
    .metric-card { background-color: #F3F4F6; padding: 15px; border-radius: 10px; border-right: 5px solid #3B82F6; }
    div[data-testid="stMetric"] { background-color: #f8fafc; padding: 10px 15px; border-radius: 8px; border: 1px solid #e2e8f0; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📊 نظام دمج تقييمات المعلمين - المطابقة الذكية والسرعة الفائقة</p>', unsafe_allow_html=True)
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
    
    # استخدام st.status لتتبع خطوات المعالجة بدقة
    with st.status("🔄 جاري معالجة البيانات وتطبيق خوارزميات المطابقة...", expanded=True) as status:
        
        start_time = time.time()
        
        st.write("⏳ قراءة وتحليل ملفات Excel داخل الذاكرة...")
        # قراءة الملفات مباشرة مع تحسين الأنواع
        mushrif_df = pd.read_excel(mushrif_file, dtype={'التقييم': str})
        admin_df = pd.read_excel(admin_file)
        
        # توحيد أسماء الأعمدة لضمان عدم حدوث أخطاء كشافات
        mushrif_df = mushrif_df.rename(columns={
            "الاسم": "اسم المعلم", "رقم الهوية": "رقم الهوية",
            "التقييم": "التقييم", "المشرف": "اسم المشرف"
        })
        admin_df = admin_df.rename(columns={
            "الاسم": "اسم المعلم", "رقم الهوية": "رقم الهوية"
        })
        
        st.write("🧼 تنظيف وتطبيع أرقام الهوية والأسماء برمجياً (Vectorized)...")
        # استخدام هندسة المصفوفات الآمنة في Pandas لمنع الـ TypeError نهائياً
        for df in [mushrif_df, admin_df]:
            # تنظيف وتوحيد رقم الهوية وعزل الكسور إن وجدت دفعة واحدة
            df["رقم الهوية_standard"] = df["رقم الهوية"].astype(str).str.strip().str.split('.').str[0]
            df["رقم الهوية_standard"] = df["رقم الهوية_standard"].replace(['nan', 'None', '<NA>'], '')
            
            # الحل الجذري هنا: حشو الأرقام المكونة من 8 خانات بصفر باستخدام قناع موجه (Vectorized Mask) آمن تماماً
            mask = (df["رقم الهوية_standard"].str.len() == 8) & (df["رقم الهوية_standard"].str.isdigit())
            df.loc[mask, "رقم الهوية_standard"] = '0' + df.loc[mask, "رقم الهوية_standard"]
            
            # تنظيف النصوص والأسماء من الفراغات الجانبية والقيم الفارغة
            df["اسم المعلم"] = df["اسم المعلم"].fillna("").astype(str).str.strip()

        # دالة قياس نسبة التشابه المحسّنة (سريعة جداً عند التطابق التام)
        def similarity(a, b):
            if not a or not b: return 0
            if a == b: return 1.0  # اختصار يوفر طاقة المعالجة الزمنية للأسماء المتطابقة
            return SequenceMatcher(None, a, b).ratio()
        
        st.write("🏗️ بناء كشافات البحث السريع لملفات المشرفين...")
        # تحويل تجميع البيانات إلى هيكل بيانات موحد وسريع الاستدعاء O(1)
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
                # تجميع تلقائي في قوائم عند تكرار نفس الهوية عند أكثر من مشرف
                supervisor_map[id_num]["المشرف"].append(row["اسم المشرف"])
                supervisor_map[id_num]["الاسم_المدخل"].append(row["اسم المعلم"])

        st.write("🧠 تشغيل خوارزمية المطابقة الذكية وعزل الأخطاء...")
        
        results, suggestions, not_found, error_details = [], [], [], []
        
        # حصر أرقام هويات المشرفين غير الموجودة في HR لتقليص نطاق البحث الذكي (Fuzzy Matching)
        admin_ids_set = set(admin_df["رقم الهوية_standard"].unique())
        unmatched_supervisor_ids = [sid for sid in supervisor_map if sid not in admin_ids_set]
        
        # حلقة معالجة أساسية رشيقة وسريعة جداً O(N)
        for admin_row in admin_df.to_dict(orient='records'):
            admin_id = admin_row["رقم الهوية_standard"]
            admin_name = admin_row["اسم المعلم"]
            
            # الحالة 1: الرقم متطابق تماماً وموجود في سجلات المشرفين
            if admin_id in supervisor_map:
                data = supervisor_map[admin_id]
                score = data["التقييم"]
                
                statuses = []
                for sup, ent_name in zip(data["المشرف"], data["الاسم_المدخل"]):
                    if ent_name == admin_name:
                        statuses.append(f"✅ صحيح عند المشرف {sup}")
                    else:
                        statuses.append(f"⚠️ خطأ في الاسم عند المشرف {sup}")
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
            
            # الحالة 2: الرقم غير موجود في الكشاف (البحث الذكي المقيد بالفروقات البسيطة)
            else:
                found_suggestion = False
                best_match_id, best_match_name = None, None
                best_id_sim, best_name_sim = 0, 0
                
                # المقارنة الذكية تتم فقط مع الهويات التي أخطأ المشرفون في إدخالها وليس كامل الملف!
                for sup_id in unmatched_supervisor_ids:
                    id_sim = similarity(admin_id, sup_id)
                    if id_sim <= 0.6: 
                        continue  # تخطي فوري لتوفير الوقت إذا كانت الأرقام متباعدة جداً
                    
                    for sup_name in supervisor_map[sup_id]["الاسم_المدخل"]:
                        name_sim = similarity(admin_name, sup_name)
                        if name_sim > 0.85 and id_sim > 0.6:
                            if id_sim > best_id_sim:
                                best_id_sim, best_name_sim = id_sim, name_sim
                                best_match_id, best_match_name = sup_id, sup_name
                                found_suggestion = True
                
                if found_suggestion:
                    suggestions.append({
                        "👤 الاسم في (HR)": admin_name,
                        "🆔 الرقم الصحيح (HR)": admin_id,
                        "🔍 الرقم الخاطئ عند المشرف": best_match_id,
                        "📝 الاسم المدخل بواسطة المشرف": best_match_name,
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
        
        if not error_details_df.empty:
            supervisor_error_summary = error_details_df.groupby("المشرف").size().reset_index()
            supervisor_error_summary.columns = ["المشرف", "عدد الأخطاء"]
        else:
            supervisor_error_summary = pd.DataFrame()
            
        end_time = time.time()
        execution_time = end_time - start_time
        
        status.update(label=f"⚡ تمت عملية المعالجة والمطابقة بنجاح فائق خلال {execution_time:.2f} ثانية!", state="complete", expanded=False)

    # ============================================================
    # عرض الإحصائيات الذكية والمؤشرات الحيوية
    # ============================================================
    st.markdown("### 📊 نظرة عامة على البيانات")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("📋 إجمالي معلمين HR", len(admin_df))
    col2.metric("✅ تم دمجهم بنجاح", len(results_df))
    col3.metric("🔍 هويات مقترح تصحيحها", len(suggestions_df))
    col4.metric("❌ معلمين بلا تقييم", len(not_found_df))
    col5.metric("⚠️ أخطاء أسماء المشرفين", len(error_details_df))
    
    st.markdown("---")
    
    # ============================================================
    # تبويبات العرض المطورة
    # ============================================================
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 الموجودون والدمج", "🔍 مقترحات تصحيح الهوية", 
        "❌ معلمين بلا تقييم", "⚠️ رادار أخطاء المشرفين", 
        "📈 التحليلات والمخططات", "📥 مركز تحميل التقارير"
    ])
    
    with tab1:
        st.subheader("📊 بيان المعلمين المدمجة تقييماتهم")
        if not results_df.empty:
            st.dataframe(results_df, use_container_width=True, height=400)
            
            c1, c2, c3 = st.columns(3)
            c1.metric("🎯 سليم تماماً", len(results_df[results_df['📌 الحالة'].str.contains("✅")]))
            c2.metric("⚙️ يحتاج مراجعة اسم", len(results_df[results_df['📌 الحالة'].str.contains("⚠️")]))
            avg_score = pd.to_numeric(results_df['⭐ التقييم'], errors='coerce').mean()
            c3.metric("⭐ متوسط تقييمات المجموعة", f"{avg_score:.2f}" if not pd.isna(avg_score) else "N/A")
        else:
            st.info("لا توجد بيانات متاحة للعرض")
            
    with tab2:
        st.subheader("🔍 نظام المقترحات الذكي لتصحيح الهويات")
        st.info("📌 الخوارزمية رصدت هذه الحالات كـ (أخطاء رقمية مطبعية) قام بها المشرفون، حيث تطابق الاسم بنسبة عالية جداً واختلف الرقم قليلاً.")
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
            st.success("✅ مذهل! جميع الأسماء المدخلة من قبل المشرفين متطابقة مع الشؤون الإدارية.")
            
    with tab5:
        st.subheader("📈 رادار المخططات البيانية والتحليل الإحصائي")
        if not results_df.empty:
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                st.markdown("##### 📌 توزيع حالات التطابق")
                status_counts = results_df['📌 الحالة'].value_counts().reset_index()
                status_counts.columns = ['الحالة', 'العدد']
                st.bar_chart(status_counts.set_index('الحالة'))
            with col_chart2:
                st.markdown("##### 📊 الكثافة العددية للأخطاء لكل مشرف")
                if not supervisor_error_summary.empty:
                    st.bar_chart(supervisor_error_summary.set_index("المشرف"))
                else:
                    st.info("لا توجد أخطاء لرسمها بيانيا.")
        else:
            st.info("قم برفع ملفات تحتوي على بيانات صالحة لتوليد المخططات.")
            
    with tab6:
        st.subheader("📥 مركز الاستخراج وتحميل التقارير المجمعة والمنفصلة")
        
        # إنشاء ملف إكسل مجمع يحتوي على أوراق منفصلة لكل تصنيف
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            if not results_df.empty: results_df.to_excel(writer, sheet_name="الموجودين_والمدمجين", index=False)
            if not suggestions_df.empty: suggestions_df.to_excel(writer, sheet_name="اقتراحات_تعديل_الهويات", index=False)
            if not not_found_df.empty: not_found_df.to_excel(writer, sheet_name="معلمين_بدون_تقييم", index=False)
            if not error_details_df.empty: error_details_df.to_excel(writer, sheet_name="تفاصيل_أخطاء_المشرفين", index=False)
            if not supervisor_error_summary.empty: supervisor_error_summary.to_excel(writer, sheet_name="ملخص_أداء_المشرفين", index=False)
        output_buffer.seek(0)
        
        st.download_button(
            label="📥 تحميل التقرير الشامل والمدمج (ملف Excel واحد بتبويبات)",
            data=output_buffer,
            file_name="التقرير_الشامل_لدمج_المعلمين_الذكي.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        st.markdown("#### 💾 استخراج تقارير منفصلة فرعية")
        dl_col1, dl_col2, dl_col3 = st.columns(3)
        
        if not results_df.empty:
            with dl_col1:
                b1 = io.BytesIO()
                results_df.to_excel(b1, index=False, engine='openpyxl')
                st.download_button("📥 تحميل ملف المدمجين فقط", data=b1.getvalue(), file_name="المعلمين_الموجودين.xlsx")
        if not suggestions_df.empty:
            with dl_col2:
                b2 = io.BytesIO()
                suggestions_df.to_excel(b2, index=False, engine='openpyxl')
                st.download_button("📥 تحميل مقترحات تعديل الأرقام", data=b2.getvalue(), file_name="مقترحات_التصحيح.xlsx")
        if not not_found_df.empty:
            with dl_col3:
                b3 = io.BytesIO()
                not_found_df.to_excel(b3, index=False, engine='openpyxl')
                st.download_button("📥 تحميل كشف غير الموجودين", data=b3.getvalue(), file_name="غير_الموجودين.xlsx")

else:
    st.info("👈 يرجى رفع ملف المشرفين وملف الشؤون الإدارية (HR) من القائمة العلوية للبدء فورا.")
    st.markdown("""
    ### 📋 التنسيق القياسي المعتمد للمدخلات:
    * **ملف المشرفين المجمع:** يجب أن يحتوي على الأعمدة التالية: `الاسم` ، `رقم الهوية` ، `التقييم` ، `المشرف`.
    * **ملف الشؤون الإدارية (HR):** يجب أن يحتوي على الأعمدة التالية: `الاسم` ، `رقم الهوية`.
    """)

st.markdown("---")
st.caption("📌 نظام دمج تقييمات المعلمين الفائق - الإصدار المطور والمسرّع v4.1 (مصحح بالكامل وآمن)")
