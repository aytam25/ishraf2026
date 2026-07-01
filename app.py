import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io

# ============================================================
# إعدادات الصفحة - عريضة
# ============================================================
st.set_page_config(
    page_title="نظام دمج تقييمات المعلمين",
    page_icon="📊",
    layout="wide"  # ✅ يجعل الصفحة عريضة
)

st.title("📊 نظام دمج تقييمات المعلمين - مع اكتشاف الأخطاء الإملائية")

# ============================================================
# رفع الملفات
# ============================================================
col1, col2 = st.columns(2)
with col1:
    mushrif_file = st.file_uploader("ارفع ملف المشرفين المجمّع (mushrif_all.xlsx)", type=["xlsx"])
with col2:
    admin_file = st.file_uploader("ارفع ملف الشؤون الإدارية (admin.xlsx)", type=["xlsx"])

if mushrif_file and admin_file:
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
    # دوال معالجة الأرقام والأسماء
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
    
    def name_similarity(name1, name2):
        if not name1 or not name2:
            return 0
        return SequenceMatcher(None, name1, name2).ratio()
    
    def id_similarity(id1, id2):
        if not id1 or not id2:
            return 0
        matches = sum(1 for a, b in zip(id1, id2) if a == b)
        max_len = max(len(id1), len(id2))
        return matches / max_len if max_len > 0 else 0
    
    def find_similar_hr_id(entered_id, hr_ids, threshold=0.6):
        if not entered_id:
            return None, 0
        best_match = None
        best_score = 0
        for hr_id in hr_ids:
            if not hr_id:
                continue
            sim = id_similarity(entered_id, hr_id)
            if sim > best_score and sim >= threshold:
                best_score = sim
                best_match = hr_id
        return best_match, best_score
    
    # تطبيق التطبيع
    mushrif_df["رقم الهوية"] = mushrif_df["رقم الهوية"].apply(normalize_id)
    mushrif_df["اسم المعلم"] = mushrif_df["اسم المعلم"].apply(clean_text)
    mushrif_df["رقم الهوية_standard"] = mushrif_df["رقم الهوية"].apply(pad_to_9)
    
    admin_df["رقم الهوية"] = admin_df["رقم الهوية"].apply(normalize_id)
    admin_df["اسم المعلم"] = admin_df["اسم المعلم"].apply(clean_text)
    admin_df["رقم الهوية_standard"] = admin_df["رقم الهوية"].apply(pad_to_9)
    
    # ============================================================
    # بناء قوائم HR للمقارنة
    # ============================================================
    
    hr_name_map = {}
    hr_ids_list = []
    for _, row in admin_df.iterrows():
        id_num = row["رقم الهوية_standard"]
        name = row["اسم المعلم"]
        if id_num and name:
            hr_name_map[id_num] = name
            hr_ids_list.append(id_num)
    
    # ============================================================
    # كشف التكرارات والأخطاء في المشرفين
    # ============================================================
    
    id_to_data = {}
    for _, row in mushrif_df.iterrows():
        id_num = row["رقم الهوية_standard"]
        if id_num:
            if id_num not in id_to_data:
                id_to_data[id_num] = {
                    "التقييم": row["التقييم"],
                    "المشرف": row["اسم المشرف"],
                    "الاسم_المدخل": row["اسم المعلم"]
                }
            else:
                existing = id_to_data[id_num]
                if isinstance(existing["المشرف"], list):
                    existing["المشرف"].append(row["اسم المشرف"])
                    existing["الاسم_المدخل"].append(row["اسم المعلم"])
                else:
                    existing["المشرف"] = [existing["المشرف"], row["اسم المشرف"]]
                    existing["الاسم_المدخل"] = [existing["الاسم_المدخل"], row["اسم المعلم"]]
    
    # ============================================================
    # المطابقة الذكية مع اكتشاف الأخطاء
    # ============================================================
    
    results = []
    similar_id_suggestions = []
    
    for _, admin_row in admin_df.iterrows():
        admin_id = admin_row["رقم الهوية_standard"]
        admin_name = admin_row["اسم المعلم"]
        
        if admin_id in id_to_data:
            data = id_to_data[admin_id]
            supervisor = data["المشرف"]
            entered_name = data["الاسم_المدخل"]
            score = data["التقييم"]
            
            if isinstance(entered_name, list):
                statuses = []
                for idx, (sup, ent_name) in enumerate(zip(supervisor, entered_name)):
                    if ent_name != admin_name:
                        statuses.append(f"⚠️ خطأ عند المشرف {sup}")
                    else:
                        statuses.append(f"✅ صحيح عند المشرف {sup}")
                status = " | ".join(statuses)
            else:
                if entered_name != admin_name:
                    status = f"⚠️ خطأ عند المشرف {supervisor}"
                else:
                    status = "✅ صحيح"
            
            results.append({
                "👤 الاسم (الإدارة)": admin_name,
                "🆔 الهوية": admin_id,
                "⭐ التقييم": score,
                "👨‍🏫 المشرف": supervisor if not isinstance(supervisor, list) else " / ".join(supervisor),
                "📝 الاسم المدخل": entered_name if not isinstance(entered_name, list) else " / ".join(entered_name),
                "📌 الحالة": status
            })
        
        else:
            best_match_id, best_sim_score = find_similar_hr_id(admin_id, list(id_to_data.keys()), threshold=0.6)
            
            if best_match_id and best_sim_score > 0.6:
                data = id_to_data[best_match_id]
                supervisor = data["المشرف"]
                entered_name = data["الاسم_المدخل"]
                score = data["التقييم"]
                name_sim = name_similarity(admin_name, entered_name if not isinstance(entered_name, list) else entered_name[0])
                
                if name_sim > 0.85:
                    status = f"🔴 خطأ في الرقم (يشبه {best_match_id} بنسبة {best_sim_score:.0%})"
                    similar_id_suggestions.append({
                        "الاسم الصحيح": admin_name,
                        "الرقم الصحيح": admin_id,
                        "الرقم الخاطئ": best_match_id,
                        "المشرف": supervisor,
                        "نسبة التشابه": f"{best_sim_score:.0%}",
                        "الاسم المدخل": entered_name
                    })
                else:
                    status = f"❌ لا توجد مطابقة (رقم مشابه '{best_match_id}' لكن اسم مختلف)"
                
                results.append({
                    "👤 الاسم (الإدارة)": admin_name,
                    "🆔 الهوية": admin_id,
                    "⭐ التقييم": score,
                    "👨‍🏫 المشرف": supervisor,
                    "📝 الاسم المدخل": entered_name,
                    "📌 الحالة": status
                })
            else:
                results.append({
                    "👤 الاسم (الإدارة)": admin_name,
                    "🆔 الهوية": admin_id,
                    "⭐ التقييم": None,
                    "👨‍🏫 المشرف": "غير موجود",
                    "📝 الاسم المدخل": "",
                    "📌 الحالة": "❌ لا توجد مطابقة"
                })
    
    merged = pd.DataFrame(results)
    
    # ============================================================
    # عرض النتائج - جدول عريض
    # ============================================================
    
    st.subheader("📊 القائمة الإدارية المكتملة")
    
    # ✅ عرض الجدول بعرض كامل مع تمرير أفقي
    st.dataframe(
        merged,
        use_container_width=True,  # ✅ يأخذ عرض الصفحة بالكامل
        height=400,  # ارتفاع ثابت مع تمرير عمودي
        column_config={
            "👤 الاسم (الإدارة)": st.column_config.TextColumn("👤 الاسم (الإدارة)", width="medium"),
            "🆔 الهوية": st.column_config.TextColumn("🆔 الهوية", width="small"),
            "⭐ التقييم": st.column_config.TextColumn("⭐ التقييم", width="small"),
            "👨‍🏫 المشرف": st.column_config.TextColumn("👨‍🏫 المشرف", width="medium"),
            "📝 الاسم المدخل": st.column_config.TextColumn("📝 الاسم المدخل", width="medium"),
            "📌 الحالة": st.column_config.TextColumn("📌 الحالة", width="large"),
        }
    )
    
    # إحصائيات سريعة
    matched = len(merged[merged['⭐ التقييم'].notna()])
    total = len(merged)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📊 إجمالي المعلمين", total)
    col2.metric("✅ تم تحديث", f"{matched} من {total}", f"{matched/total*100:.1f}%")
    col3.metric("⚠️ أخطاء", len(merged[merged['📌 الحالة'].str.contains("خطأ", na=False)]))
    col4.metric("❌ بدون مطابقة", len(merged[merged['📌 الحالة'].str.contains("لا توجد", na=False)]))
    
    # ============================================================
    # 📊 مخططات بيانية خفيفة (تستخدم cache للحفاظ على الأداء)
    # ============================================================
    
    st.subheader("📊 التحليلات البيانية")
    
    # تبويبات للمخططات
    tab1, tab2, tab3 = st.tabs(["📈 توزيع الحالات", "👨‍🏫 أخطاء المشرفين", "📊 إحصائيات عامة"])
    
    with tab1:
        # توزيع الحالات
        status_counts = merged['📌 الحالة'].value_counts().reset_index()
        status_counts.columns = ['الحالة', 'العدد']
        
        col1, col2 = st.columns(2)
        with col1:
            st.dataframe(status_counts, use_container_width=True)
        with col2:
            st.bar_chart(
                status_counts.set_index('الحالة'),
                use_container_width=True
            )
    
    with tab2:
        # أخطاء المشرفين
        error_data = merged[merged['📌 الحالة'].str.contains("خطأ", na=False)]
        if not error_data.empty:
            supervisor_errors = error_data['👨‍🏫 المشرف'].value_counts().reset_index()
            supervisor_errors.columns = ['المشرف', 'عدد الأخطاء']
            
            col1, col2 = st.columns(2)
            with col1:
                st.dataframe(supervisor_errors, use_container_width=True)
            with col2:
                st.bar_chart(
                    supervisor_errors.set_index('المشرف'),
                    use_container_width=True
                )
        else:
            st.success("✅ لا توجد أخطاء لعرضها")
    
    with tab3:
        col1, col2 = st.columns(2)
        with col1:
            # نسبة المطابقات
            match_rates = pd.DataFrame({
                'الفئة': ['مطابق', 'غير مطابق'],
                'العدد': [
                    len(merged[merged['📌 الحالة'].str.contains("صحيح|خطأ", na=False)]),
                    len(merged[merged['📌 الحالة'].str.contains("لا توجد", na=False)])
                ]
            })
            st.dataframe(match_rates, use_container_width=True)
            
            # إضافة شريط تقدم بسيط
            match_pct = len(merged[merged['⭐ التقييم'].notna()]) / len(merged) * 100
            st.progress(match_pct / 100, text=f"نسبة المطابقة: {match_pct:.1f}%")
        
        with col2:
            # متوسط التقييمات (إن كانت رقمية)
            numeric_scores = pd.to_numeric(merged['⭐ التقييم'], errors='coerce')
            if not numeric_scores.dropna().empty:
                st.metric("📊 متوسط التقييم", f"{numeric_scores.mean():.2f}")
                st.metric("📈 أعلى تقييم", f"{numeric_scores.max():.0f}")
                st.metric("📉 أقل تقييم", f"{numeric_scores.min():.0f}")
            else:
                st.info("لا توجد بيانات تقييم رقمية للتحليل")
    
    # ============================================================
    # عرض الأخطاء المقترحة (أرقام متشابهة)
    # ============================================================
    
    if similar_id_suggestions:
        with st.expander("🔍 اقتراحات: أرقام هوية متشابهة (ربما خطأ إملائي)"):
            st.warning(f"⚠️ تم اكتشاف {len(similar_id_suggestions)} حالة")
            suggestions_df = pd.DataFrame(similar_id_suggestions)
            st.dataframe(suggestions_df, use_container_width=True)
    
    # ============================================================
    # تفاصيل الأخطاء لكل مشرف
    # ============================================================
    
    with st.expander("📋 تفاصيل الأخطاء حسب المشرف"):
        error_list = []
        for _, row in merged.iterrows():
            if "خطأ" in str(row['📌 الحالة']):
                error_list.append({
                    "اسم المعلم": row['👤 الاسم (الإدارة)'],
                    "رقم الهوية": row['🆔 الهوية'],
                    "المشرف": row['👨‍🏫 المشرف'],
                    "الحالة": row['📌 الحالة']
                })
        
        if error_list:
            errors_df = pd.DataFrame(error_list)
            st.dataframe(errors_df, use_container_width=True)
            
            st.subheader("📊 ملخص الأخطاء لكل مشرف")
            error_summary = errors_df.groupby("المشرف").size().reset_index()
            error_summary.columns = ["المشرف", "عدد الأخطاء"]
            st.dataframe(error_summary, use_container_width=True)
            
            # رسم بياني لأخطاء المشرفين
            st.bar_chart(error_summary.set_index("المشرف"), use_container_width=True)
        else:
            st.success("✅ لا توجد أخطاء")
    
    # ============================================================
    # تحميل الملفات
    # ============================================================
    
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        merged.to_excel(writer, sheet_name="القائمة_الإدارية", index=False)
        if error_list:
            errors_df.to_excel(writer, sheet_name="تفاصيل_الأخطاء", index=False)
        if similar_id_suggestions:
            suggestions_df.to_excel(writer, sheet_name="أرقام_متشابهة", index=False)
    output_buffer.seek(0)
    
    st.subheader("⬇️ تحميل الملفات الناتجة")
    st.download_button(
        label="📥 تحميل القائمة الإدارية المكتملة (Excel)",
        data=output_buffer,
        file_name="admin_completed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # ============================================================
    # خيارات العرض
    # ============================================================
    
    with st.expander("🎯 خيارات التصفية"):
        col1, col2 = st.columns(2)
        with col1:
            if st.button("عرض المعلمين الذين لديهم مشاكل فقط"):
                filtered = merged[merged["📌 الحالة"].str.contains("خطأ|لا توجد", na=False)].copy()
                if filtered.empty:
                    st.info("✅ لا توجد مشاكل حالياً")
                else:
                    st.dataframe(filtered, use_container_width=True)
        
        with col2:
            supervisors_list = mushrif_df["اسم المشرف"].dropna().unique().tolist()
            if supervisors_list:
                selected_supervisor = st.selectbox("اختر اسم المشرف:", supervisors_list)
                if selected_supervisor:
                    filtered_by_supervisor = merged[merged["👨‍🏫 المشرف"].str.contains(selected_supervisor, na=False)]
                    if filtered_by_supervisor.empty:
                        st.info(f"لا توجد بيانات للمدرسين تحت إشراف {selected_supervisor}")
                    else:
                        st.dataframe(filtered_by_supervisor, use_container_width=True)

else:
    st.info("👈 يرجى رفع ملف المشرفين وملف الشؤون الإدارية للبدء")
