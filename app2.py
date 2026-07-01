import streamlit as st
import pandas as pd
from difflib import SequenceMatcher

st.title("نظام دمج تقييمات المعلمين - مع اكتشاف الأخطاء الإملائية")

# رفع الملفات
mushrif_file = st.file_uploader("ارفع ملف المشرفين المجمّع (mushrif_all.xlsx)", type=["xlsx"])
admin_file = st.file_uploader("ارفع ملف الشؤون الإدارية (admin.xlsx)", type=["xlsx"])

if mushrif_file and admin_file:
    mushrif_df = pd.read_excel(mushrif_file)
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
        """حساب نسبة التشابه بين اسمين"""
        if not name1 or not name2:
            return 0
        return SequenceMatcher(None, name1, name2).ratio()
    
    def id_similarity(id1, id2):
        """حساب نسبة التشابه بين رقمين هوية"""
        if not id1 or not id2:
            return 0
        # مقارنة حرف بحرف
        matches = sum(1 for a, b in zip(id1, id2) if a == b)
        max_len = max(len(id1), len(id2))
        return matches / max_len if max_len > 0 else 0
    
    def find_similar_hr_id(entered_id, hr_ids, threshold=0.7):
        """البحث عن رقم هوية مشابه في ملف الإدارة"""
        if not entered_id:
            return None, 0
        
        best_match = None
        best_score = 0
        
        for hr_id in hr_ids:
            if not hr_id:
                continue
            
            # مقارنة الأرقام
            sim = id_similarity(entered_id, hr_id)
            
            # إذا كان التشابه مرتفعاً
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
    
    # قاموس: الرقم ← الاسم الصحيح
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
    
    # قاموس المطابقة: الرقم ← (التقييم, المشرف, الاسم المدخل)
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
                # في حالة التكرار، نضيف المشرف الآخر
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
    error_details = []
    similar_id_suggestions = []
    
    for _, admin_row in admin_df.iterrows():
        admin_id = admin_row["رقم الهوية_standard"]
        admin_name = admin_row["اسم المعلم"]
        
        # الحالة 1: مطابقة تامة بالرقم
        if admin_id in id_to_data:
            data = id_to_data[admin_id]
            supervisor = data["المشرف"]
            entered_name = data["الاسم_المدخل"]
            score = data["التقييم"]
            
            # التحقق من صحة الاسم
            if isinstance(entered_name, list):
                # عدة مشرفين أدخلوا نفس الرقم
                statuses = []
                for idx, (sup, ent_name) in enumerate(zip(supervisor, entered_name)):
                    if ent_name != admin_name:
                        statuses.append(f"⚠️ خطأ عند المشرف {sup}: أدخل '{ent_name}' والصحيح '{admin_name}'")
                    else:
                        statuses.append(f"✅ صحيح عند المشرف {sup}")
                status = " | ".join(statuses)
                correct_name = admin_name
            else:
                if entered_name != admin_name:
                    status = f"⚠️ خطأ في الاسم عند المشرف {supervisor}: أدخل '{entered_name}' والصحيح '{admin_name}'"
                    correct_name = admin_name
                else:
                    status = "✅ صحيح"
                    correct_name = admin_name
            
            results.append({
                "اسم المعلم (من الإدارة)": admin_name,
                "رقم الهوية": admin_id,
                "التقييم": score,
                "المشرف المسؤول": supervisor if not isinstance(supervisor, list) else " / ".join(supervisor),
                "الاسم المدخل (من المشرف)": entered_name if not isinstance(entered_name, list) else " / ".join(entered_name),
                "الحالة": status
            })
        
        # الحالة 2: الرقم غير موجود في المشرفين → نحاول البحث عن رقم مشابه
        else:
            # البحث عن اسم مشابه في المشرفين (ربما خطأ في الرقم)
            best_match_id, best_sim_score = find_similar_hr_id(admin_id, list(id_to_data.keys()), threshold=0.6)
            
            if best_match_id and best_sim_score > 0.6:
                # وجدنا رقم مشابه في المشرفين
                data = id_to_data[best_match_id]
                supervisor = data["المشرف"]
                entered_name = data["الاسم_المدخل"]
                score = data["التقييم"]
                
                # هل الاسم المدخل يشبه الاسم الصحيح؟
                name_sim = name_similarity(admin_name, entered_name if not isinstance(entered_name, list) else entered_name[0])
                
                if name_sim > 0.85:
                    status = f"🔴 خطأ في رقم الهوية: المدخل '{admin_id}' يشبه '{best_match_id}' بنسبة {best_sim_score:.0%}، والاسم متطابق تقريباً"
                    similar_id_suggestions.append({
                        "الاسم الصحيح": admin_name,
                        "الرقم الصحيح": admin_id,
                        "الرقم الخاطئ (من المشرف)": best_match_id,
                        "المشرف": supervisor,
                        "نسبة التشابه في الرقم": f"{best_sim_score:.0%}",
                        "الاسم المدخل": entered_name
                    })
                    
                    results.append({
                        "اسم المعلم (من الإدارة)": admin_name,
                        "رقم الهوية": admin_id,
                        "التقييم": score,
                        "المشرف المسؤول": supervisor,
                        "الاسم المدخل (من المشرف)": entered_name,
                        "الحالة": status
                    })
                else:
                    # الرقم مشابه لكن الاسم مختلف
                    results.append({
                        "اسم المعلم (من الإدارة)": admin_name,
                        "رقم الهوية": admin_id,
                        "التقييم": None,
                        "المشرف المسؤول": "غير موجود",
                        "الاسم المدخل (من المشرف)": "",
                        "الحالة": f"❌ لا توجد مطابقة (رقم مشابه '{best_match_id}' لكن الاسم مختلف)"
                    })
            else:
                # لا يوجد رقم مشابه
                results.append({
                    "اسم المعلم (من الإدارة)": admin_name,
                    "رقم الهوية": admin_id,
                    "التقييم": None,
                    "المشرف المسؤول": "غير موجود",
                    "الاسم المدخل (من المشرف)": "",
                    "الحالة": "❌ لا توجد مطابقة"
                })
    
    merged = pd.DataFrame(results)
    
    # ============================================================
    # عرض النتائج
    # ============================================================
    
    st.subheader("📊 القائمة الإدارية المكتملة")
    st.dataframe(merged)
    
    matched = len(merged[merged['التقييم'].notna()])
    total = len(merged)
    st.info(f"✅ تم تحديث {matched} من {total} سجل ({matched/total*100:.1f}%)")
    
    # ============================================================
    # عرض الأخطاء المقترحة (أرقام متشابهة)
    # ============================================================
    
    if similar_id_suggestions:
        st.subheader("🔍 اقتراحات: أرقام هوية متشابهة (ربما خطأ إملائي)")
        st.warning(f"⚠️ تم اكتشاف {len(similar_id_suggestions)} حالة لرقم هوية خاطئ يشبه رقماً آخر")
        suggestions_df = pd.DataFrame(similar_id_suggestions)
        st.dataframe(suggestions_df)
        
        # عرض تفصيل للحالة المذكورة
        st.subheader("📝 مثال: ميساء محمد عمر صلاح")
        example = suggestions_df[suggestions_df["الاسم الصحيح"].str.contains("ميساء", na=False)]
        if not example.empty:
            st.info(f"""
            **المشكلة:**  
            - الاسم الصحيح: {example.iloc[0]['الاسم الصحيح']}  
            - الرقم الصحيح: {example.iloc[0]['الرقم الصحيح']}  
            - الرقم الخاطئ (من المشرف): {example.iloc[0]['الرقم الخاطئ (من المشرف)']}  
            - المشرف: {example.iloc[0]['المشرف']}  
            - نسبة التشابه في الرقم: {example.iloc[0]['نسبة التشابه في الرقم']}  
            
            **ملاحظة:** يبدو أن المشرف أخطأ في إدخال رقم واحد (6 أصبح 3)
            """)
    
    # ============================================================
    # تفاصيل الأخطاء لكل مشرف
    # ============================================================
    
    st.subheader("📋 تفاصيل الأخطاء حسب المشرف")
    
    # استخراج الأخطاء من الحالة
    error_list = []
    for _, row in merged.iterrows():
        if "خطأ" in row["الحالة"]:
            error_list.append({
                "اسم المعلم": row["اسم المعلم (من الإدارة)"],
                "رقم الهوية": row["رقم الهوية"],
                "المشرف": row["المشرف المسؤول"],
                "الحالة": row["الحالة"]
            })
    
    if error_list:
        errors_df = pd.DataFrame(error_list)
        st.dataframe(errors_df)
        
        # ملخص الأخطاء لكل مشرف
        st.subheader("📊 ملخص الأخطاء لكل مشرف")
        error_summary = errors_df.groupby("المشرف").size().reset_index()
        error_summary.columns = ["المشرف", "عدد الأخطاء"]
        st.dataframe(error_summary)
    else:
        st.success("✅ لا توجد أخطاء")
    
    # ============================================================
    # تحميل الملفات
    # ============================================================
    
    import io
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
    
    st.subheader("خيارات العرض")
    if st.button("عرض المعلمين الذين لديهم مشاكل فقط"):
        filtered = merged[merged["الحالة"].str.contains("خطأ|لا توجد", na=False)].copy()
        if filtered.empty:
            st.info("لا توجد مشاكل حالياً ✅")
        else:
            st.dataframe(filtered)
    
    # تصفية حسب المشرف
    st.subheader("تصفية حسب المشرف")
    supervisors_list = mushrif_df["اسم المشرف"].dropna().unique().tolist()
    selected_supervisor = st.selectbox("اختر اسم المشرف:", supervisors_list)
    if selected_supervisor:
        filtered_by_supervisor = merged[merged["المشرف المسؤول"].str.contains(selected_supervisor, na=False)]
        if filtered_by_supervisor.empty:
            st.info(f"لا توجد بيانات للمدرسين تحت إشراف {selected_supervisor} ✅")
        else:
            st.dataframe(filtered_by_supervisor)