import streamlit as st
import pandas as pd
from difflib import SequenceMatcher
import io

# ============================================================
# إعدادات الصفحة
# ============================================================
st.set_page_config(
    page_title="نظام دمج تقييمات المعلمين - الذكي",
    page_icon="📊",
    layout="wide"
)

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
        
        results = []           # المعلمين الموجودين في الإدارة
        not_found = []         # المعلمين غير الموجودين
        suggestions = []       # اقتراحات تصحيح الرقم
        
        for _, admin_row in admin_df.iterrows():
            admin_id = admin_row["رقم الهوية_standard"]
            admin_name = admin_row["اسم المعلم"]
            
            # الحالة 1: الرقم موجود في المشرفين
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
                    status = " | ".join(statuses)
                    display_supervisor = " / ".join(supervisor)
                    display_name = " / ".join(entered_name)
                else:
                    if entered_name == admin_name:
                        status = "✅ صحيح"
                    else:
                        status = f"⚠️ خطأ في الاسم: أدخل '{entered_name}' والصحيح '{admin_name}'"
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
            
            # الحالة 2: الرقم غير موجود في المشرفين
            else:
                # نبحث عن اسم مشابه في HR
                found_suggestion = False
                best_match_id = None
                best_match_name = None
                best_sim = 0
                
                for hr_id, hr_name in hr_map.items():
                    if hr_id != admin_id:
                        id_sim = similarity(admin_id, hr_id)
                        name_sim = similarity(admin_name, hr_name)
                        
                        # إذا كان الاسم مشابهاً جداً، نعتبره اقتراح تصحيح
                        if name_sim > 0.85 and id_sim > 0.6:
                            if id_sim > best_sim:
                                best_sim = id_sim
                                best_match_id = hr_id
                                best_match_name = hr_name
                                found_suggestion = True
                
                if found_suggestion:
                    # اقتراح تصحيح الرقم
                    suggestions.append({
                        "👤 الاسم (المدخل)": admin_name,
                        "🆔 الرقم (المدخل)": admin_id,
                        "🔍 الرقم المقترح": best_match_id,
                        "✅ الاسم الصحيح في HR": best_match_name,
                        "📊 نسبة التشابه": f"{best_sim:.0%}"
                    })
                else:
                    # معلم غير موجود تماماً
                    not_found.append({
                        "👤 الاسم (المدخل)": admin_name,
                        "🆔 الرقم (المدخل)": admin_id,
                        "📌 الملاحظة": "لم يتم العثور على هذا المعلم في HR ولا يوجد اسم مشابه"
                    })
        
        # ============================================================
        # عرض النتائج
        # ============================================================
        
        # 1. المعلمين الموجودين
        if results:
            st.subheader("📊 المعلمين الموجودين في الإدارة")
            merged = pd.DataFrame(results)
            st.dataframe(merged, use_container_width=True, height=400)
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("📊 إجمالي الموجودين", len(merged))
            matched = len(merged[merged['⭐ التقييم'].notna()])
            col2.metric("✅ تم تحديث", f"{matched} من {len(merged)}")
            errors = len(merged[merged['📌 الحالة'].str.contains("خطأ", na=False)])
            col3.metric("⚠️ أخطاء في الاسم", errors)
            col4.metric("✅ صحيح", len(merged[merged['📌 الحالة'] == "✅ صحيح"]))
        
        # 2. اقتراحات التصحيح
        if suggestions:
            st.subheader("🔍 اقتراحات تصحيح رقم الهوية")
            st.info("📌 هذه الحالات قد يكون فيها خطأ إملائي في رقم واحد أو اثنين، والاسم متطابق")
            suggestions_df = pd.DataFrame(suggestions)
            st.dataframe(suggestions_df, use_container_width=True)
        
        # 3. المعلمين غير الموجودين
        if not_found:
            st.subheader("❌ المعلمين غير الموجودين في الإدارة")
            st.warning(f"⚠️ {len(not_found)} معلم غير موجود في قاعدة البيانات")
            not_found_df = pd.DataFrame(not_found)
            st.dataframe(not_found_df, use_container_width=True)
        
        # ============================================================
        # تحميل الملفات
        # ============================================================
        
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            if results:
                merged.to_excel(writer, sheet_name="الموجودين", index=False)
            if suggestions:
                suggestions_df.to_excel(writer, sheet_name="اقتراحات_تصحيح", index=False)
            if not_found:
                not_found_df.to_excel(writer, sheet_name="غير_الموجودين", index=False)
        output_buffer.seek(0)
        
        st.subheader("⬇️ تحميل التقرير الكامل")
        st.download_button(
            label="📥 تحميل ملف Excel",
            data=output_buffer,
            file_name="تقرير_المعلمين.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("👈 يرجى رفع ملف المشرفين وملف الشؤون الإدارية للبدء")
    st.markdown("""
    ### 📋 تنسيق الملفات المطلوب:
    
    **ملف المشرفين:**
    - يحتوي على أعمدة: `الاسم`, `رقم الهوية`, `التقييم`, `المشرف`
    
    **ملف الشؤون الإدارية (HR):**
    - يحتوي على أعمدة: `الاسم`, `رقم الهوية`
    """)
