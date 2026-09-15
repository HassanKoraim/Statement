import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- دوال المساعدة ---
def calculate_contract_duration(start, end):
    if pd.isna(start) or pd.isna(end):
        return "غير محدد"
    delta = end - start
    years = delta.days // 365
    months = (delta.days % 365) // 30
    
    parts = []
    if years > 0:
        parts.append(f"{years} سنة")
    if months > 0:
        parts.append(f"{months} شهر")
        
    return " و ".join(parts) if parts else f"{delta.days} يوم"

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='الإفادة')
    processed_data = output.getvalue()
    return processed_data

# --- إعدادات الصفحة ---
st.set_page_config(page_title="نظام الإفادات", page_icon="📊", layout="wide")
st.title("📊 نظام استخراج إفادات شؤون الموظفين")

# --- إعدادات الفلترة ---
st.sidebar.header("⚙️ إعدادات الفلترة")
st.sidebar.info("حدد نطاق الأيام للبحث عن الإفادات التي ستنتهي خلال هذه الفترة:")
min_days, max_days = st.sidebar.slider("نطاق الأيام (من - إلى):", min_value=0, max_value=90, value=(0, 5))

# --- رفع الملف ---
uploaded_file = st.file_uploader('قم برفع ملف "إفادات للجيمناي.xlsx"', type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # قراءة الملف
        df = pd.read_excel(uploaded_file)
        
        # توحيد أسماء الأعمدة
        df.columns = df.columns.str.strip()
        col_mapping = {
            'الرقم الوظيفي': 'رقم وظيفي',
            'تاريخ الانضام الفعلي': 'تاريخ الانضمام الفعلي', 
            'تاريخ بداية العقد ': 'تاريخ بداية العقد',
            'فترة الاشعار': 'فترة الإشعار'
        }
        df.rename(columns=col_mapping, inplace=True)
        
        # إضافة الأعمدة الجديدة للقائمة المتوقعة
        expected_cols = [
            'رقم وظيفي', 'الاسم', 'صاحب العمل', 'تاريخ الانضمام الفعلي', 
            'تاريخ بداية العقد', 'تاريخ المباشرة بالعقد', 'تاريخ نهاية العقد', 
            'فترة التجربة', 'فترة الإشعار'
        ]
        missing_cols = [col for col in expected_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"عذراً، الشيت المرفوع ينقصه الأعمدة التالية: {', '.join(missing_cols)}")
        else:
            # تحويل أعمدة التواريخ (تمت إضافة تاريخ المباشرة بالعقد)
            date_cols = ['تاريخ الانضمام الفعلي', 'تاريخ بداية العقد', 'تاريخ المباشرة بالعقد', 'تاريخ نهاية العقد']
            for col in date_cols:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                
            df['فترة التجربة'] = pd.to_numeric(df['فترة التجربة'], errors='coerce').fillna(0)
            df['فترة الإشعار'] = pd.to_numeric(df['فترة الإشعار'], errors='coerce').fillna(0)
            
            today = pd.to_datetime(datetime.today().date())
            
            st.divider()
            
            # --- 1. إفادة تجديد العقود ---
            st.subheader("1️⃣ إفادة تجديد العقود")
            df['المدة المتبقية لإشعار المدير المباشر'] = (df['تاريخ نهاية العقد'] - pd.to_timedelta(df['فترة الإشعار'] + 60, unit='d') - today).dt.days
            rep1_df = df[(df['المدة المتبقية لإشعار المدير المباشر'] >= min_days) & (df['المدة المتبقية لإشعار المدير المباشر'] <= max_days)].copy()
            
            if not rep1_df.empty:
                rep1_df['مدة العقد'] = rep1_df.apply(lambda row: calculate_contract_duration(row['تاريخ بداية العقد'], row['تاريخ نهاية العقد']), axis=1)
                
                # إضافة عمود "الاسم"
                export_df1 = rep1_df[['رقم وظيفي', 'الاسم', 'مدة العقد', 'المدة المتبقية لإشعار المدير المباشر']]
                st.dataframe(export_df1, use_container_width=True)
                
                st.download_button(
                    label="📥 تحميل إفادة تجديد العقود (Excel)",
                    data=to_excel(export_df1),
                    file_name="تجديد_العقود.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info(f"لا يوجد إفادات لتجديد العقود خلال {min_days} إلى {max_days} أيام.")

            empty_employer = df['صاحب العمل'].isnull() | (df['صاحب العمل'].astype(str).str.strip().isin(['', '-', 'nan', 'null']))
            
            # --- 2. إفادة فترة تجربة (فعلي) ---
            st.subheader("2️⃣ إفادة فترة تجربة (فعلي)")
            df['المدة المتبقية لإشعار المدير المباشر'] = (df['تاريخ الانضمام الفعلي'] + pd.to_timedelta(df['فترة التجربة'] - 30, unit='d') - today).dt.days
            rep2_df = df[empty_employer & (df['المدة المتبقية لإشعار المدير المباشر'] >= min_days) & (df['المدة المتبقية لإشعار المدير المباشر'] <= max_days)].copy()
            
            if not rep2_df.empty:
                # إضافة عمود "الاسم"
                export_df2 = rep2_df[['رقم وظيفي', 'الاسم', 'المدة المتبقية لإشعار المدير المباشر', 'تاريخ الانضمام الفعلي']]
                export_df2['تاريخ الانضمام الفعلي'] = export_df2['تاريخ الانضمام الفعلي'].dt.strftime('%Y-%m-%d')
                
                st.dataframe(export_df2, use_container_width=True)
                st.download_button(
                    label="📥 تحميل إفادة تجربة الفعلي (Excel)",
                    data=to_excel(export_df2),
                    file_name="تجربة_فعلي.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info(f"لا يوجد موظفين ستنتهي فترة تجربتهم الفعلية خلال {min_days} إلى {max_days} أيام.")

            # --- 3. إفادة فترة تجربة (قوى) ---
            st.subheader("3️⃣ إفادة فترة تجربة (قوى)")
            # تغيير الحسبة لتعتمد على "تاريخ المباشرة بالعقد" بدلاً من "تاريخ بداية العقد"
            df['المدة المتبقية لإشعار المدير المباشر'] = (df['تاريخ المباشرة بالعقد'] + pd.to_timedelta(df['فترة التجربة'] - 30, unit='d') - today).dt.days
            rep3_df = df[(~empty_employer) & (df['المدة المتبقية لإشعار المدير المباشر'] >= min_days) & (df['المدة المتبقية لإشعار المدير المباشر'] <= max_days)].copy()
            
            if not rep3_df.empty:
                # إضافة عمود "الاسم" وعرض "تاريخ المباشرة بالعقد"
                export_df3 = rep3_df[['رقم وظيفي', 'الاسم', 'المدة المتبقية لإشعار المدير المباشر', 'تاريخ المباشرة بالعقد', 'صاحب العمل']]
                export_df3['تاريخ المباشرة بالعقد'] = export_df3['تاريخ المباشرة بالعقد'].dt.strftime('%Y-%m-%d')
                
                st.dataframe(export_df3, use_container_width=True)
                st.download_button(
                    label="📥 تحميل إفادة تجربة قوى (Excel)",
                    data=to_excel(export_df3),
                    file_name="تجربة_قوى.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info(f"لا يوجد موظفين ستنتهي فترة تجربتهم في قوى خلال {min_days} إلى {max_days} أيام.")

    except Exception as e:
        st.error(f"حدث خطأ أثناء معالجة الملف: {e}")