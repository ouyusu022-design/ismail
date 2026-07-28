import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import json
from shapely.geometry import Polygon, MultiPolygon
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import base64

# -------------------------------------------------------------
# 1. اعدادات الصفحة والمظهر
# -------------------------------------------------------------
st.set_page_config(
    page_title="نظام إدارة رخص المناجم (Direction des Mines)",
    page_icon="🔒",
    layout="wide"
)

# -------------------------------------------------------------
# 2. دوال إدارة المستخدمين (JSON)
# -------------------------------------------------------------
USERS_FILE = 'users.json'

def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    # حساب افتراضي في حال عدم وجود الملف
    return {
        "admin": {
            "password": "admin",
            "role": "admin"
        }
    }

def save_users(users_data):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users_data, f, ensure_ascii=False, indent=4)

# -------------------------------------------------------------
# 3. إدارة جلسة المستخدم (Session State)
# -------------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'username' not in st.session_state:
    st.session_state['username'] = ''
if 'role' not in st.session_state:
    st.session_state['role'] = ''

# -------------------------------------------------------------
# 🔐 الواجهة الأولى: تسجيل الدخول فقط (بدون خيار إنشاء حساب)
# -------------------------------------------------------------
if not st.session_state['logged_in']:
    st.title("🔒 نظام إدارة رخص المناجم (Direction des Mines)")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("🔑 تسجيل الدخول")
        st.write("المرجو إدخال اسم المستخدم وكلمة السر للولوج إلى النظام.")
        
        username_input = st.text_input("اسم المستخدم (Nom d'utilisateur):")
        password_input = st.text_input("كلمة السر (Mot de passe):", type="password")
        
        if st.button("تسجيل الدخول 🔑", use_container_width=True):
            users = load_users()
            if username_input in users and users[username_input]['password'] == password_input:
                st.session_state['logged_in'] = True
                st.session_state['username'] = username_input
                st.session_state['role'] = users[username_input].get('role', 'user')
                st.success(f"مرحباً بك {username_input}")
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة السر غير صحيحة!")

# -------------------------------------------------------------
# 🏢 الواجهة الثانية: داخل المنظومة بعد تسجيل الدخول
# -------------------------------------------------------------
else:
    # القائمة الجانبية والمعلومات الشخصية
    st.sidebar.title("🏢 Direction des Mines")
    st.sidebar.markdown(f"👤 **المستخدم:** `{st.session_state['username']}`")
    st.sidebar.markdown(f"🔰 ** الصلاحية:** `{st.session_state['role']}`")
    st.sidebar.markdown("---")
    
    if st.sidebar.button("تسجيل الخروج 🚪", use_container_width=True):
        st.session_state['logged_in'] = False
        st.session_state['username'] = ''
        st.session_state['role'] = ''
        st.rerun()

    st.title("🔒 نظام إدارة رخص المناجم (Direction des Mines)")

    # إظهار التبويبات حسب صلاحية الحساب (Admin vs User)
    if st.session_state['role'] == 'admin':
        tab_main, tab_admin = st.tabs(["📊 إدارة الرخص والبيانات", "👨‍💼 لوحة التحكم (إضافة موظف)"])
    else:
        tab_main = st.container()
        tab_admin = None

    # --- تبويب العمل الرئيسي (إدارة الرخص) ---
    with tab_main if st.session_state['role'] == 'admin' else st.container():
        st.subheader("📋 لوحة معالجة الرخص والبيانات")
        
        # قسم تحميل واستعراض الملفات
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            excel_file = st.file_uploader("رفع ملف الرخص (Excel/CSV)", type=["xlsx", "csv"])
        with col_f2:
            shp_files = st.file_uploader("رفع ملفات الشيب فايل (Shapefile)", type=["shp", "dbf", "shx", "prj"], accept_multiple_files=True)

        if excel_file is not None:
            try:
                if excel_file.name.endswith('.csv'):
                    df = pd.read_csv(excel_file)
                else:
                    df = pd.read_excel(excel_file)
                st.success("تم تحميل جدول البيانات بنجاح!")
                st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

    # --- تبويب لوحة التحكم بالموظفين (للأدمن فقط) ---
    if tab_admin is not None:
        with tab_admin:
            st.subheader("👨‍💼 لوحة إدارة حسابات الموظفين")
            st.info("💡 من هنا يمكنك إضافة حسابات جديدة للموظفين للولوج إلى النظام.")
            
            with st.form("add_user_form", clear_on_submit=True):
                new_user = st.text_input("اسم الموظف الجديد (Nom d'utilisateur):")
                new_pass = st.text_input("كلمة السر (Mot de passe):", type="password")
                new_role = st.selectbox("نوع الصلاحية (Rôle):", ["user", "admin"], format_func=lambda x: "موظف عادي (User)" if x == "user" else "مسؤول (Admin)")
                
                submit_btn = st.form_submit_button("إضافة الموظف 💾")
                
                if submit_btn:
                    if new_user.strip() != "" and new_pass.strip() != "":
                        users = load_users()
                        if new_user in users:
                            st.warning("⚠️ هذا الاسم مستخدم بالفعل، اختر اسم آخر.")
                        else:
                            users[new_user] = {
                                "password": new_pass,
                                "role": new_role
                            }
                            save_users(users)
                            st.success(f"✅ تمت إضافة الموظف ({new_user}) بنجاح!")
                    else:
                        st.error("⚠️ المرجو ملء جميع الحقول المطلوبة.")

            st.markdown("---")
            st.write("📋 **قائمة الحسابات المسجلة حالياً:**")
            current_users = load_users()
            users_list = [{"المستخدم": u, "الصلاحية": current_users[u].get('role', 'user')} for u in current_users]
            st.table(pd.DataFrame(users_list))
