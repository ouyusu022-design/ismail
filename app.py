import streamlit as st

# 1. إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="نظام إدارة الملفات",
    page_icon="📂",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. كود CSS الشامل لإخفاء جميع عناصر Streamlit الخارجية (Viewer Badge, Manage App, Header, Footer)
st.markdown("""
    <style>
    /* إخفاء الهيدر العلوي وشريط الأدوات */
    header, #MainMenu, [data-testid="stHeader"], [data-testid="stToolbar"] {
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
    }

    /* إخفاء الفوتر */
    footer {
        display: none !important;
        visibility: hidden !important;
    }

    /* إخفاء زر Manage App، الـ Viewer Badge، وأيقونة الأوفاتار */
    [data-testid="stAppToolbar"],
    [data-testid="stStatusWidget"],
    [data-testid="stAppViewerBadge"],
    [data-testid="stViewerBadge"],
    .stAppViewerBadge,
    .stStatusWidget,
    #ManageAppButton,
    div[class*="viewerBadge"],
    div[class*="StatusWidget"],
    div[class*="styles_viewerBadge"],
    div[class*="Profile"],
    div[data-testid="stDecoration"],
    iframe ~ div,
    button[aria-label="Manage app"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        width: 0px !important;
        height: 0px !important;
        pointer-events: none !important;
    }

    /* ضبط المسافات الفوقانية والتحتانية */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 0rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 3. إدارة حالة الجلسة (Session State)
# -------------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# -------------------------------------------------------------
# 4. واجهة تسجيل الدخول (Login Page)
# -------------------------------------------------------------
if not st.session_state.logged_in:
    st.title("🔐 تسجيل الدخول للمنصة")
    st.write("مرحباً بك، يرجى إدخال معطيات الدخول للوصول إلى الملفات.")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        username = st.text_input("اسم المستخدم (Username)")
        password = st.text_input("كلمة السر (Password)", type="password")
        
        if st.button("دخول", use_container_width=True):
            # يمكنك تعديل معطيات الدخول هنا
            if username == "admin" and password == "1234":
                st.session_state.logged_in = True
                st.success("تم تسجيل الدخول بنجاح!")
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة السر غير صحيحة")

# -------------------------------------------------------------
# 5. الواجهة الرئيسية للمنصة بعد الدخول (Main Dashboard)
# -------------------------------------------------------------
else:
    # القائمة الجانبية (Sidebar)
    with st.sidebar:
        st.title("📋 القائمة الرئيسية")
        st.write("مرحباً بك فـ النظام")
        st.divider()
        if st.button("تسجيل الخروج", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    # المحتوى الرئيسي للصفحة
    st.title("📂 لوحة التحكم وإدارة الملفات")
    st.write("الواجهة دابا جاهزة ونقية 100% للمستخدمين.")
