import os
import sys
import subprocess

def main():
    # معرفة المسار الحالي
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    app_path = os.path.join(base_dir, 'app.py')
    
    # تشغيل Streamlit ومنعه من فتح نوافذ متكررة
    # Streamlit غادي يفتح النافذة غير مرة وحدة أوتوماتيكياً
    subprocess.run([
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.headless=false",
        "--browser.gatherUsageStats=false"
    ])

if __name__ == "__main__":
    main()