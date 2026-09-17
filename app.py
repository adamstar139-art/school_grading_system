import streamlit as st
import pandas as pd
import sqlite3
import openpyxl
import urllib.parse
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import streamlit.components.v1 as components
import requests

### ==============================================================================
### 1. تهيئة قاعدة البيانات والبيانات المعتمدة للطلاب
### ==============================================================================
DB_NAME = "school_grading_system.db"
STUDENTS_INIT = [
    # الصف الأول المتوسط - فصل 1
    {"id": "1167628468", "name": "ابراهيم بن محمد بن علي الوهيبي", "grade": "الأول المتوسط", "class": 1, "phone": "966504158122"},
    {"id": "2395664317", "name": "بلال عبدالرزاق عيسى العيسى", "grade": "الأول المتوسط", "class": 1, "phone": "966507448712"},
    {"id": "1170582165", "name": "حسام بن محمد بن علي ال رايان البارقي", "grade": "الأول المتوسط", "class": 1, "phone": "966504445699"},
    {"id": "1169004353", "name": "ريان عبدالله جابر الاسمري", "grade": "الأول المتوسط", "class": 1, "phone": "966554260960"},
    {"id": "2446713998", "name": "زيد زياد عبد اللطيف ابو قبع", "grade": "الأول المتوسط", "class": 1, "phone": "966590123455"},
    {"id": "2527104554", "name": "سامي سعد عباس حمد", "grade": "الأول المتوسط", "class": 1, "phone": "966591781701"},
    {"id": "1170111759", "name": "سعد ناصر سعد السيف", "grade": "الأول المتوسط", "class": 1, "phone": "966503219351"},
    {"id": "1195559479", "name": "عبدالعزيز عبدالله عبدالعزيز العمار", "grade": "الأول المتوسط", "class": 1, "phone": "966555838394"},
    {"id": "1153310501", "name": "عبدالله بن سليمان بن عبدالله الراجحي", "grade": "الأول المتوسط", "class": 1, "phone": "0551418881"},
    {"id": "1170836520", "name": "عبدالله سعد بن محمد العيشان", "grade": "الأول المتوسط", "class": 1, "phone": "966504217660"},
    {"id": "1171448515", "name": "علي احمد علي كريري", "grade": "الأول المتوسط", "class": 1, "phone": "966558885481"},
    {"id": "1170853053", "name": "علي سعد علي القحطاني", "grade": "الأول المتوسط", "class": 1, "phone": "966505466546"},
    {"id": "1172018036", "name": "عمر عبدالله سعد الجبرين", "grade": "الأول المتوسط", "class": 1, "phone": "966555249420"},
    {"id": "2552851368", "name": "مازن اسلام احمد ابراهيم موسى", "grade": "الأول المتوسط", "class": 1, "phone": "966550490495"},
    {"id": "013609321", "name": "محمد أحمد علي عقيل", "grade": "الأول المتوسط", "class": 1, "phone": "966546000184"},
    {"id": "2394606749", "name": "محمد اشرف مسعود ابوخاطر", "grade": "الأول المتوسط", "class": 1, "phone": "966501276888"},
    {"id": "1170042046", "name": "محمد بن فيصل بن مصلح الشمراني", "grade": "الأول المتوسط", "class": 1, "phone": "966555832145"},
    {"id": "1169174164", "name": "محمد نايف فراج الدعجاني", "grade": "الأول المتوسط", "class": 1, "phone": "966554444782"},
    {"id": "2380890976", "name": "وائل بولعيش", "grade": "الأول المتوسط", "class": 1, "phone": "966591534495"},
    
    # الصف الأول المتوسط - فصل 2
    {"id": "1170348286", "name": "الوليد ابن خالد بن فهد العتيبي", "grade": "الأول المتوسط", "class": 2, "phone": "966558522229"},
    {"id": "1172433185", "name": "باسل محمد فرج الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966537589781"},
    {"id": "1173391556", "name": "بسام بن عبدالكريم بن عبدالله الحرقان الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966534467820"},
    {"id": "1169185053", "name": "تركي عبدالله مسفر الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966505258369"},
    {"id": "1170108078", "name": "تميم فهد عبدالعزيز العزاز", "grade": "الأول المتوسط", "class": 2, "phone": "966554435692"},
    {"id": "1170970741", "name": "جاسر بن عبدالله بن منصور المطاطحة الحارثي", "grade": "الأول المتوسط", "class": 2, "phone": "966559455545"},
    {"id": "1168982427", "name": "راكان عبدالله يحي كريري", "grade": "الأول المتوسط", "class": 2, "phone": "966581727444"},
    {"id": "1172590968", "name": "ريان عبدالله منصور السبر", "grade": "الأول المتوسط", "class": 2, "phone": "966566959670"},
    {"id": "2392863888", "name": "ريان وليد حلاق", "grade": "الأول المتوسط", "class": 2, "phone": "966530527662"},
    {"id": "1170420473", "name": "سيف عبدالكريم بريك العصيمي", "grade": "الأول المتوسط", "class": 2, "phone": "966504277904"},
    {"id": "1168942108", "name": "صالح حسن فتحي سندي", "grade": "الأول المتوسط", "class": 2, "phone": "966557553922"},
    {"id": "1173182138", "name": "عبدالرحمن ابراهيم عبدالله الحضيف", "grade": "الأول المتوسط", "class": 2, "phone": "966558822674"},
    {"id": "1172448548", "name": "عبدالله صالح حمد الصفيان", "grade": "الأول المتوسط", "class": 2, "phone": "966558889978"},
    {"id": "1170000945", "name": "فهد ابن احمد بن فهد العثمان", "grade": "الأول المتوسط", "class": 2, "phone": "966504484898"},
    {"id": "1167092616", "name": "فهد عويض ثعيل المطيري", "grade": "الأول المتوسط", "class": 2, "phone": "966508271056"},
    {"id": "1170413171", "name": "فهد نايف فهد الحسينان", "grade": "الأول المتوسط", "class": 2, "phone": "966504140616"},
    {"id": "1170294118", "name": "فيصل موينع عبدالله بن موينع", "grade": "الأول المتوسط", "class": 2, "phone": "966541600918"},
    {"id": "1171524604", "name": "فيصل ناصر سيف العريفي", "grade": "الأول المتوسط", "class": 2, "phone": "966505474606"},
    {"id": "2502333707", "name": "محمد اسلام محمد دراز", "grade": "الأول المتوسط", "class": 2, "phone": "966556124553"},
    {"id": "1170374993", "name": "مشاري عثمان سعد ناصر السعد", "grade": "الأول المتوسط", "class": 2, "phone": "966500330693"},
    {"id": "1170884165", "name": "يزن محمد علي اليحيا", "grade": "الأول المتوسط", "class": 2, "phone": "966557072133"},
    {"id": "1170548737", "name": "يوسف محمد عبدالله الدوسري", "grade": "الأول المتوسط", "class": 2, "phone": "966556666176"}
]

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            grade TEXT NOT NULL,
            class INT NOT NULL,
            phone TEXT NOT NULL
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS grades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT NOT NULL,
            term TEXT NOT NULL,
            week TEXT NOT NULL,
            score REAL,
            is_absent INT DEFAULT 0,
            UNIQUE(student_id, term, week),
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
    ''')
    c.execute("SELECT COUNT(*) FROM students")
    count = c.fetchone()[0]
    if count == 0:
        for st_data in STUDENTS_INIT:
            c.execute("INSERT OR REPLACE INTO students (id, name, grade, class, phone) VALUES (?, ?, ?, ?, ?)",
                      (st_data["id"], st_data["name"], st_data["grade"], st_data["class"], st_data["phone"]))
        conn.commit()
    conn.close()

init_db()

### ==============================================================================
### 2. الدوال المساعدة وصياغة الرسائل والربط مع Mora SMS API
### ==============================================================================
def create_whatsapp_url(phone, text):
    phone_clean = str(phone).strip().replace("+", "").replace(" ", "").replace("-", "")
    if phone_clean.startswith("05"):
        phone_clean = "966" + phone_clean[1:]
    elif phone_clean.startswith("5"):
        phone_clean = "966" + phone_clean
    encoded_text = urllib.parse.quote(text)
    return f"https://api.whatsapp.com/send?phone={phone_clean}&text={encoded_text}"

def send_mora_sms(phone, message, username="0560229124", password="@THA0508634881", sender_name="AlThaghr", api_url="https://mora-sa.com/api/v1/sendsms"):
    """
    دالة الإرسال لـ Mora SMS المحدثة مع طباعة تفاصيل استجابة الخادم لسهولة التشخيص
    """
    phone_clean = str(phone).strip().replace("+", "").replace(" ", "").replace("-", "")
    if phone_clean.startswith("05"):
        phone_clean = "966" + phone_clean[1:]
    elif phone_clean.startswith("5"):
        phone_clean = "966" + phone_clean
        
    payload = {
        "username": username,
        "password": password,
        "sender": sender_name,
        "numbers": phone_clean,
        "message": message,
        "unicode": "E"
    }
    
    headers = {"Content-Type": "application/x-www-form-urlencoded; charset=utf-8"}
    
    try:
        # المحاولة عبر POST أولاً
        res = requests.post(api_url, data=payload, headers=headers, timeout=10)
        
        # إذا كانت الاستجابة 200 وتتضمن مؤشرات النجاح
        if res.status_code == 200 and ("100" in res.text or "success" in res.text.lower() or "sent" in res.text.lower()):
            return True, f"✅ تم الإرسال بنجاح (الاستجابة: {res.text})"
        else:
            # تجربة المحاولة عبر GET إذا فشل POST
            params_str = urllib.parse.urlencode(payload)
            get_res = requests.get(f"{api_url}?{params_str}", timeout=10)
            if get_res.status_code == 200 and ("100" in get_res.text or "success" in get_res.text.lower()):
                return True, f"✅ تم الإرسال بنجاح عبر GET (الاستجابة: {get_res.text})"
            else:
                return False, f"⚠️ تفاصيل الاستجابة من مورا: POST [{res.status_code}] -> {res.text} | GET [{get_res.status_code}] -> {get_res.text}"
    except Exception as e:
        return False, f"❌ خطأ في الاتصال بالسيرفر: {str(e)}"

def generate_parent_message(student_name, score, is_absent):
    if is_absent:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نود التنبيه على غياب الطالب هذا الأسبوع، "
            f"ونحثكم على متابعة الانتظام وحضور الاختبارات لتجنب حسم الدرجات والتأثير على مستواه التحصيلي: متوسطة الثغر النموذجية الأهلية."
        )
    sc_str = f"{score}%" if score is not None else "أقل من 50%"
    if score is None or score < 50:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نفيدكم بأن نسبة إتقان الطالب هذا الأسبوع هي ({sc_str}). "
            f"حرصاً منا على مصلحة ابنكم ومستقبله الدراسي، نود إشعاركم بوجود تراجع ملحوظ في مستواه التحصيلي مؤخراً، "
            f"ونرجو منكم تكثيف المتابعة المنزلية والتواصل معنا للوقوف على أسباب هذا التراجع ووضع خطة لتحسين أدائه. مع تحياتنا متوسطة الثغر النموذجية الأهلية."
        )
    elif score <= 75:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نفيدكم بأن نسبة إتقان الطالب هذا الأسبوع هي ({sc_str}). "
            f"نود إحاطتكم علماً بأن المستوى التحصيلي لابنكم جيد ومستقر بشكل عام، ولكنه يمتلك قدرات أعلى تؤهله لتحقيق درجات أفضل. "
            f"نأمل منكم التركيز معه في الفترة القادمة لرفع كفاءته الدراسية. شاكرين لكم تعاونكم الدائم. مع تحياتنا متوسطة الثغر النموذجية الأهلية."
        )
    else:
        return (
            f"المكرم ولي أمر الطالب/ {student_name}، نتقدم بخالص الشكر والتقدير لكم وللطالب على الاهتمام والتفوق بنسبة إتقان ممتازة ({sc_str})، "
            f"يسعدنا إبلاغكم بأن ابنكم قدم أداءً تحصيلياً متميزاً وسلوكاً رائعاً داخل الفصل، وحصل على درجات ممتازة في التقييمات الأخيرة. "
            f"نشكر لكم حسن المتابعة والاهتمام، ونرجو الاستمرار في هذا الدعم المتبادل للحفاظ على هذا المستوى المتفوق. مع تحياتنا متوسطة الثغر النموذجية الأهلية."
        )

def render_printable_html_view(html_content, title="طباعة التقرير"):
    full_html = f'''<!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #333; }}
            .header-table {{ width: 100%; border-bottom: 2px solid #1e3c72; margin-bottom: 20px; padding-bottom: 10px; }}
            .data-table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            .data-table th {{ background-color: #1e3c72; color: white; padding: 10px; border: 1px solid #ccc; text-align: center; }}
            .data-table td {{ padding: 8px; border: 1px solid #ccc; text-align: center; }}
            .badge-red {{ color: #dc2626; font-weight: bold; }}
            .badge-blue {{ color: #2563eb; font-weight: bold; }}
            .badge-green {{ color: #16a34a; font-weight: bold; }}
            .badge-gray {{ color: #4b5563; font-weight: bold; }}
            .signatures {{ width: 100%; margin-top: 40px; text-align: center; font-weight: bold; }}
        </style>
    </head>
    <body>
        {html_content}
    </body>
    </html>'''
    st.components.v1.html(full_html, height=500, scrolling=True)

### ==============================================================================
### 3. واجهة برنامج Streamlit
### ==============================================================================
st.set_page_config(
    page_title="برنامج رصد الدرجات - متوسطة الثغر النموذجية الأهلية",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.sidebar.title("📌 القائمة الرئيسية")
page = st.sidebar.radio("اختر الصفحة:", ["📝 صفحة الرصد", "🏫 إدارة المدرسة وتقارير أولياء الأمور"])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ إعدادات Mora SMS")
mora_user = st.sidebar.text_input("اسم المستخدم في مورا:", value="966560229124")
mora_pass = st.sidebar.text_input("كلمة المرور:", value="THA@0508634881", type="password")
mora_sender = st.sidebar.text_input("اسم المرسل المعتمد:", value="S")
mora_endpoint = st.sidebar.text_input("رابط API مورا:", value="https://mora-sa.com/api/v1/sendsms")

### ==============================================================================
### الصفحة الأولى: صفحة الرصد (RECORDING SHEET)
### ==============================================================================
if page == "📝 صفحة الرصد":
    st.subheader("📝 صفحة رصد درجات الإتقان الأسبوعية")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        term = st.selectbox("الفصل الدراسي:", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني"])
    with col2:
        grade = st.selectbox("الصف الدراسي:", ["الأول المتوسط", "الثاني المتوسط", "الثالث المتوسط"])
    with col3:
        class_num = st.selectbox("الفصل / الشعبة:", [1, 2, 3])
    with col4:
        weeks = [f"الأسبوع {i}" for i in range(1, 19)]
        week = st.selectbox("الأسبوع المستهدف:", weeks)

    st.markdown("---")

    conn = get_db_connection()
    query = """
        SELECT s.id, s.name, s.grade, s.class, g.score, g.is_absent 
        FROM students s
        LEFT JOIN grades g ON s.id = g.student_id AND g.term = ? AND g.week = ?
        WHERE s.grade = ? AND s.class = ?
        ORDER BY s.name ASC
    """
    df_students = pd.read_sql_query(query, conn, params=(term, week, grade, class_num))
    conn.close()

    if df_students.empty:
        st.warning("لا يوجد طلاب مسجلين في هذا الصف والشعبة.")
    else:
        st.info(f"📊 عدد الطلاب في {grade} - فصل ({class_num}): **{len(df_students)} طالب** | {term} - {week}")
        
        with st.form("recording_form"):
            st.markdown("##### 📥 أدخل/عدّل درجات الطلاب وحالة الغياب:")
            
            updated_data = []
            for idx, row in df_students.iterrows():
                col_name, col_score, col_absent = st.columns([3, 2, 1])
                with col_name:
                    st.markdown(f'<div style="padding: 8px;">📌 {idx+1}. {row["name"]} ({row["id"]})</div>', unsafe_allow_html=True)
                with col_score:
                    curr_score = float(row["score"]) if (pd.notnull(row["score"]) and row["score"] is not None) else 0.0
                    sc = st.number_input(f"الدرجة (100)", min_value=0.0, max_value=100.0, value=curr_score, step=1.0, key=f"sc_{row['id']}")
                with col_absent:
                    curr_abs = bool(row["is_absent"]) if pd.notnull(row["is_absent"]) else False
                    is_abs = st.checkbox("غائب ⚪", value=curr_abs, key=f"abs_{row['id']}")
                
                final_score = 0.0 if is_abs else sc
                updated_data.append({
                    "student_id": row["id"],
                    "name": row["name"],
                    "score": final_score,
                    "is_absent": 1 if is_abs else 0
                })
            
            save_btn = st.form_submit_button("💾 حفظ البيانات والتحديث")
            
        if save_btn:
            conn = get_db_connection()
            c = conn.cursor()
            for item in updated_data:
                c.execute("""
                    INSERT INTO grades (student_id, term, week, score, is_absent)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(student_id, term, week) DO UPDATE SET
                    score = excluded.score,
                    is_absent = excluded.is_absent
                """, (item["student_id"], term, week, item["score"], item["is_absent"]))
            conn.commit()
            conn.close()
            st.success("✅ تم حفظ البيانات وتحديث جدول الرصد وقاعدة البيانات بنجاح!")
            st.rerun()

### ==============================================================================
### الصفحة الثانية: إدارة المدرسة وتقارير أولياء الأمور (ADMIN & PARENT REPORTS)
### ==============================================================================
elif page == "🏫 إدارة المدرسة وتقارير أولياء الأمور":
    st.subheader("🏫 إدارة المدرسة وإرسال وتقارير أولياء الأمور")
    
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        weeks = [f"الأسبوع {i}" for i in range(1, 19)]
        selected_week = st.selectbox("اختر الأسبوع لعرض التقرير والرسائل:", weeks)
    with col_w2:
        selected_term = st.selectbox("اختر الفصل الدراسي:", ["الفصل الدراسي الأول", "الفصل الدراسي الثاني"])
        
    st.markdown("---")

    conn = get_db_connection()
    query = """
        SELECT s.id, s.name, s.grade, s.class, s.phone, g.score, g.is_absent
        FROM students s
        LEFT JOIN grades g ON s.id = g.student_id AND g.term = ? AND g.week = ?
        ORDER BY s.grade, s.class, s.name
    """
    df_reports = pd.read_sql_query(query, conn, params=(selected_term, selected_week))
    conn.close()

    cat_red, cat_blue, cat_green, cat_gray = [], [], [], []

    for idx, r in df_reports.iterrows():
        sc = r["score"]
        is_abs = r["is_absent"]
        msg = generate_parent_message(r["name"], sc, is_abs)
        
        row_dict = {
            "id": r["id"],
            "name": r["name"],
            "grade": r["grade"],
            "class": r["class"],
            "phone": r["phone"],
            "score": sc if (sc is not None and is_abs == 0) else 0.0,
            "is_absent": is_abs,
            "message": msg
        }
        
        if is_abs == 1:
            cat_gray.append(row_dict)
        elif sc is None or sc < 50:
            cat_red.append(row_dict)
        elif sc <= 75:
            cat_blue.append(row_dict)
        else:
            cat_green.append(row_dict)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🔴 فئة أقل من 50%", f"{len(cat_red)} طالب")
    m2.metric("🔵 فئة 50% - 75%", f"{len(cat_blue)} طالب")
    m3.metric("🟢 فئة 76% - 100%", f"{len(cat_green)} طالب")
    m4.metric("⚪ فئة الغياب", f"{len(cat_gray)} طالب")

    st.markdown("---")

    st.markdown("### 📲 إرسال الرسائل لولي الأمر (WhatsApp + Mora SMS):")

    tab1, tab2, tab3, tab4 = st.tabs([
        f"🔴 أقل من 50% ({len(cat_red)})",
        f"🔵 50% - 75% ({len(cat_blue)})",
        f"🟢 76% - 100% ({len(cat_green)})",
        f"⚪ الغياب ({len(cat_gray)})"
    ])

    def show_category_tab(cat_list, cat_name):
        if not cat_list:
            st.info(f"لا يوجد طلاب في {cat_name} بهذا الأسبوع.")
        else:
            col_b1, col_b2 = st.columns([2, 1])
            with col_b1:
                st.markdown(f"##### 📲 قائمة رسائل {cat_name}:")
            with col_b2:
                if st.button(f"🚀 إرسال SMS جماعي لكل طلاب {cat_name}", key=f"bulk_{cat_name}"):
                    success_count = 0
                    fail_count = 0
                    details_log = []
                    with st.spinner(f"جاري إرسال SMS لجميع طلاب {cat_name}..."):
                        for item in cat_list:
                            ok, resp_msg = send_mora_sms(
                                item['phone'], item['message'],
                                username=mora_user, password=mora_pass,
                                sender_name=mora_sender, api_url=mora_endpoint
                            )
                            if ok:
                                success_count += 1
                            else:
                                fail_count += 1
                            details_log.append(f"• {item['name']} ({item['phone']}): {resp_msg}")
                    
                    if success_count > 0:
                        st.success(f"🎉 اكتملت عملية الإرسال! النجاح: {success_count} | الفشل: {fail_count}")
                    else:
                        st.error(f"❌ اكتملت عملية الإرسال! النجاح: {success_count} | الفشل: {fail_count}")
                        
                    with st.expander("🔍 اضغط هنا لرؤية تفاصيل الاستجابة والتشخيص الدقيق من خادم مورا:"):
                        for log_line in details_log:
                            st.write(log_line)

            for item in cat_list:
                wa_link = create_whatsapp_url(item['phone'], item['message'])
                with st.expander(f"👤 {item['name']} ({item['grade']} - فصل {item['class']}) | جوال: {item['phone']}"):
                    st.write(f"**رقم الهوية:** {item['id']}")
                    st.write(f"**النسبة المئوية / الدرجة:** {item['score']}%" if item['is_absent'] == 0 else "**الحالة:** غائب ⚪")
                    st.info(f"💬 **نص الرسالة:**\n\n{item['message']}")
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(f'''
                        <a href="{wa_link}" target="_blank" style="text-decoration:none;">
                            <div style="background-color:#25D366; color:white; padding:10px; border-radius:6px; text-align:center; font-weight:bold;">
                                💬 إرسال عبر الواتساب (WhatsApp)
                            </div>
                        </a>
                        ''', unsafe_allow_html=True)
                    with c2:
                        if st.button(f"📱 إرسال SMS فردي عبر مورا", key=f"sms_indiv_{item['id']}"):
                            with st.spinner("جاري الإرسال..."):
                                ok, resp_msg = send_mora_sms(
                                    item['phone'], item['message'],
                                    username=mora_user, password=mora_pass,
                                    sender_name=mora_sender, api_url=mora_endpoint
                                )
                                if ok:
                                    st.success(f"✅ {resp_msg}")
                                else:
                                    st.error(f"❌ {resp_msg}")

    with tab1: show_category_tab(cat_red, "فئة أقل من 50%")
    with tab2: show_category_tab(cat_blue, "فئة 50% - 75%")
    with tab3: show_category_tab(cat_green, "فئة 76% - 100%")
    with tab4: show_category_tab(cat_gray, "فئة الغياب")
