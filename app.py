
"""
Human Resource Development and Management
Classification Society Training Platform
Single-file Streamlit App + Excel Database

Features:
- Login ID and password
- Captcha security verification
- Admin-created users only
- Auto-generated User ID, Login ID and temporary password
- Role-based portal access
- Admin-only Excel database download and full database views
- Trainer can upload training content and generate MCQs
- Surveyor / Plan Appraiser can take MCQ test
- Auto certificate after passing
- Excel database auto-updates after each action

Run:
    pip install -r requirements.txt
    streamlit run app.py

Default Admin:
    Login ID: admin
    Password: admin123
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import hashlib
import random
import re
import secrets
import string
import uuid

import pandas as pd
import streamlit as st


APP_TITLE = "Human Resource Development and Management"
APP_SUBTITLE = "Classification Society Training, Competency and Certification Platform"
APP_VERSION = "2026.05.21-secure"
DB_FILE = "hrdm_training_database.xlsx"


# ============================================================
# MODULE 1 — BASIC UTILITIES
# ============================================================

def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_text() -> str:
    return date.today().strftime("%Y-%m-%d")


def new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:8].upper()}"


def safe_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value)


def is_valid_url(value: str) -> bool:
    value = (value or "").strip()
    if not value:
        return True
    return value.startswith(("http://", "https://"))


def password_hash(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def generate_password(length: int = 10) -> str:
    chars = string.ascii_letters + string.digits + "@#$"
    return "".join(secrets.choice(chars) for _ in range(length))


def make_login_id(name: str, users_df: pd.DataFrame) -> str:
    base = re.sub(r"[^a-z0-9]", "", name.lower().strip().replace(" ", "."))
    if not base:
        base = "user"
    username = base
    count = 1
    existing = set(users_df["Login_ID"].astype(str).str.lower()) if not users_df.empty and "Login_ID" in users_df.columns else set()
    while username.lower() in existing:
        count += 1
        username = f"{base}{count}"
    return username


def actor_value(actor: dict, key: str, default: str = "") -> str:
    if not isinstance(actor, dict):
        return default
    aliases = {
        "User_ID": ["User_ID", "Staff_ID", "ID"],
        "Name": ["Name", "Staff_Name", "User_Name"],
        "Role": ["Role"],
        "Department": ["Department"],
        "Email": ["Email"],
        "Login_ID": ["Login_ID", "Username"],
    }
    for possible_key in aliases.get(key, [key]):
        if possible_key in actor and safe_text(actor.get(possible_key)).strip():
            return safe_text(actor.get(possible_key))
    return default


def read_sheet(sheet_name: str) -> pd.DataFrame:
    if not Path(DB_FILE).exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(DB_FILE, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        return pd.DataFrame()


def write_sheet(sheet_name: str, df: pd.DataFrame) -> None:
    if Path(DB_FILE).exists():
        with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def append_row(sheet_name: str, row: dict) -> None:
    df = read_sheet(sheet_name)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    write_sheet(sheet_name, df)


def show_table(df: pd.DataFrame) -> None:
    st.dataframe(df.fillna(""), width="stretch", hide_index=True)


# ============================================================
# MODULE 2 — FRONTEND STYLE
# ============================================================

def apply_style() -> None:
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
        .main-header {
            padding: 1.35rem 1.5rem;
            border-radius: 24px;
            background: linear-gradient(135deg, #0f172a, #1e293b);
            color: white;
            margin-bottom: 1rem;
        }
        .main-header h1 {margin:0; font-size:2.1rem;}
        .main-header p {margin:0.4rem 0 0 0; color:#cbd5e1;}
        .info-card {
            padding: 1rem;
            border-radius: 18px;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 4px rgba(15,23,42,0.08);
            margin-bottom: 1rem;
        }
        .step-box {
            padding: 0.75rem 1rem;
            border-radius: 14px;
            background: #f1f5f9;
            border-left: 4px solid #0f172a;
            margin-bottom: 0.45rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def app_header() -> None:
    st.markdown(
        f"""
        <div class="main-header">
            <h1>{APP_TITLE}</h1>
            <p>{APP_SUBTITLE} | Version {APP_VERSION}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def info_card(title: str, text: str) -> None:
    st.markdown(
        f"""
        <div class="info-card">
            <h3 style="margin-top:0;">{title}</h3>
            <p style="margin-bottom:0;">{text}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_steps(steps: list[str]) -> None:
    for index, step in enumerate(steps, start=1):
        st.markdown(
            f'<div class="step-box"><b>{index}.</b> {step}</div>',
            unsafe_allow_html=True,
        )


# ============================================================
# MODULE 3 — DATABASE STRUCTURE
# ============================================================

SHEET_COLUMNS = {
    "Users": [
        "User_ID", "Name", "Role", "Department", "Email", "Login_ID",
        "Password_Hash", "Status", "Created_By", "Created_On",
        "Last_Login", "Failed_Attempts",
    ],
    "Trainings": [
        "Training_ID", "Training_Title", "Category", "Target_Role",
        "Trainer_ID", "Trainer_Name", "Slides_Link", "Video_Link",
        "Reference_Link", "Schedule_Date", "Schedule_Time", "Meeting_Link",
        "Recording_Link", "Passing_Marks", "Status", "Created_By",
        "Created_On", "Last_Updated",
    ],
    "Training_Content": [
        "Content_ID", "Training_ID", "File_Name", "File_Type", "Content_Text",
        "Uploaded_By", "Uploaded_On",
    ],
    "Question_Bank": [
        "Question_ID", "Training_ID", "Question", "Option_A", "Option_B",
        "Option_C", "Option_D", "Correct_Answer", "Marks", "Generated_On",
    ],
    "Training_Records": [
        "Record_ID", "User_ID", "Name", "Role", "Training_ID",
        "Training_Title", "Status", "Slides_Opened", "Video_Opened",
        "Live_Attendance", "Recording_Opened", "Test_Status", "Score",
        "Passing_Marks", "Certificate_Status", "Certificate_Link",
        "Due_Date", "Completed_On", "Progress_%", "Remarks", "Last_Updated",
    ],
    "Notifications": [
        "Notification_ID", "Training_ID", "User_ID", "Name", "Email",
        "Subject", "Message", "Status", "Generated_On", "Sent_On",
        "Generated_By",
    ],
    "Certificates": [
        "Certificate_ID", "User_ID", "Name", "Role", "Training_ID",
        "Training_Title", "Score", "Issued_On", "Certificate_Link", "Status",
        "Issued_By",
    ],
    "Activity_Log": [
        "Log_ID", "Date_Time", "Activity", "Actor_ID", "Actor_Name",
        "Actor_Role", "User_ID", "Training_ID", "Status", "Remarks",
    ],
    "Dashboard": ["Metric", "Value"],
    "System": ["Key", "Value"],
}


def database_is_valid() -> bool:
    if not Path(DB_FILE).exists():
        return False
    users = read_sheet("Users")
    if users.empty:
        return False
    return set(SHEET_COLUMNS["Users"]).issubset(set(users.columns))


def create_database(reset: bool = False) -> None:
    if Path(DB_FILE).exists() and not reset and database_is_valid():
        return

    users = pd.DataFrame(
        [
            [
                "USR-ADMIN", "Admin User", "Admin", "Administration",
                "admin@classsociety.org", "admin", password_hash("admin123"),
                "Active", "System", today_text(), "", 0,
            ],
            [
                "USR-MGMT", "Management User", "Management", "Management",
                "management@classsociety.org", "management", password_hash("mgmt123"),
                "Active", "System", today_text(), "", 0,
            ],
            [
                "USR-TRN-001", "Training Officer", "Trainer", "Training Department",
                "trainer@classsociety.org", "trainer", password_hash("trainer123"),
                "Active", "System", today_text(), "", 0,
            ],
            [
                "USR-SUR-001", "Muhammad Ali", "Surveyor", "Electrical Survey",
                "ali@classsociety.org", "surveyor", password_hash("surveyor123"),
                "Active", "System", today_text(), "", 0,
            ],
            [
                "USR-APP-001", "Ahmed Khan", "Plan Appraiser", "Plan Appraisal",
                "ahmed@classsociety.org", "appraiser", password_hash("appraiser123"),
                "Active", "System", today_text(), "", 0,
            ],
        ],
        columns=SHEET_COLUMNS["Users"],
    )

    system = pd.DataFrame(
        [["APP_TITLE", APP_TITLE], ["APP_VERSION", APP_VERSION]],
        columns=SHEET_COLUMNS["System"],
    )

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        users.to_excel(writer, sheet_name="Users", index=False)
        system.to_excel(writer, sheet_name="System", index=False)
        for sheet, columns in SHEET_COLUMNS.items():
            if sheet not in ["Users", "System"]:
                pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet, index=False)

    log_activity("Database Created", actor_name="System", actor_role="System", remarks="Initial database created.", update=False)
    update_dashboard()


# ============================================================
# MODULE 4 — LOGGING, DASHBOARD AND PROGRESS
# ============================================================

def log_activity(
    activity: str,
    actor_id: str = "",
    actor_name: str = "",
    actor_role: str = "",
    user_id: str = "",
    training_id: str = "",
    status: str = "Success",
    remarks: str = "",
    update: bool = True,
) -> None:
    append_row(
        "Activity_Log",
        {
            "Log_ID": new_id("LOG"),
            "Date_Time": now_text(),
            "Activity": activity,
            "Actor_ID": actor_id,
            "Actor_Name": actor_name,
            "Actor_Role": actor_role,
            "User_ID": user_id,
            "Training_ID": training_id,
            "Status": status,
            "Remarks": remarks,
        },
    )
    if update:
        update_dashboard()


def calculate_progress(record: pd.Series) -> int:
    steps = [
        record.get("Slides_Opened") == "Yes",
        record.get("Video_Opened") == "Yes" or record.get("Recording_Opened") == "Yes",
        record.get("Live_Attendance") in ["Present", "Recording Viewed"],
        record.get("Test_Status") == "Passed",
        record.get("Certificate_Status") == "Issued",
    ]
    return int((sum(steps) / len(steps)) * 100)


def refresh_training_records() -> None:
    records = read_sheet("Training_Records")
    if records.empty:
        update_dashboard()
        return

    for idx, row in records.iterrows():
        progress = calculate_progress(row)
        records.at[idx, "Progress_%"] = progress
        records.at[idx, "Status"] = "Completed" if progress == 100 else "Pending"
        if progress == 100 and str(row.get("Completed_On", "")).strip() in ["", "nan", "NaT"]:
            records.at[idx, "Completed_On"] = today_text()
        records.at[idx, "Last_Updated"] = now_text()

    write_sheet("Training_Records", records)
    update_dashboard()


def update_dashboard() -> None:
    users = read_sheet("Users")
    trainings = read_sheet("Trainings")
    records = read_sheet("Training_Records")
    certs = read_sheet("Certificates")
    questions = read_sheet("Question_Bank")
    logs = read_sheet("Activity_Log")

    completed = len(records[records["Status"] == "Completed"]) if not records.empty else 0
    pending = len(records[records["Status"] != "Completed"]) if not records.empty else 0
    average_progress = round(records["Progress_%"].fillna(0).mean(), 2) if not records.empty else 0

    dashboard = pd.DataFrame(
        [
            ["Total Users", len(users)],
            ["Surveyors", len(users[users["Role"] == "Surveyor"]) if not users.empty else 0],
            ["Plan Appraisers", len(users[users["Role"] == "Plan Appraiser"]) if not users.empty else 0],
            ["Trainers", len(users[users["Role"] == "Trainer"]) if not users.empty else 0],
            ["Total Trainings", len(trainings)],
            ["Generated MCQs", len(questions)],
            ["Completed Records", completed],
            ["Pending Records", pending],
            ["Average Progress %", average_progress],
            ["Certificates Issued", len(certs)],
            ["Activity Logs", len(logs)],
            ["Last Updated", now_text()],
        ],
        columns=SHEET_COLUMNS["Dashboard"],
    )

    write_sheet("Dashboard", dashboard)


# ============================================================
# MODULE 5 — AUTHENTICATION AND CAPTCHA
# ============================================================

def generate_captcha() -> None:
    a = random.randint(2, 15)
    b = random.randint(2, 15)
    st.session_state["captcha_question"] = f"{a} + {b}"
    st.session_state["captcha_answer"] = str(a + b)


def normalize_session_user() -> None:
    user = st.session_state.get("user", {})
    if not isinstance(user, dict):
        st.session_state["logged_in"] = False
        st.session_state["user"] = {}
        return
    required = ["User_ID", "Name", "Role", "Email"]
    if st.session_state.get("logged_in", False) and not all(k in user for k in required):
        st.session_state["logged_in"] = False
        st.session_state["user"] = {}


def authenticate(login_id: str, password: str) -> dict | None:
    users = read_sheet("Users")
    if users.empty:
        return None

    login_id = login_id.strip()
    entered_hash = password_hash(password.strip())

    match = users[
        (users["Login_ID"].astype(str) == login_id)
        & (users["Password_Hash"].astype(str) == entered_hash)
        & (users["Status"].astype(str) == "Active")
    ]

    if match.empty:
        return None

    idx = match.index[0]
    users.at[idx, "Last_Login"] = now_text()
    users.at[idx, "Failed_Attempts"] = 0
    write_sheet("Users", users)

    return users.loc[idx].to_dict()


def register_failed_login(login_id: str) -> None:
    users = read_sheet("Users")
    if users.empty or "Login_ID" not in users.columns:
        return

    match = users[users["Login_ID"].astype(str) == login_id.strip()]
    if match.empty:
        return

    idx = match.index[0]
    try:
        failed = int(users.at[idx, "Failed_Attempts"])
    except Exception:
        failed = 0

    failed += 1
    users.at[idx, "Failed_Attempts"] = failed

    if failed >= 5:
        users.at[idx, "Status"] = "Inactive"

    write_sheet("Users", users)


def login_page() -> None:
    st.subheader("Secure Login")
    st.caption("Only users created by Admin can enter the system.")

    if "captcha_question" not in st.session_state or "captcha_answer" not in st.session_state:
        generate_captcha()

    with st.form("login_form"):
        login_id = st.text_input("Login ID")
        password = st.text_input("Password", type="password")
        captcha_input = st.text_input(f"Security Verification: {st.session_state['captcha_question']} = ?")
        submitted = st.form_submit_button("Login")

    if submitted:
        if captcha_input.strip() != st.session_state["captcha_answer"]:
            st.error("Security verification failed. Please try again.")
            generate_captcha()
            st.rerun()

        user = authenticate(login_id, password)
        if user is None:
            register_failed_login(login_id)
            st.error("Invalid Login ID/password, inactive user, or too many failed attempts.")
            generate_captcha()
        else:
            st.session_state["logged_in"] = True
            st.session_state["user"] = user
            log_activity(
                "User Login",
                actor_id=user["User_ID"],
                actor_name=user["Name"],
                actor_role=user["Role"],
                user_id=user["User_ID"],
                remarks="Successful login.",
            )
            st.success("Login successful.")
            st.rerun()

    with st.expander("Default login for first testing"):
        st.write("Login ID: `admin`")
        st.write("Password: `admin123`")


def require_login() -> dict:
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "user" not in st.session_state:
        st.session_state["user"] = {}

    normalize_session_user()

    if not st.session_state["logged_in"]:
        login_page()
        st.stop()

    return st.session_state["user"]


# ============================================================
# MODULE 6 — BUSINESS LOGIC
# ============================================================

def default_meeting_link(training_id: str) -> str:
    return f"https://teams.microsoft.com/l/meetup-join/{training_id}"


def default_certificate_link(user_id: str, training_id: str) -> str:
    return f"https://certificate.classsociety.org/{user_id}/{training_id}"


def create_notifications(training_id: str, actor: dict) -> None:
    trainings = read_sheet("Trainings")
    records = read_sheet("Training_Records")
    users = read_sheet("Users")

    selected_training = trainings[trainings["Training_ID"] == training_id]
    if selected_training.empty:
        return

    training = selected_training.iloc[0]
    assigned_records = records[records["Training_ID"] == training_id]

    for _, record in assigned_records.iterrows():
        person_match = users[users["User_ID"] == record["User_ID"]]
        if person_match.empty:
            continue

        person = person_match.iloc[0]
        message = (
            f"Dear {person['Name']},\n\n"
            f"You have been assigned the following training:\n\n"
            f"Training: {training['Training_Title']}\n"
            f"Schedule: {training['Schedule_Date']} at {training['Schedule_Time']}\n"
            f"Meeting Link: {training['Meeting_Link']}\n"
            f"Slides Link: {training['Slides_Link']}\n"
            f"Video Link: {training['Video_Link']}\n\n"
            f"Please complete the training activities and submit the MCQ test.\n\n"
            f"Regards,\nHR Development and Management"
        )

        append_row(
            "Notifications",
            {
                "Notification_ID": new_id("NOTIF"),
                "Training_ID": training_id,
                "User_ID": person["User_ID"],
                "Name": person["Name"],
                "Email": person["Email"],
                "Subject": f"Training Assigned: {training['Training_Title']}",
                "Message": message,
                "Status": "Generated",
                "Generated_On": now_text(),
                "Sent_On": "",
                "Generated_By": actor_value(actor, "Name", "System"),
            },
        )

    log_activity(
        activity="Notifications Generated",
        actor_id=actor_value(actor, "User_ID"),
        actor_name=actor_value(actor, "Name"),
        actor_role=actor_value(actor, "Role"),
        training_id=training_id,
        remarks="Notification records generated in Excel.",
    )


def issue_certificate(actor: dict, training_id: str, score: int) -> None:
    user_id = actor_value(actor, "User_ID")
    name = actor_value(actor, "Name")
    role = actor_value(actor, "Role")

    records = read_sheet("Training_Records")
    certificates = read_sheet("Certificates")

    mask = (records["User_ID"] == user_id) & (records["Training_ID"] == training_id)
    if records[mask].empty:
        return

    record = records[mask].iloc[0]
    link = default_certificate_link(user_id, training_id)

    records.loc[mask, "Certificate_Status"] = "Issued"
    records.loc[mask, "Certificate_Link"] = link
    records.loc[mask, "Remarks"] = "Certificate issued after passing MCQ test."
    records.loc[mask, "Last_Updated"] = now_text()
    write_sheet("Training_Records", records)

    exists = certificates[
        (certificates["User_ID"] == user_id)
        & (certificates["Training_ID"] == training_id)
    ]

    if exists.empty:
        certificates = pd.concat(
            [
                certificates,
                pd.DataFrame(
                    [
                        {
                            "Certificate_ID": new_id("CERT"),
                            "User_ID": user_id,
                            "Name": name,
                            "Role": role,
                            "Training_ID": training_id,
                            "Training_Title": record["Training_Title"],
                            "Score": score,
                            "Issued_On": today_text(),
                            "Certificate_Link": link,
                            "Status": "Issued",
                            "Issued_By": "System",
                        }
                    ]
                ),
            ],
            ignore_index=True,
        )
        write_sheet("Certificates", certificates)

    log_activity(
        activity="Certificate Issued",
        actor_id=user_id,
        actor_name=name,
        actor_role=role,
        user_id=user_id,
        training_id=training_id,
        remarks=link,
    )
    refresh_training_records()


# ============================================================
# MODULE 7 — CONTENT EXTRACTION AND MCQ GENERATION
# ============================================================

def extract_docx_text_safely(uploaded_file) -> str:
    try:
        import importlib
        docx_module = importlib.import_module("docx")
        document = docx_module.Document(uploaded_file)
        paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except ModuleNotFoundError:
        st.error("DOCX support requires python-docx. Run: pip install python-docx")
        return ""
    except Exception as exc:
        st.error(f"Could not read Word file: {exc}")
        return ""


def extract_text_from_uploaded_file(uploaded_file) -> str:
    if uploaded_file is None:
        return ""

    name = uploaded_file.name.lower()

    if name.endswith(".txt"):
        return uploaded_file.getvalue().decode("utf-8", errors="ignore")

    if name.endswith(".docx"):
        return extract_docx_text_safely(uploaded_file)

    st.error("Only .txt and .docx files are supported.")
    return ""


def extract_keywords(text: str) -> list[str]:
    words = re.findall(r"\b[A-Za-z][A-Za-z\-]{4,}\b", text)
    stopwords = {
        "training", "system", "should", "shall", "which", "there", "their", "about",
        "through", "during", "after", "before", "within", "using", "based", "these",
        "those", "where", "under", "requirements", "procedure", "document",
        "classification", "society", "survey", "surveyor", "appraisal",
    }

    cleaned = []
    for word in words:
        w = word.strip(".,:;()[]{}").lower()
        if w not in stopwords and len(w) >= 5:
            cleaned.append(w.title())

    unique = []
    for word in cleaned:
        if word not in unique:
            unique.append(word)

    return unique[:80]


def generate_mcqs_from_text(training_id: str, text: str, max_questions: int = 10) -> pd.DataFrame:
    sentences = re.split(r"(?<=[.!?])\s+", text.replace("\n", " "))
    sentences = [s.strip() for s in sentences if 50 <= len(s.strip()) <= 240]

    keywords = extract_keywords(text)
    if len(keywords) < 4:
        return pd.DataFrame(columns=SHEET_COLUMNS["Question_Bank"])

    questions = []
    random.shuffle(sentences)

    for sentence in sentences:
        if len(questions) >= max_questions:
            break

        answer = None
        for keyword in keywords:
            if re.search(rf"\b{re.escape(keyword)}\b", sentence, re.IGNORECASE):
                answer = keyword
                break

        if not answer:
            continue

        distractors = [k for k in keywords if k.lower() != answer.lower()]
        if len(distractors) < 3:
            continue

        options = random.sample(distractors, 3) + [answer]
        random.shuffle(options)

        question_text = re.sub(
            rf"\b{re.escape(answer)}\b",
            "__________",
            sentence,
            flags=re.IGNORECASE,
            count=1,
        )

        questions.append(
            {
                "Question_ID": new_id("Q"),
                "Training_ID": training_id,
                "Question": question_text,
                "Option_A": options[0],
                "Option_B": options[1],
                "Option_C": options[2],
                "Option_D": options[3],
                "Correct_Answer": answer,
                "Marks": 1,
                "Generated_On": now_text(),
            }
        )

    return pd.DataFrame(questions, columns=SHEET_COLUMNS["Question_Bank"])


# ============================================================
# MODULE 8 — SIDEBAR AND COMMON UI
# ============================================================

def sidebar_common(actor: dict) -> None:
    name = actor_value(actor, "Name", "Unknown User")
    role = actor_value(actor, "Role", "Unknown Role")
    email = actor_value(actor, "Email", "")

    st.sidebar.success(f"{name} ({role})")
    st.sidebar.caption(email)
    st.sidebar.divider()

    if st.sidebar.button("Logout"):
        log_activity(
            "User Logout",
            actor_id=actor_value(actor, "User_ID"),
            actor_name=name,
            actor_role=role,
            user_id=actor_value(actor, "User_ID"),
            remarks="User logged out.",
        )
        st.session_state["logged_in"] = False
        st.session_state["user"] = {}
        generate_captcha()
        st.rerun()

    if role == "Admin":
        if st.sidebar.button("Reset Excel Database"):
            create_database(reset=True)
            st.session_state["logged_in"] = False
            st.session_state["user"] = {}
            st.sidebar.success("Database reset. Please login again.")
            st.rerun()

        if Path(DB_FILE).exists():
            with open(DB_FILE, "rb") as file:
                st.sidebar.download_button(
                    "Download Excel Database",
                    file,
                    file_name=DB_FILE,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# ============================================================
# MODULE 9 — DASHBOARD
# ============================================================

def dashboard_page(actor: dict) -> None:
    role = actor_value(actor, "Role", "Unknown Role")
    st.header(f"{role} Dashboard")

    if role == "Admin":
        steps = [
            "Create users with auto-generated User ID, Login ID and Password.",
            "Assign role to each user.",
            "Create training and assign trainer.",
            "Assign surveyors / plan appraisers to training.",
            "Download Excel database and monitor all records.",
        ]
    elif role == "Trainer":
        steps = [
            "Open assigned training.",
            "Add links and upload training content file.",
            "Generate MCQ test from uploaded content.",
            "Schedule training and generate notification records.",
            "Mark attendance and add recording link.",
        ]
    elif role in ["Surveyor", "Plan Appraiser"]:
        steps = [
            "Open assigned training.",
            "Mark slides/video/recording complete.",
            "Take MCQ test.",
            "Certificate is issued automatically if passed.",
        ]
    else:
        steps = [
            "View HRD progress.",
            "Review training records and certification status.",
            "Monitor training effectiveness.",
        ]

    info_card("Role Workflow", "Your dashboard shows only your permitted functions.")
    show_steps(steps)

    dashboard = read_sheet("Dashboard")
    if not dashboard.empty:
        st.subheader("System Summary")
        cols = st.columns(4)
        for i, row in dashboard.iterrows():
            cols[i % 4].metric(str(row["Metric"]), str(row["Value"]))


# ============================================================
# MODULE 10 — ADMIN PANEL
# ============================================================

def admin_panel(actor: dict) -> None:
    st.header("Admin Panel")
    info_card("Admin Control", "Admin creates users, assigns roles, creates trainings, assigns trainees, and views/downloads Excel data.")

    tab_users, tab_training, tab_assign, tab_database = st.tabs(
        ["Users & Roles", "Create Training", "Assign Trainees", "Admin Database"]
    )

    with tab_users:
        st.subheader("Current Users")
        show_table(read_sheet("Users").drop(columns=["Password_Hash"], errors="ignore"))

        with st.form("user_form"):
            st.markdown("### Create New User")
            name = st.text_input("Name")
            role = st.selectbox("Role", ["Admin", "Management", "Trainer", "Surveyor", "Plan Appraiser"])
            department = st.text_input("Department")
            email = st.text_input("Email")
            status = st.selectbox("Status", ["Active", "Inactive"])
            submitted = st.form_submit_button("Create User and Generate Login Credentials")

        if submitted:
            if not name or not email:
                st.error("Name and email are required.")
            else:
                users = read_sheet("Users")
                user_id = new_id("USR")
                login_id = make_login_id(name, users)
                temp_password = generate_password(10)

                row = {
                    "User_ID": user_id,
                    "Name": name,
                    "Role": role,
                    "Department": department,
                    "Email": email,
                    "Login_ID": login_id,
                    "Password_Hash": password_hash(temp_password),
                    "Status": status,
                    "Created_By": actor_value(actor, "Name"),
                    "Created_On": today_text(),
                    "Last_Login": "",
                    "Failed_Attempts": 0,
                }

                users = pd.concat([users, pd.DataFrame([row])], ignore_index=True)
                write_sheet("Users", users)
                log_activity("User Created", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), user_id=user_id)
                update_dashboard()

                st.success("User created successfully.")
                st.info("Share these credentials with the user securely. The password is shown only now.")
                st.code(f"User ID: {user_id}\nLogin ID: {login_id}\nTemporary Password: {temp_password}")

        st.subheader("Reset User Password / Change Status")
        users = read_sheet("Users")
        if not users.empty:
            selected_user = st.selectbox("Select User", users["Name"] + " — " + users["User_ID"])
            selected_user_id = selected_user.split(" — ")[-1]
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Reset Password"):
                    new_password = generate_password(10)
                    users = read_sheet("Users")
                    idx = users[users["User_ID"] == selected_user_id].index[0]
                    users.at[idx, "Password_Hash"] = password_hash(new_password)
                    users.at[idx, "Failed_Attempts"] = 0
                    users.at[idx, "Status"] = "Active"
                    write_sheet("Users", users)
                    log_activity("Password Reset", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), user_id=selected_user_id)
                    st.success("Password reset.")
                    st.code(f"New Temporary Password: {new_password}")
            with col2:
                new_status = st.selectbox("New Status", ["Active", "Inactive"])
                if st.button("Update Status"):
                    users = read_sheet("Users")
                    idx = users[users["User_ID"] == selected_user_id].index[0]
                    users.at[idx, "Status"] = new_status
                    write_sheet("Users", users)
                    log_activity("User Status Updated", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), user_id=selected_user_id, remarks=new_status)
                    st.success("Status updated.")
                    st.rerun()

    with tab_training:
        users = read_sheet("Users")
        trainers = users[(users["Role"] == "Trainer") & (users["Status"] == "Active")]

        if trainers.empty:
            st.warning("Create at least one active Trainer first.")
        else:
            with st.form("training_form"):
                st.markdown("### Create Training")
                title = st.text_input("Training Title")
                category = st.selectbox(
                    "Category",
                    ["Basic Survey", "Electrical Survey", "Hull Survey", "Machinery Survey", "Plan Appraisal", "Statutory Survey", "Report Writing"],
                )
                target_role = st.selectbox("Target Role", ["Surveyor", "Plan Appraiser"])
                trainer_display = st.selectbox("Trainer", trainers["Name"] + " — " + trainers["User_ID"])
                passing_marks = st.number_input("Passing Marks (%)", min_value=1, max_value=100, value=75)
                submitted_training = st.form_submit_button("Create Training")

            if submitted_training:
                if not title:
                    st.error("Training title is required.")
                else:
                    training_id = new_id("TRN")
                    trainer_name, trainer_id = trainer_display.split(" — ")

                    trainings = read_sheet("Trainings")
                    row = {
                        "Training_ID": training_id,
                        "Training_Title": title,
                        "Category": category,
                        "Target_Role": target_role,
                        "Trainer_ID": trainer_id,
                        "Trainer_Name": trainer_name,
                        "Slides_Link": "",
                        "Video_Link": "",
                        "Reference_Link": "",
                        "Schedule_Date": "",
                        "Schedule_Time": "",
                        "Meeting_Link": "",
                        "Recording_Link": "",
                        "Passing_Marks": int(passing_marks),
                        "Status": "Draft",
                        "Created_By": actor_value(actor, "Name"),
                        "Created_On": now_text(),
                        "Last_Updated": now_text(),
                    }

                    trainings = pd.concat([trainings, pd.DataFrame([row])], ignore_index=True)
                    write_sheet("Trainings", trainings)
                    log_activity("Training Created", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), training_id=training_id)
                    update_dashboard()
                    st.success(f"Training created. Training ID: {training_id}")
                    st.rerun()

    with tab_assign:
        users = read_sheet("Users")
        trainings = read_sheet("Trainings")

        if trainings.empty:
            st.warning("No training created yet.")
        else:
            training_display = st.selectbox("Select Training", trainings["Training_Title"] + " — " + trainings["Training_ID"])
            training_id = training_display.split(" — ")[-1]
            training = trainings[trainings["Training_ID"] == training_id].iloc[0]

            eligible = users[
                (users["Role"] == training["Target_Role"])
                & (users["Status"] == "Active")
            ]
            selected_trainees = st.multiselect("Select Trainees", eligible["Name"] + " — " + eligible["User_ID"])
            due_date = st.date_input("Due Date")

            if st.button("Assign Selected Trainees"):
                if not selected_trainees:
                    st.warning("Select at least one trainee.")
                else:
                    records = read_sheet("Training_Records")

                    for trainee in selected_trainees:
                        trainee_name, trainee_id = trainee.split(" — ")
                        trainee_row = users[users["User_ID"] == trainee_id].iloc[0]

                        exists = records[
                            (records["User_ID"] == trainee_id)
                            & (records["Training_ID"] == training_id)
                        ]
                        if not exists.empty:
                            continue

                        row = {
                            "Record_ID": new_id("REC"),
                            "User_ID": trainee_id,
                            "Name": trainee_name,
                            "Role": trainee_row["Role"],
                            "Training_ID": training_id,
                            "Training_Title": training["Training_Title"],
                            "Status": "Pending",
                            "Slides_Opened": "No",
                            "Video_Opened": "No",
                            "Live_Attendance": "Not Marked",
                            "Recording_Opened": "No",
                            "Test_Status": "Not Attempted",
                            "Score": "",
                            "Passing_Marks": int(training["Passing_Marks"]),
                            "Certificate_Status": "Not Issued",
                            "Certificate_Link": "",
                            "Due_Date": str(due_date),
                            "Completed_On": "",
                            "Progress_%": 0,
                            "Remarks": "Assigned by Admin.",
                            "Last_Updated": now_text(),
                        }

                        records = pd.concat([records, pd.DataFrame([row])], ignore_index=True)
                        log_activity(
                            "Training Assigned",
                            actor_value(actor, "User_ID"),
                            actor_value(actor, "Name"),
                            actor_value(actor, "Role"),
                            user_id=trainee_id,
                            training_id=training_id,
                        )

                    write_sheet("Training_Records", records)
                    refresh_training_records()
                    st.success("Trainees assigned.")
                    st.rerun()

    with tab_database:
        st.subheader("Admin Database View")
        st.warning("This section is visible only to Admin.")
        for title, sheet in [
            ("Dashboard", "Dashboard"),
            ("Users", "Users"),
            ("Trainings", "Trainings"),
            ("Training Records", "Training_Records"),
            ("Training Content", "Training_Content"),
            ("Question Bank", "Question_Bank"),
            ("Notifications", "Notifications"),
            ("Certificates", "Certificates"),
            ("Activity Log", "Activity_Log"),
        ]:
            st.markdown(f"### {title}")
            df = read_sheet(sheet)
            if sheet == "Users":
                df = df.drop(columns=["Password_Hash"], errors="ignore")
            show_table(df)


# ============================================================
# MODULE 11 — TRAINER PANEL
# ============================================================

def trainer_panel(actor: dict) -> None:
    st.header("Trainer Panel")
    info_card("Trainer Work", "Add links, upload training content, generate MCQs, schedule training and track attendance.")

    trainings = read_sheet("Trainings")
    my_trainings = trainings[trainings["Trainer_ID"] == actor_value(actor, "User_ID")]

    if my_trainings.empty:
        st.warning("No training assigned to this trainer.")
        return

    training_display = st.selectbox("Assigned Training", my_trainings["Training_Title"] + " — " + my_trainings["Training_ID"])
    training_id = training_display.split(" — ")[-1]
    training = my_trainings[my_trainings["Training_ID"] == training_id].iloc[0]

    tab_material, tab_content, tab_schedule, tab_attendance, tab_results = st.tabs(
        ["Links & Material", "Upload Content & Generate MCQs", "Schedule & Notify", "Attendance & Recording", "Results"]
    )

    with tab_material:
        with st.form("material_form"):
            slides = st.text_input("Slides Link", value=safe_text(training.get("Slides_Link", "")))
            video = st.text_input("Video Link", value=safe_text(training.get("Video_Link", "")))
            reference = st.text_input("Reference Link", value=safe_text(training.get("Reference_Link", "")))
            passing_marks = st.number_input("Passing Marks (%)", min_value=1, max_value=100, value=int(training.get("Passing_Marks", 75)))
            save_material = st.form_submit_button("Save Links & Material Settings")

        if save_material:
            if not is_valid_url(slides) or not is_valid_url(video) or not is_valid_url(reference):
                st.error("Links must start with http:// or https://")
            else:
                all_trainings = read_sheet("Trainings")
                idx = all_trainings[all_trainings["Training_ID"] == training_id].index[0]

                all_trainings.at[idx, "Slides_Link"] = slides.strip()
                all_trainings.at[idx, "Video_Link"] = video.strip()
                all_trainings.at[idx, "Reference_Link"] = reference.strip()
                all_trainings.at[idx, "Passing_Marks"] = int(passing_marks)
                all_trainings.at[idx, "Status"] = "Material Added"
                all_trainings.at[idx, "Last_Updated"] = now_text()

                write_sheet("Trainings", all_trainings)

                records = read_sheet("Training_Records")
                records.loc[records["Training_ID"] == training_id, "Passing_Marks"] = int(passing_marks)
                write_sheet("Training_Records", records)

                log_activity("Training Material Links Saved", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), training_id=training_id)
                update_dashboard()
                st.success("Links and settings saved.")
                st.rerun()

    with tab_content:
        st.subheader("Upload Training Content")
        st.caption("Upload .txt always works. .docx works only if python-docx is installed.")

        uploaded = st.file_uploader("Upload Training Content File", type=["txt", "docx"])
        mcq_count = st.slider("Number of MCQs to generate", min_value=5, max_value=20, value=10)

        if st.button("Extract Content and Generate MCQs"):
            content_text = extract_text_from_uploaded_file(uploaded)

            if not content_text.strip():
                st.error("No readable content found.")
            else:
                append_row(
                    "Training_Content",
                    {
                        "Content_ID": new_id("CONTENT"),
                        "Training_ID": training_id,
                        "File_Name": uploaded.name,
                        "File_Type": uploaded.name.split(".")[-1].lower(),
                        "Content_Text": content_text[:30000],
                        "Uploaded_By": actor_value(actor, "Name"),
                        "Uploaded_On": now_text(),
                    },
                )

                new_questions = generate_mcqs_from_text(training_id, content_text, max_questions=mcq_count)

                if new_questions.empty:
                    st.error("Could not generate MCQs. Please upload content with clear technical sentences.")
                else:
                    question_bank = read_sheet("Question_Bank")
                    question_bank = pd.concat([question_bank, new_questions], ignore_index=True)
                    write_sheet("Question_Bank", question_bank)

                    log_activity("MCQs Generated", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), training_id=training_id, remarks=f"{len(new_questions)} questions generated.")
                    update_dashboard()
                    st.success(f"{len(new_questions)} MCQs generated and saved.")
                    show_table(new_questions)

        st.subheader("Existing MCQs for This Training")
        qbank = read_sheet("Question_Bank")
        show_table(qbank[qbank["Training_ID"] == training_id] if not qbank.empty else qbank)

    with tab_schedule:
        schedule_date = st.date_input("Schedule Date")
        schedule_time = st.text_input("Schedule Time", value="10:00 AM")
        link = st.text_input("Meeting Link", value=default_meeting_link(training_id))

        if st.button("Schedule Training & Generate Notifications"):
            if not is_valid_url(link):
                st.error("Meeting link must start with http:// or https://")
            else:
                all_trainings = read_sheet("Trainings")
                idx = all_trainings[all_trainings["Training_ID"] == training_id].index[0]

                all_trainings.at[idx, "Schedule_Date"] = str(schedule_date)
                all_trainings.at[idx, "Schedule_Time"] = schedule_time
                all_trainings.at[idx, "Meeting_Link"] = link.strip()
                all_trainings.at[idx, "Status"] = "Scheduled"
                all_trainings.at[idx, "Last_Updated"] = now_text()

                write_sheet("Trainings", all_trainings)
                create_notifications(training_id, actor)
                update_dashboard()
                st.success("Training scheduled and notifications generated.")
                st.rerun()

        st.subheader("Notifications for This Training")
        notifications = read_sheet("Notifications")
        show_table(notifications[notifications["Training_ID"] == training_id] if not notifications.empty else notifications)

    with tab_attendance:
        records = read_sheet("Training_Records")
        trainees = records[records["Training_ID"] == training_id]

        if trainees.empty:
            st.warning("No trainees assigned yet.")
        else:
            show_table(trainees[["User_ID", "Name", "Live_Attendance", "Status", "Progress_%"]])
            trainee_display = st.selectbox("Select Trainee", trainees["Name"] + " — " + trainees["User_ID"])
            trainee_id = trainee_display.split(" — ")[-1]
            attendance = st.selectbox("Attendance", ["Present", "Absent"])

            if st.button("Mark Attendance"):
                records = read_sheet("Training_Records")
                mask = (records["Training_ID"] == training_id) & (records["User_ID"] == trainee_id)

                records.loc[mask, "Live_Attendance"] = attendance
                records.loc[mask, "Remarks"] = f"Attendance marked: {attendance}"
                records.loc[mask, "Last_Updated"] = now_text()

                write_sheet("Training_Records", records)
                log_activity("Attendance Marked", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), user_id=trainee_id, training_id=training_id, remarks=attendance)
                refresh_training_records()
                st.success("Attendance saved.")
                st.rerun()

        recording = st.text_input("Recording Link")
        if st.button("Save Recording Link"):
            if not is_valid_url(recording):
                st.error("Recording link must start with http:// or https://")
            else:
                all_trainings = read_sheet("Trainings")
                idx = all_trainings[all_trainings["Training_ID"] == training_id].index[0]

                all_trainings.at[idx, "Recording_Link"] = recording.strip()
                all_trainings.at[idx, "Status"] = "Recorded"
                all_trainings.at[idx, "Last_Updated"] = now_text()

                write_sheet("Trainings", all_trainings)
                log_activity("Recording Saved", actor_value(actor, "User_ID"), actor_value(actor, "Name"), actor_value(actor, "Role"), training_id=training_id)
                update_dashboard()
                st.success("Recording link saved.")
                st.rerun()

    with tab_results:
        records = read_sheet("Training_Records")
        show_table(records[records["Training_ID"] == training_id] if not records.empty else records)


# ============================================================
# MODULE 12 — TRAINEE PANEL
# ============================================================

def trainee_activity(actor: dict, training_id: str, field: str, remarks: str) -> None:
    user_id = actor_value(actor, "User_ID")
    records = read_sheet("Training_Records")
    mask = (records["User_ID"] == user_id) & (records["Training_ID"] == training_id)

    records.loc[mask, field] = "Yes"
    records.loc[mask, "Remarks"] = remarks
    records.loc[mask, "Last_Updated"] = now_text()

    write_sheet("Training_Records", records)
    log_activity(field.replace("_", " "), user_id, actor_value(actor, "Name"), actor_value(actor, "Role"), user_id, training_id, remarks=remarks)
    refresh_training_records()
    st.success("Activity updated.")
    st.rerun()


def trainee_panel(actor: dict) -> None:
    role = actor_value(actor, "Role")
    user_id = actor_value(actor, "User_ID")
    st.header(f"{role} Training Portal")
    info_card("Training Actions", "Complete assigned training material and take MCQ test.")

    records = read_sheet("Training_Records")
    my_records = records[records["User_ID"] == user_id]

    if my_records.empty:
        st.warning("No training assigned yet.")
        return

    training_display = st.selectbox("My Training", my_records["Training_Title"] + " — " + my_records["Training_ID"])
    training_id = training_display.split(" — ")[-1]

    record = my_records[my_records["Training_ID"] == training_id].iloc[0]
    trainings = read_sheet("Trainings")
    training = trainings[trainings["Training_ID"] == training_id].iloc[0]

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Progress", f"{record['Progress_%']}%")
    c2.metric("Slides", record["Slides_Opened"])
    c3.metric("Video", record["Video_Opened"])
    c4.metric("Test", record["Test_Status"])
    c5.metric("Certificate", record["Certificate_Status"])

    tab_material, tab_test, tab_record = st.tabs(["Training Material", "MCQ Test", "My Record"])

    with tab_material:
        m1, m2, m3 = st.columns(3)

        with m1:
            st.markdown("#### Slides")
            st.code(safe_text(training.get("Slides_Link", "")))
            if st.button("Mark Slides Complete"):
                trainee_activity(actor, training_id, "Slides_Opened", "Slides completed.")

        with m2:
            st.markdown("#### Video")
            st.code(safe_text(training.get("Video_Link", "")))
            if st.button("Mark Video Complete"):
                trainee_activity(actor, training_id, "Video_Opened", "Video completed.")

        with m3:
            st.markdown("#### Recording")
            st.code(safe_text(training.get("Recording_Link", "")))
            if st.button("Mark Recording Complete"):
                records = read_sheet("Training_Records")
                mask = (records["User_ID"] == user_id) & (records["Training_ID"] == training_id)

                records.loc[mask, "Recording_Opened"] = "Yes"
                records.loc[mask, "Video_Opened"] = "Yes"
                records.loc[mask, "Live_Attendance"] = "Recording Viewed"
                records.loc[mask, "Remarks"] = "Recording viewed."
                records.loc[mask, "Last_Updated"] = now_text()

                write_sheet("Training_Records", records)
                log_activity("Recording Viewed", user_id, actor_value(actor, "Name"), role, user_id, training_id)
                refresh_training_records()
                st.success("Recording completed.")
                st.rerun()

    with tab_test:
        qbank = read_sheet("Question_Bank")
        questions = qbank[qbank["Training_ID"] == training_id] if not qbank.empty else pd.DataFrame()

        if questions.empty:
            st.warning("MCQ test has not been generated by trainer yet.")
        else:
            st.caption("Select answers and submit. Certificate is issued automatically if score meets passing marks.")

            with st.form("mcq_test_form"):
                answers = {}
                for i, (_, q) in enumerate(questions.iterrows(), start=1):
                    st.markdown(f"**Q{i}. {q['Question']}**")
                    options = [q["Option_A"], q["Option_B"], q["Option_C"], q["Option_D"]]
                    answers[q["Question_ID"]] = st.radio(
                        "Select answer",
                        options,
                        key=f"q_{q['Question_ID']}",
                        label_visibility="collapsed",
                    )

                submitted = st.form_submit_button("Submit MCQ Test")

            if submitted:
                total = len(questions)
                correct = 0
                for _, q in questions.iterrows():
                    if answers.get(q["Question_ID"]) == q["Correct_Answer"]:
                        correct += 1

                score = round((correct / total) * 100, 2) if total else 0
                passing_marks = int(record["Passing_Marks"]) if str(record["Passing_Marks"]).strip() else 75
                result = "Passed" if score >= passing_marks else "Failed"

                records = read_sheet("Training_Records")
                mask = (records["User_ID"] == user_id) & (records["Training_ID"] == training_id)

                records.loc[mask, "Score"] = score
                records.loc[mask, "Test_Status"] = result
                records.loc[mask, "Remarks"] = f"MCQ test submitted. Correct: {correct}/{total}"
                records.loc[mask, "Last_Updated"] = now_text()

                write_sheet("Training_Records", records)
                log_activity("MCQ Test Submitted", user_id, actor_value(actor, "Name"), role, user_id, training_id, remarks=f"{score}% - {result}")

                if result == "Passed":
                    issue_certificate(actor, training_id, int(score))

                refresh_training_records()
                st.success(f"Test submitted. Score: {score}%. Result: {result}")
                st.rerun()

    with tab_record:
        records = read_sheet("Training_Records")
        show_table(records[records["User_ID"] == user_id])


# ============================================================
# MODULE 13 — MANAGEMENT PANEL
# ============================================================

def management_panel(actor: dict) -> None:
    st.header("Management Dashboard")
    st.info("Management can view progress and records, but Excel download and full database control are restricted to Admin.")

    for title, sheet in [
        ("Dashboard", "Dashboard"),
        ("Trainings", "Trainings"),
        ("Training Records", "Training_Records"),
        ("Certificates", "Certificates"),
    ]:
        st.subheader(title)
        show_table(read_sheet(sheet))


# ============================================================
# MODULE 14 — MAIN APP ROUTER
# ============================================================

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")

    if not database_is_valid():
        create_database(reset=True)

    apply_style()
    app_header()

    actor = require_login()
    sidebar_common(actor)

    name = actor_value(actor, "Name", "Unknown User")
    role = actor_value(actor, "Role", "Unknown Role")
    dept = actor_value(actor, "Department", "")
    email = actor_value(actor, "Email", "")

    st.markdown(
        f"""
        <div class="info-card">
            <b>{name}</b> | {role} | {dept} | {email}
        </div>
        """,
        unsafe_allow_html=True,
    )

    if role == "Admin":
        page = st.sidebar.radio("Menu", ["Dashboard", "Admin Panel"])
        if page == "Dashboard":
            dashboard_page(actor)
        else:
            admin_panel(actor)

    elif role == "Trainer":
        page = st.sidebar.radio("Menu", ["Dashboard", "Trainer Panel"])
        if page == "Dashboard":
            dashboard_page(actor)
        else:
            trainer_panel(actor)

    elif role in ["Surveyor", "Plan Appraiser"]:
        page = st.sidebar.radio("Menu", ["Dashboard", "Training Portal"])
        if page == "Dashboard":
            dashboard_page(actor)
        else:
            trainee_panel(actor)

    elif role == "Management":
        page = st.sidebar.radio("Menu", ["Dashboard", "Management View"])
        if page == "Dashboard":
            dashboard_page(actor)
        else:
            management_panel(actor)

    else:
        st.error("Unknown role. Please contact Admin.")

    if role == "Admin":
        st.divider()
        with st.expander("Admin Excel Dashboard Sheet"):
            show_table(read_sheet("Dashboard"))

        with st.expander("Admin Activity Log"):
            show_table(read_sheet("Activity_Log"))


if __name__ == "__main__":
    main()