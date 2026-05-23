from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import hashlib, importlib, io, os, random, re, secrets, string, uuid, smtplib, subprocess, sys, urllib.parse
from email.message import EmailMessage

import pandas as pd
import streamlit as st

APP_TITLE = "Human Resource Development and Management"
APP_SUBTITLE = "Classification Society Training, Competency and Certification Platform"
DB_FILE = "hrdm_training_database.xlsx"
PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "https://pakistan-shipping-bureau.psbureau.org")
DEFAULT_ADMIN_PASSWORD = os.getenv("DEFAULT_ADMIN_PASSWORD", "Admin@1234")
DEFAULT_MGMT_PASSWORD = os.getenv("DEFAULT_MGMT_PASSWORD", "Mgmt@1234")
DEFAULT_TRAINER_PASSWORD = os.getenv("DEFAULT_TRAINER_PASSWORD", "Trainer@1234")
DEFAULT_SURVEYOR_PASSWORD = os.getenv("DEFAULT_SURVEYOR_PASSWORD", "Surveyor@1234")
DEFAULT_APPRAISER_PASSWORD = os.getenv("DEFAULT_APPRAISER_PASSWORD", "Appraiser@1234")
DEFAULT_QMR_PASSWORD = os.getenv("DEFAULT_QMR_PASSWORD", "QMR@1234")
DEFAULT_TRAINEE_PASSWORD = os.getenv("DEFAULT_TRAINEE_PASSWORD", "Trainee@1234")
FAILED_LOGIN_ALERT_THRESHOLD = int(os.getenv("FAILED_LOGIN_ALERT_THRESHOLD", "5"))

AVAILABLE_ROLES = ["Admin", "Management", "Trainer", "Surveyor", "Plan Appraiser", "Quality Management Representative", "Rule Development Rep", "Trainee", "On Probation"]
AVAILABLE_DEPARTMENTS = ["Plan Appraisal", "Survey", "Rule Development", "Quality Management System", "Industry", "Support/Admin"]

SCHEMA = {
    "Users": ["User_ID","Name","Role","Department","Assigned_Duty","Email","Login_ID","Password_Hash","Password","Status","Created_By","Created_On","Last_Login","Failed_Attempts","Reset_Token","Reset_Expires"],
    "Trainings": ["Training_ID","Training_Title","Category","Target_Role","Trainer_ID","Trainer_Name","Slides_Link","Video_Link","Reference_Link","Schedule_Date","Schedule_Time","Meeting_Link","Recording_Link","Passing_Marks","Status","Created_By","Created_On","Last_Updated"],
    "Training_Content": ["Content_ID","Training_ID","File_Name","File_Type","Content_Text","Uploaded_By","Uploaded_On"],
    "Question_Bank": ["Question_ID","Training_ID","Question","Option_A","Option_B","Option_C","Option_D","Correct_Answer","Marks","Generated_On"],
    "Training_Records": ["Record_ID","User_ID","Name","Role","Training_ID","Training_Title","Status","Slides_Opened","Video_Opened","Live_Attendance","Recording_Opened","Test_Status","Score","Passing_Marks","Certificate_Status","Certificate_Link","Due_Date","Completed_On","Progress_%","Remarks","Last_Updated"],
    "Notifications": ["Notification_ID","Training_ID","User_ID","Name","Email","Subject","Message","Status","Generated_On","Sent_On","Generated_By"],
    "Certificates": ["Certificate_ID","User_ID","Name","Role","Training_ID","Training_Title","Score","Issued_On","Certificate_Link","Status","Issued_By"],
    "Activity_Log": ["Log_ID","Date_Time","Activity","Actor_ID","Actor_Name","Actor_Role","User_ID","Training_ID","Status","Remarks"],
    "Dashboard": ["Metric","Value"],
    "System": ["Key","Value"],
}


def now(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def today(): return date.today().strftime("%Y-%m-%d")
def uid(prefix): return f"{prefix}-{uuid.uuid4().hex[:8].upper()}"
def phash(password: str) -> str: return hashlib.sha256(password.encode("utf-8")).hexdigest()
def temp_password(n=10): return "".join(secrets.choice(string.ascii_letters + string.digits + "@#$") for _ in range(n))

def clean(v):
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except Exception: pass
    return str(v)

def url_ok(v):
    v = clean(v).strip()
    return v == "" or v.startswith(("http://", "https://"))

# Safe session-state helpers to avoid StreamlitAPIException when session state
# cannot be modified (e.g., during certain lifecycle stages). These swallow
# errors and ensure the app continues to run.
def safe_session_set(key, value):
    try:
        if key not in st.session_state:
            st.session_state[key] = value
    except Exception:
        try:
            st.session_state.update({key: value})
        except Exception:
            pass

def safe_session_update(d: dict):
    try:
        st.session_state.update(d)
    except Exception:
        for k, v in d.items():
            try:
                st.session_state[k] = v
            except Exception:
                pass

# Support multi-role Target_Role values stored as comma/semicolon/pipe-separated strings.
def parse_target_roles(value):
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    text = clean(value)
    if not text:
        return []
    return [part.strip() for part in re.split(r"[,;|]+", text) if part.strip()]

def actor_get(actor, key, default=""):
    aliases = {"User_ID":["User_ID","Staff_ID"],"Name":["Name","Staff_Name"],"Role":["Role"],"Department":["Department"],"Email":["Email"],"Login_ID":["Login_ID","Username"]}
    for k in aliases.get(key,[key]):
        if isinstance(actor, dict) and k in actor and clean(actor[k]).strip(): return clean(actor[k])
    return default

def login_id_from_name(name, users):
    base = re.sub(r"[^a-z0-9]", "", name.lower().replace(" ", ".")) or "user"
    existing = set(users["Login_ID"].astype(str).str.lower()) if not users.empty and "Login_ID" in users.columns else set()
    login = base; i = 1
    while login.lower() in existing:
        i += 1; login = f"{base}{i}"
    return login

# ---------------- Excel DB ----------------
def read_sheet(sheet):
    cols = SCHEMA.get(sheet, [])
    if not Path(DB_FILE).exists(): return pd.DataFrame(columns=cols)
    try:
        df = pd.read_excel(DB_FILE, sheet_name=sheet, engine="openpyxl")
    except Exception:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns: df[c] = ""
    df = df[cols] if cols else df
    if not df.empty:
        df = df.astype(object)
    return df

def write_sheet(sheet, df):
    cols = SCHEMA.get(sheet, [])
    for c in cols:
        if c not in df.columns: df[c] = ""
    if cols: df = df[cols]
    mode = "a" if Path(DB_FILE).exists() else "w"
    kwargs = {"engine":"openpyxl", "mode":mode}
    if mode == "a": kwargs["if_sheet_exists"] = "replace"
    with pd.ExcelWriter(DB_FILE, **kwargs) as w:
        df.to_excel(w, sheet_name=sheet, index=False)

def append_row(sheet, row):
    df = read_sheet(sheet)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    write_sheet(sheet, df)

def db_valid():
    users = read_sheet("Users")
    return Path(DB_FILE).exists() and not users.empty and set(SCHEMA["Users"]).issubset(users.columns)

def create_db(reset=False):
    if Path(DB_FILE).exists() and not reset and db_valid(): return
    users = pd.DataFrame([
        ["USR-ADMIN","Admin User","Admin","Support/Admin","System Admin","usama.saleem@psbureau.org","admin",phash(DEFAULT_ADMIN_PASSWORD),DEFAULT_ADMIN_PASSWORD,"Active","System",today(),"",0,"",""],
        ["USR-MGMT","Management User","Management","Support/Admin","Management Oversight","management@psbureau.org","management",phash(DEFAULT_MGMT_PASSWORD),DEFAULT_MGMT_PASSWORD,"Active","System",today(),"",0,"",""],
        ["USR-TRAINER","Training Officer","Trainer","Support/Admin","Training Delivery","trainer@psbureau.org","trainer",phash(DEFAULT_TRAINER_PASSWORD),DEFAULT_TRAINER_PASSWORD,"Active","System",today(),"",0,"",""],
        ["USR-SURVEYOR","Sample Surveyor","Surveyor","Survey","Survey Field Work","surveyor@psbureau.org","surveyor",phash(DEFAULT_SURVEYOR_PASSWORD),DEFAULT_SURVEYOR_PASSWORD,"Active","System",today(),"",0,"",""],
        ["USR-APPRAISER","Sample Plan Appraiser","Plan Appraiser","Plan Appraisal","Plan Appraisal Review","appraiser@psbureau.org","appraiser",phash(DEFAULT_APPRAISER_PASSWORD),DEFAULT_APPRAISER_PASSWORD,"Active","System",today(),"",0,"",""],
        ["USR-QMR","Quality Management Rep","Quality Management Representative","Quality Management System","QMS Oversight","qmr@psbureau.org","qmr",phash(DEFAULT_QMR_PASSWORD),DEFAULT_QMR_PASSWORD,"Active","System",today(),"",0,"",""],
        ["USR-TRAINEE","Sample Trainee","Trainee","Industry","Training Completion","trainee@psbureau.org","trainee",phash(DEFAULT_TRAINEE_PASSWORD),DEFAULT_TRAINEE_PASSWORD,"Active","System",today(),"",0,"",""],
    ], columns=SCHEMA["Users"])
    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as w:
        users.to_excel(w, sheet_name="Users", index=False)
        pd.DataFrame([["APP_TITLE",APP_TITLE],["CREATED_ON",now()]], columns=SCHEMA["System"]).to_excel(w, sheet_name="System", index=False)
        for s,c in SCHEMA.items():
            if s not in ["Users","System"]: pd.DataFrame(columns=c).to_excel(w, sheet_name=s, index=False)
    log("Database Created", actor_name="System", actor_role="System", update=False)
    update_dashboard()

# ---------------- Logs / Status ----------------
def log(activity, actor_id="", actor_name="", actor_role="", user_id="", training_id="", status="Success", remarks="", update=True):
    append_row("Activity_Log", {"Log_ID":uid("LOG"),"Date_Time":now(),"Activity":activity,"Actor_ID":actor_id,"Actor_Name":actor_name,"Actor_Role":actor_role,"User_ID":user_id,"Training_ID":training_id,"Status":status,"Remarks":remarks})
    if update: update_dashboard()

def progress(row):
    steps = [row.get("Slides_Opened")=="Yes", row.get("Video_Opened")=="Yes" or row.get("Recording_Opened")=="Yes", row.get("Live_Attendance") in ["Present","Recording Viewed"], row.get("Test_Status")=="Passed", row.get("Certificate_Status")=="Issued"]
    return int(sum(steps)/len(steps)*100)

def refresh_records():
    rec = read_sheet("Training_Records")
    if rec.empty: update_dashboard(); return
    for i,row in rec.iterrows():
        p = progress(row); rec.at[i,"Progress_%"] = p; rec.at[i,"Status"] = "Completed" if p == 100 else "Pending"; rec.at[i,"Last_Updated"] = now()
        if p == 100 and clean(row.get("Completed_On")) == "": rec.at[i,"Completed_On"] = today()
    write_sheet("Training_Records", rec); update_dashboard()

def update_dashboard():
    users, trainings, rec, certs, qs, logs = [read_sheet(s) for s in ["Users","Trainings","Training_Records","Certificates","Question_Bank","Activity_Log"]]
    completed = len(rec[rec["Status"]=="Completed"]) if not rec.empty else 0
    pending = len(rec[rec["Status"]!="Completed"]) if not rec.empty else 0
    avg = round(rec["Progress_%"].fillna(0).mean(),2) if not rec.empty else 0
    dash = pd.DataFrame([
        ["Total Users",len(users)], ["Active Users",len(users[users["Status"]=="Active"]) if not users.empty else 0], ["Surveyors",len(users[users["Role"]=="Surveyor"]) if not users.empty else 0],
        ["Plan Appraisers",len(users[users["Role"]=="Plan Appraiser"]) if not users.empty else 0], ["Trainers",len(users[users["Role"]=="Trainer"]) if not users.empty else 0], ["Total Trainings",len(trainings)],
        ["Generated MCQs",len(qs)], ["Completed Records",completed], ["Pending Records",pending], ["Average Progress %",avg], ["Certificates Issued",len(certs)], ["Activity Logs",len(logs)], ["Last Updated",now()]
    ], columns=SCHEMA["Dashboard"])
    write_sheet("Dashboard", dash)

# ---------------- Security ----------------
def generate_captcha():
    a,b = random.randint(2,15), random.randint(2,15)
    safe_session_update({"captcha_question": f"{a} + {b}", "captcha_answer": str(a+b)})

def reset_session():
    safe_session_update({"logged_in": False, "user": {}})
    generate_captcha()

def authenticate(login_id, password):
    users = read_sheet("Users")
    login_id = login_id.strip()
    password_hash = phash(password.strip())
    match = (users["Login_ID"].astype(str)==login_id) | (users["Email"].astype(str).str.lower()==login_id.lower())
    m = users[match & (users["Password_Hash"].astype(str)==password_hash) & (users["Status"].astype(str)=="Active")]
    if m.empty: return None
    idx=m.index[0]; users.at[idx,"Last_Login"]=now(); users.at[idx,"Failed_Attempts"]=0; write_sheet("Users",users); return users.loc[idx].to_dict()

def failed_login(login_id):
    login_id = login_id.strip()
    users = read_sheet("Users")
    m = users[(users["Login_ID"].astype(str) == login_id) | (users["Email"].astype(str).str.lower() == login_id.lower())]
    if m.empty:
        return
    idx = m.index[0]
    fail = int(users.at[idx, "Failed_Attempts"] or 0) + 1
    users.at[idx, "Failed_Attempts"] = fail
    write_sheet("Users", users)
    # If attempts reach threshold, alert admins but do NOT deactivate the user automatically.
    if fail >= FAILED_LOGIN_ALERT_THRESHOLD:
        try:
            user_id = users.at[idx, "User_ID"] if "User_ID" in users.columns else ""
            user_email = users.at[idx, "Email"] if "Email" in users.columns else ""
        except Exception:
            user_id = ""
            user_email = ""
        alert_admin_failed_login(login_id, user_id, user_email, fail)

def send_email(to_email, subject, body):
    host = os.getenv("SMTP_HOST", "")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER", "")
    password = os.getenv("SMTP_PASSWORD", "")
    from_addr = os.getenv("SMTP_FROM", user or "noreply@example.com")
    if not host or not user or not password:
        return False
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = to_email
    msg.set_content(body)
    try:
        with smtplib.SMTP(host, port) as smtp:
            smtp.starttls()
            smtp.login(user, password)
            smtp.send_message(msg)
        return True
    except Exception:
        return False

def alert_admin_failed_login(login_id, user_id, user_email, attempts):
    """Notify Admin(s) when a user has multiple failed login attempts."""
    users = read_sheet("Users")
    admins = []
    try:
        if not users.empty and "Role" in users.columns and "Email" in users.columns:
            admins = users[(users["Role"] == "Admin") & (users["Status"] == "Active")]["Email"].astype(str).dropna().unique().tolist()
    except Exception:
        admins = []
    subject = f"Security Alert: Multiple failed login attempts for {login_id}"
    message = f"User Login: {login_id}\nUser ID: {user_id}\nUser Email: {user_email}\nFailed Attempts: {attempts}\n\nPlease review activity and consider contacting the user."
    # Create a notification record
    try:
        append_row("Notifications", {"Notification_ID": uid("NOT"), "Training_ID": "", "User_ID": user_id, "Name": "", "Email": ",".join(admins) if admins else "", "Subject": subject, "Message": message, "Status": "Pending", "Generated_On": now(), "Sent_On": "", "Generated_By": "System"})
    except Exception:
        pass
    # Log the alert
    log("Failed Login Alert", actor_name="System", actor_role="System", remarks=f"{login_id} — {attempts} attempts")
    # Attempt to email admins if SMTP is configured
    for a in admins:
        try:
            send_email(a, subject, message)
        except Exception:
            pass

def generate_reset_token():
    return secrets.token_urlsafe(16)

def request_password_reset(login_id):
    users = read_sheet("Users")
    login_id = login_id.strip()
    m = users[(users["Login_ID"].astype(str)==login_id) | (users["Email"].astype(str).str.lower()==login_id.lower())]
    if m.empty: return False, "No matching user found."
    idx = m.index[0]
    if not clean(users.at[idx,"Email"]):
        return False, "The user does not have an email address on file."
    token = generate_reset_token()
    expires = (datetime.now() + timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
    users.at[idx,"Reset_Token"] = token
    users.at[idx,"Reset_Expires"] = expires
    write_sheet("Users",users)
    body = f"""Dear {users.at[idx,'Name']},

A password reset request has been received for your account in the Human Resource Development and Management Platform.

RESET TOKEN (Valid for 1 hour):
{token}

INSTRUCTIONS:
1. Visit the login page of the platform
2. Click "Forgot Password? Secure Reset"
3. Enter your Login ID or Email: {users.at[idx,'Email']}
4. Click "Request Reset Email" to get this token
5. Paste the reset token above in the "Reset Token" field
6. Enter your new password in the "New Password" field
7. Click "Reset Password with Token"

SECURITY NOTES:
- This token expires at: {expires}
- If you did not request this reset, please ignore this email or contact support immediately
- Do not share this token with anyone
- Support: support@psbureau.org

Regards,
Human Resource Development and Management System
Pakistan Shipping Bureau"""
    sent = send_email(users.at[idx,"Email"], "Password Reset Request — Action Required", body)
    if sent:
        return True, "A password reset email has been sent."
    return False, "Unable to send reset email. Please contact Admin."

def reset_password_with_token(login_id, token, new_password):
    users = read_sheet("Users")
    login_id = login_id.strip()
    m = users[(users["Login_ID"].astype(str)==login_id) | (users["Email"].astype(str).str.lower()==login_id.lower())]
    if m.empty: return False, "No matching user found."
    idx = m.index[0]
    if clean(users.at[idx,"Reset_Token"]) != token:
        return False, "Invalid reset token."
    expires = users.at[idx,"Reset_Expires"]
    if expires and datetime.now() > datetime.strptime(expires, "%Y-%m-%d %H:%M:%S"):
        return False, "Reset token has expired."
    users.at[idx,"Password_Hash"] = phash(new_password)
    users.at[idx,"Password"] = new_password
    users.at[idx,"Reset_Token"] = ""
    users.at[idx,"Reset_Expires"] = ""
    write_sheet("Users",users)
    return True, "Password has been reset successfully."

def login_page():
    st.subheader("Secure Login")
    st.caption("Only Admin-created users can enter the system.")
    if "captcha_question" not in st.session_state: generate_captcha()
    with st.form("login_form"):
        login_id=st.text_input("Login ID or Email"); pw=st.text_input("Password", type="password"); cap=st.text_input(f"Security Verification: {st.session_state['captcha_question']} = ?")
        sub=st.form_submit_button("Login")
    if sub:
        if cap.strip()!=st.session_state.get("captcha_answer",""):
            st.error("Security verification failed."); generate_captcha()
        user=authenticate(login_id,pw)
        if not user:
            failed_login(login_id); st.error("Invalid credentials or inactive user."); generate_captcha()
        else:
                safe_session_update({"logged_in": True, "user": user})
                log("User Login",user["User_ID"],user["Name"],user["Role"],user_id=user["User_ID"],remarks="Successful login")
    with st.expander("Forgot Password? Secure Reset"):
        reset_id = st.text_input("Enter Login ID or Email to request a reset", key="reset_login")
        if st.button("Request Reset Email"):
            if not reset_id.strip():
                st.error("Enter your Login ID or Email.")
            else:
                ok,msg=request_password_reset(reset_id)
                if ok: st.success(msg)
                else: st.error(msg)
        st.markdown("---")
        token = st.text_input("Reset Token", key="reset_token")
        new_pw = st.text_input("New Password", type="password", key="reset_new_pw")
        if st.button("Reset Password with Token"):
            if not reset_id.strip() or not token.strip() or not new_pw:
                st.error("Provide Login ID/Email, reset token, and a new password.")
            else:
                ok,msg=reset_password_with_token(reset_id, token.strip(), new_pw)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

def require_login():
    safe_session_set("logged_in", False)
    safe_session_set("user", {})
    u = st.session_state.get("user", {})
    if st.session_state.get("logged_in") and not all(k in u for k in ["User_ID", "Name", "Role", "Email"]):
        reset_session()
    if not st.session_state.get("logged_in"):
        login_page()
        if not st.session_state.get("logged_in"):
            st.stop()
    return st.session_state.get("user", {})

# ---------------- Content and MCQs ----------------
EXTRACTOR_PACKAGES = {
    "docx": "python-docx",
    "pptx": "python-pptx",
    "PyPDF2": "PyPDF2",
}

def ensure_module(import_name, package_name):
    """Import module; if missing, install it automatically in the current Python environment."""
    try:
        return importlib.import_module(import_name)
    except ModuleNotFoundError:
        try:
            with st.spinner(f"Installing missing package: {package_name} ..."):
                subprocess.check_call([sys.executable, "-m", "pip", "install", package_name, "--quiet"])
            importlib.invalidate_caches()
            module = importlib.import_module(import_name)
            st.success(f"{package_name} installed successfully.")
            return module
        except Exception as e:
            st.error(f"Unable to install {package_name}. In terminal run: python -m pip install {package_name}. Details: {e}")
            return None

def check_extractors_panel():
    st.caption("TXT works directly. DOCX/PPTX/PDF support is checked and installed automatically if missing.")
    if st.button("Check / Install File Extraction Packages"):
        ok = True
        for mod, pkg in EXTRACTOR_PACKAGES.items():
            ok = ensure_module(mod, pkg) is not None and ok
        if ok:
            st.success("DOCX, PPTX and PDF extraction support is ready.")

def read_upload(uploaded):
    if uploaded is None:
        return ""
    name = uploaded.name.lower()
    if name.endswith(".txt"):
        try:
            return uploaded.getvalue().decode("utf-8", errors="ignore")
        except Exception as e:
            st.error(f"Could not read TXT {uploaded.name}: {e}")
            return ""
    def decode_bytes():
        try:
            return uploaded.getvalue().decode("utf-8", errors="ignore")
        except Exception:
            try:
                return uploaded.getvalue().decode("latin-1", errors="ignore")
            except Exception:
                return ""

    if name.endswith(".docx"):
        docx = ensure_module("docx", "python-docx")
        if docx is not None:
            try:
                d = docx.Document(io.BytesIO(uploaded.getvalue()))
                return "\n".join(p.text.strip() for p in d.paragraphs if p.text.strip())
            except Exception:
                pass
        return decode_bytes()
    if name.endswith(".doc"):
        return decode_bytes()
    if name.endswith(".pdf"):
        pypdf = ensure_module("PyPDF2", "PyPDF2")
        if pypdf is not None:
            try:
                reader = pypdf.PdfReader(io.BytesIO(uploaded.getvalue()))
                text = "\n".join((page.extract_text() or "") for page in reader.pages).strip()
                if text:
                    return text
            except Exception:
                pass
        return decode_bytes()
    if name.endswith(".pptx"):
        pptx = ensure_module("pptx", "python-pptx")
        if pptx is not None:
            try:
                prs = pptx.Presentation(io.BytesIO(uploaded.getvalue()))
                text = []
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            content = shape.text.strip()
                            if content:
                                text.append(content)
                result = "\n".join(text).strip()
                if result:
                    return result
            except Exception:
                pass
        return decode_bytes()
    if name.endswith(".ppt"):
        return decode_bytes()
    if name.endswith(".txt"):
        try:
            return uploaded.getvalue().decode("utf-8", errors="ignore")
        except Exception:
            return decode_bytes()
    st.error(f"{uploaded.name}: unsupported file type. Use TXT, DOC, DOCX, PDF, PPT or PPTX.")
    return ""

def read_uploads(uploaded_files):
    if not uploaded_files:
        return "", pd.DataFrame(columns=["File_Name", "File_Type", "Readable", "Words"])
    all_text = []
    report = []
    for uploaded in uploaded_files:
        text = read_upload(uploaded)
        words = len(text.split()) if text else 0
        report.append({
            "File_Name": uploaded.name,
            "File_Type": uploaded.name.split(".")[-1].lower(),
            "Readable": "Yes" if words > 0 else "No",
            "Words": words,
        })
        if text:
            all_text.append(f"\n\n--- SOURCE FILE: {uploaded.name} ---\n{text}")
    return "\n".join(all_text).strip(), pd.DataFrame(report)

def keywords(text):
    sw={"training","system","should","shall","which","there","their","about","through","during","after","before","within","using","based","these","those","where","under","requirements","procedure","document","classification","society","survey","surveyor","appraisal","management","development"}
    out=[]
    for w in re.findall(r"\b[A-Za-z][A-Za-z\-]{4,}\b", text):
        x=w.lower().strip(".,:;()[]{}")
        if x not in sw and len(x)>=5 and x.title() not in out: out.append(x.title())
    return out[:80]

def generate_mcqs(training_id, text, max_q=10):
    sents=[s.strip() for s in re.split(r"(?<=[.!?])\s+", text.replace("\n"," ")) if 45<=len(s.strip())<=240]
    keys=keywords(text)
    rows=[]; random.shuffle(sents)
    if len(keys)<4: return pd.DataFrame(columns=SCHEMA["Question_Bank"])
    for s in sents:
        if len(rows)>=max_q: break
        ans=next((k for k in keys if re.search(rf"\b{re.escape(k)}\b", s, re.I)), None)
        if not ans: continue
        d=[k for k in keys if k.lower()!=ans.lower()]
        if len(d)<3: continue
        opts=random.sample(d,3)+[ans]; random.shuffle(opts)
        q=re.sub(rf"\b{re.escape(ans)}\b","__________",s,flags=re.I,count=1)
        rows.append({"Question_ID":uid("Q"),"Training_ID":training_id,"Question":q,"Option_A":opts[0],"Option_B":opts[1],"Option_C":opts[2],"Option_D":opts[3],"Correct_Answer":ans,"Marks":1,"Generated_On":now()})
    return pd.DataFrame(rows, columns=SCHEMA["Question_Bank"])

# ---------------- Business ----------------
def meeting_link(tid, title="Training Session", schedule_date="", schedule_time="10:00 AM"):
    """Create a Microsoft Teams meeting creation link.

    This opens Microsoft Teams with meeting details pre-filled. After Teams creates
    the real meeting, paste the final Teams join link back into the Meeting Link
    field before clicking Schedule and Notify.
    """
    subject = f"{title} - {tid}".strip()
    body = (
        f"Training ID: {tid}\n"
        f"Training Title: {title}\n"
        f"Schedule: {schedule_date} {schedule_time}\n\n"
        "Please join this HRDM training session through Microsoft Teams."
    )
    return (
        "https://teams.microsoft.com/l/meeting/new?"
        f"subject={urllib.parse.quote_plus(subject)}"
        f"&content={urllib.parse.quote_plus(body)}"
    )
def cert_link(user_id, tid): return f"https://certificate.psbureau.org/{user_id}/{tid}"

def notify_training(tid, actor):
    trainings, rec, users = read_sheet("Trainings"), read_sheet("Training_Records"), read_sheet("Users")
    tr = trainings[trainings["Training_ID"]==tid]
    if tr.empty: return
    t=tr.iloc[0]
    for _,r in rec[rec["Training_ID"]==tid].iterrows():
        u=users[users["User_ID"]==r["User_ID"]]
        if u.empty: continue
        p=u.iloc[0]
        msg=f"Dear {p['Name']},\n\nTraining: {t['Training_Title']}\nSchedule: {t['Schedule_Date']} {t['Schedule_Time']}\nMeeting: {t['Meeting_Link']}\nSlides: {t['Slides_Link']}\nVideo: {t['Video_Link']}\n\nRegards,\nHuman Resource Development and Management"
        append_row("Notifications", {"Notification_ID":uid("NOTIF"),"Training_ID":tid,"User_ID":p["User_ID"],"Name":p["Name"],"Email":p["Email"],"Subject":f"Training Assigned: {t['Training_Title']}","Message":msg,"Status":"Generated","Generated_On":now(),"Sent_On":"","Generated_By":actor_get(actor,"Name")})
    log("Notifications Generated",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),training_id=tid)

def issue_certificate(actor, tid, score):
    user_id=actor_get(actor,"User_ID"); rec=read_sheet("Training_Records"); cert=read_sheet("Certificates")
    mask=(rec["User_ID"]==user_id)&(rec["Training_ID"]==tid)
    if rec[mask].empty: return
    row=rec[mask].iloc[0]; link=cert_link(user_id,tid)
    rec.loc[mask,"Certificate_Status"]="Issued"; rec.loc[mask,"Certificate_Link"]=link; rec.loc[mask,"Remarks"]="Certificate issued after passing test."; rec.loc[mask,"Last_Updated"]=now(); write_sheet("Training_Records",rec)
    exists=cert[(cert["User_ID"]==user_id)&(cert["Training_ID"]==tid)]
    if exists.empty:
        cert=pd.concat([cert,pd.DataFrame([{"Certificate_ID":uid("CERT"),"User_ID":user_id,"Name":actor_get(actor,"Name"),"Role":actor_get(actor,"Role"),"Training_ID":tid,"Training_Title":row["Training_Title"],"Score":score,"Issued_On":today(),"Certificate_Link":link,"Status":"Issued","Issued_By":"System"}])],ignore_index=True)
        write_sheet("Certificates",cert)
    log("Certificate Issued",user_id,actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=user_id,training_id=tid); refresh_records()

# ---------------- UI ----------------
def table(df): st.dataframe(df.fillna(""), width="stretch", hide_index=True)

def style():
    st.markdown("""
    <style>
    :root { color-scheme: dark light; }
    .block-container { padding-top:1.2rem; padding-bottom:2rem; font-family: Inter, system-ui, sans-serif; }
    .hdr { padding:1.4rem 1.6rem; border-radius:24px; background: linear-gradient(135deg,#0b1120,#14203a); color:#f8fafc; margin-bottom:1.3rem; box-shadow:0 20px 50px rgba(15,23,42,.17); }
    .hdr h1 { margin:0; font-size:2.4rem; letter-spacing:.02em; line-height:1.05; }
    .hdr p { margin:.55rem 0 0; color:#cbd5e1; font-size:1rem; max-width:780px; }
    .info-box { padding:1rem 1.2rem; border-radius:18px; background:#f8fafc; color:#0f172a; border:1px solid #e2e8f0; box-shadow:0 10px 30px rgba(15,23,42,.06); margin-bottom:1rem; }
    .info-box a { color:#0f172a; text-decoration:none; font-weight:600; }
    .card { padding:1.1rem 1.2rem; border-radius:18px; background:#ffffff; border:1px solid #e2e8f0; box-shadow:0 10px 25px rgba(15,23,42,.05); margin-bottom:1rem; }
    .card h3 { margin-top:0; font-size:1.05rem; color:#0f172a; }
    .card p { margin:.65rem 0 0; color:#334155; line-height:1.6; }
    .step { padding:.9rem 1rem; border-radius:14px; background:#f8fafc; border-left:4px solid #0f172a; margin-bottom:.55rem; color:#1e293b; }
    .footer { padding:1rem 0; margin-top:2rem; border-top:1px solid #e2e8f0; color:#64748b; font-size:.95rem; text-align:center; }
    .footer a { color:#0f172a; text-decoration:none; font-weight:600; }
    </style>""", unsafe_allow_html=True)

def header():
    st.markdown(f"<div class='hdr'><h1>{APP_TITLE}</h1><p>{APP_SUBTITLE}</p></div>", unsafe_allow_html=True)
    if PUBLIC_URL:
        st.markdown(f"<div class='info-box'>Professional access URL: <a href='{PUBLIC_URL}' target='_blank'>{PUBLIC_URL}</a></div>", unsafe_allow_html=True)

def card(title, body): st.markdown(f"<div class='card'><h3>{title}</h3><p>{body}</p></div>", unsafe_allow_html=True)
def steps(items):
    for i,x in enumerate(items,1): st.markdown(f"<div class='step'><b>{i}.</b> {x}</div>", unsafe_allow_html=True)

def footer():
    st.markdown("<div class='footer'>Powered by Human Resource Development and Management Platform. Secure training, assessment, and certification for modern workforce operations. <a href='https://psbureau.org' target='_blank'>Learn more</a>.</div>", unsafe_allow_html=True)

def sidebar(actor):
    role=actor_get(actor,"Role"); name=actor_get(actor,"Name")
    st.sidebar.success(f"{name} ({role})")
    st.sidebar.caption(actor_get(actor,"Email"))
    st.sidebar.markdown("**Platform**\nHuman Resource Development and Management")
    st.sidebar.markdown("**Support**\n[support@psbureau.org](mailto:support@psbureau.org)")
    st.sidebar.divider()
    if st.sidebar.button("Logout"):
        log("User Logout",actor_get(actor,"User_ID"),name,role,user_id=actor_get(actor,"User_ID")); reset_session()
    if role=="Admin":
        if Path(DB_FILE).exists():
            with open(DB_FILE,"rb") as f: st.sidebar.download_button("Download Excel Database",f,file_name=DB_FILE,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def dashboard(actor):
    role=actor_get(actor,"Role"); st.header(f"{role} Dashboard")
    flows={"Admin":["Create users with automatic Login ID and temporary password.","Create training and assign trainer.","Assign trainees and monitor Excel database."],"Trainer":["Add material links.","Upload content and generate MCQs.","Schedule training, notify trainees, mark attendance."],"Surveyor":["Complete assigned training.","Take MCQ test.","Receive certificate after passing."],"Plan Appraiser":["Complete assigned training.","Take MCQ test.","Receive certificate after passing."],"Quality Management Representative":["Monitor quality standards.","Review training compliance.","Generate quality reports."],"Rule Development Rep":["Complete assigned training.","Take MCQ test.","Participate in rule development updates."],"Trainee":["Complete assigned training.","Take MCQ test.","Receive certificate after passing."],"On Probation":["Complete assigned training.","Take MCQ test.","Receive certificate after passing."],"Management":["Monitor HR development progress.","Review training and certificates."]}
    card("Role Workflow","Only allowed role functions are shown."); steps(flows.get(role,[]))
    dash=read_sheet("Dashboard")
    if not dash.empty:
        st.subheader("System Summary"); cols=st.columns(4)
        for i,r in dash.iterrows(): cols[i%4].metric(str(r["Metric"]), str(r["Value"]))

def admin_page(actor):
    st.header("Admin Panel")
    card("Admin Control","Admin manages users, roles, trainings, trainee assignments, and full Excel database access.")
    t1,t2,t3,t4=st.tabs(["Users & Roles","Create Training","Assign Trainees","Admin Database"])
    with t1:
        users=read_sheet("Users"); st.subheader("Users"); table(users.drop(columns=["Password_Hash","Password"], errors="ignore"))
        with st.form("new_user"):
            name=st.text_input("Name"); role=st.selectbox("Role",AVAILABLE_ROLES); dept=st.selectbox("Department",AVAILABLE_DEPARTMENTS); duty=st.text_input("Assigned Duty"); email=st.text_input("Email"); password=st.text_input("Password (leave blank to auto-generate)", type="password"); status=st.selectbox("Status",["Active","Inactive"]); sub=st.form_submit_button("Create User and Generate Login")
        if sub:
            if not name or not email: st.error("Name and Email are required.")
            else:
                users=read_sheet("Users"); user_id=uid("USR"); login=login_id_from_name(name,users); pw=password.strip() or temp_password()
                row={"User_ID":user_id,"Name":name,"Role":role,"Department":dept,"Assigned_Duty":duty.strip(),"Email":email,"Login_ID":login,"Password_Hash":phash(pw),"Password":pw,"Status":status,"Created_By":actor_get(actor,"Name"),"Created_On":today(),"Last_Login":"","Failed_Attempts":0,"Reset_Token":"","Reset_Expires":""}
                users=pd.concat([users,pd.DataFrame([row])],ignore_index=True); write_sheet("Users",users); log("User Created",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=user_id); update_dashboard()
                st.success("User created. Share securely:"); st.code(f"User ID: {user_id}\nLogin ID: {login}\nEmail: {email}\nPassword: {pw}")
        st.subheader("Password Reset / Status")
        users=read_sheet("Users")
        if not users.empty:
            selected=st.selectbox("Select User", users["Name"].astype(str)+" — "+users["User_ID"].astype(str))
            sid=selected.split(" — ")[-1]
            c1,c2,c3=st.columns([2,2,1])
            with c1:
                if st.button("Reset Password"):
                    pw=temp_password(); users=read_sheet("Users"); idx=users[users["User_ID"]==sid].index[0]; users.at[idx,"Password_Hash"]=phash(pw); users.at[idx,"Password"]=pw; users.at[idx,"Failed_Attempts"]=0; users.at[idx,"Status"]="Active"; users.at[idx,"Reset_Token"]=""; users.at[idx,"Reset_Expires"]=""; write_sheet("Users",users); log("Password Reset",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=sid); st.success("New password:"); st.code(pw)
            with c2:
                if st.button("Send Reset Email"):
                    users=read_sheet("Users"); idx=users[users["User_ID"]==sid].index[0]; ok,msg=request_password_reset(users.at[idx,"Login_ID"] or users.at[idx,"Email"])
                    if ok: st.success(msg)
                    else: st.error(msg)
            with c3:
                status=st.selectbox("New Status",["Active","Inactive","Left"], key="status_select")
                if st.button("Save Status"):
                    users=read_sheet("Users"); idx=users[users["User_ID"]==sid].index[0]; users.at[idx,"Status"]=status; write_sheet("Users",users); log("User Status Updated",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=sid,remarks=status)
        st.subheader("Assign Duty / Remove User")
        if not users.empty:
            selected2=st.selectbox("Select User to Update", users["Name"].astype(str)+" — "+users["User_ID"].astype(str), key="manage_user")
            sid2=selected2.split(" — ")[-1]
            idx2=users[users["User_ID"]==sid2].index[0]
            duty=st.text_input("Assigned Duty", users.at[idx2,"Assigned_Duty"], key="assigned_duty")
            if st.button("Assign Duty"):
                users=read_sheet("Users"); idx=users[users["User_ID"]==sid2].index[0]; users.at[idx,"Assigned_Duty"]=duty.strip(); write_sheet("Users",users); log("Duty Assigned",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=sid2,remarks=f"Duty: {duty.strip()}"); st.success("Duty assigned.")
            if st.button("Mark User as Left"):
                users=read_sheet("Users"); idx=users[users["User_ID"]==sid2].index[0]; users.at[idx,"Status"]="Left"; write_sheet("Users",users); log("User Marked Left",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=sid2,remarks="Left company"); st.success("User marked as left. Records preserved.")
    with t2:
        users=read_sheet("Users"); trainers=users[(users["Role"]=="Trainer")&(users["Status"]=="Active")]
        if trainers.empty: st.warning("Create an active Trainer first.")
        else:
            safe_session_set("training_title", "")
            safe_session_set("training_category", "Basic Survey")
            safe_session_set("training_target", [])
            if "training_trainer" not in st.session_state:
                safe_session_set("training_trainer", (trainers["Name"].astype(str)+" — "+trainers["User_ID"].astype(str)).iloc[0])
            with st.form("training"):
                title=st.text_input("Training Title", key="training_title")
                cat=st.selectbox("Category",["Basic Survey","Electrical Survey","Hull Survey","Machinery Survey","Plan Appraisal","Statutory Survey","Report Writing","Safety","Quality Management","Rule Development"], key="training_category")
                target=st.multiselect("Target Role",["Surveyor","Plan Appraiser","Quality Management Representative","Rule Development Rep","Trainee","On Probation"], key="training_target")
                trainer=st.selectbox("Trainer",trainers["Name"].astype(str)+" — "+trainers["User_ID"].astype(str), key="training_trainer")
                passing=st.number_input("Passing Marks (%)",1,100,75, key="training_passing")
                sub=st.form_submit_button("Create Training")
            if sub:
                if not title:
                    st.error("Training title is required.")
                elif not target:
                    st.error("Select at least one target role.")
                else:
                    tid=uid("TRN")
                    tname,trainer_id=trainer.split(" — ")
                    target_val=", ".join(target)
                    tr=read_sheet("Trainings")
                    row={"Training_ID":tid,"Training_Title":title,"Category":cat,"Target_Role":target_val,"Trainer_ID":trainer_id,"Trainer_Name":tname,"Slides_Link":"","Video_Link":"","Reference_Link":"","Schedule_Date":"","Schedule_Time":"","Meeting_Link":"","Recording_Link":"","Passing_Marks":int(passing),"Status":"Draft","Created_By":actor_get(actor,"Name"),"Created_On":now(),"Last_Updated":now()}
                    tr=pd.concat([tr,pd.DataFrame([row])],ignore_index=True)
                    write_sheet("Trainings",tr)
                    log("Training Created",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),training_id=tid)
                    st.success(f"Training created: {tid}")
                    safe_session_update({
                        "training_title": "",
                        "training_category": "Basic Survey",
                        "training_target": [],
                        "training_trainer": (trainers["Name"].astype(str)+" — "+trainers["User_ID"].astype(str)).iloc[0],
                        "training_passing": 75,
                    })
    with t3:
        users=read_sheet("Users"); trainings=read_sheet("Trainings")
        if trainings.empty: st.warning("No training created.")
        else:
            display=st.selectbox("Training",trainings["Training_Title"].astype(str)+" — "+trainings["Training_ID"].astype(str), key="assign_training_display")
            tid=display.split(" — ")[-1]; tr=trainings[trainings["Training_ID"]==tid].iloc[0]
            target_roles = parse_target_roles(tr["Target_Role"])
            eligible=users[(users["Role"].isin(target_roles))&(users["Status"]=="Active")]
            selected=st.multiselect("Select Trainees",eligible["Name"].astype(str)+" — "+eligible["User_ID"].astype(str), key="assign_selected")
            due=st.date_input("Due Date", key="assign_due_date")
            if st.button("Assign Trainees"):
                if not selected: st.warning("Select at least one trainee.")
                else:
                    rec=read_sheet("Training_Records"); added=0
                    for item in selected:
                        name,user_id=item.split(" — "); u=users[users["User_ID"]==user_id].iloc[0]
                        if not rec[(rec["User_ID"]==user_id)&(rec["Training_ID"]==tid)].empty: continue
                        row={"Record_ID":uid("REC"),"User_ID":user_id,"Name":name,"Role":u["Role"],"Training_ID":tid,"Training_Title":tr["Training_Title"],"Status":"Pending","Slides_Opened":"No","Video_Opened":"No","Live_Attendance":"Not Marked","Recording_Opened":"No","Test_Status":"Not Attempted","Score":"","Passing_Marks":int(tr["Passing_Marks"]),"Certificate_Status":"Not Issued","Certificate_Link":"","Due_Date":str(due),"Completed_On":"","Progress_%":0,"Remarks":"Assigned by Admin","Last_Updated":now()}
                        rec=pd.concat([rec,pd.DataFrame([row])],ignore_index=True)
                        log("Training Assigned",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=user_id,training_id=tid)
                        added+=1
                    write_sheet("Training_Records",rec); refresh_records(); st.success(f"{added} trainee(s) assigned.")
                    safe_session_update({"assign_selected": [], "assign_due_date": today()})
    with t4:
        st.warning("Admin-only database view.")
        for title,sheet in [("Dashboard","Dashboard"),("Users","Users"),("Trainings","Trainings"),("Training Records","Training_Records"),("Training Content","Training_Content"),("Question Bank","Question_Bank"),("Notifications","Notifications"),("Certificates","Certificates"),("Activity Log","Activity_Log")]:
            st.subheader(title); df=read_sheet(sheet); table(df.drop(columns=["Password_Hash"], errors="ignore") if sheet=="Users" else df)

def trainer_page(actor):
    st.header("Trainer Panel"); card("Trainer Work","Add links, upload content, generate MCQs, schedule training, mark attendance, and save recordings.")
    trainings=read_sheet("Trainings"); mine=trainings[trainings["Trainer_ID"]==actor_get(actor,"User_ID")]
    if mine.empty: st.warning("No training assigned."); return
    display=st.selectbox("Assigned Training",mine["Training_Title"].astype(str)+" — "+mine["Training_ID"].astype(str)); tid=display.split(" — ")[-1]; tr=mine[mine["Training_ID"]==tid].iloc[0]
    a,b,c,d,e=st.tabs(["Links","Content & MCQs","Schedule","Attendance/Recording","Results"])
    with a:
        with st.form("links"):
            slides=st.text_input("Slides Link",clean(tr["Slides_Link"])); video=st.text_input("Video Link",clean(tr["Video_Link"])); ref=st.text_input("Reference Link",clean(tr["Reference_Link"])); passing=st.number_input("Passing Marks (%)",1,100,int(tr["Passing_Marks"])); sub=st.form_submit_button("Save Links")
        if sub:
            if not all(url_ok(x) for x in [slides,video,ref]): st.error("Links must start with http:// or https://")
            else:
                trainings=read_sheet("Trainings"); idx=trainings[trainings["Training_ID"]==tid].index[0]; trainings.at[idx,"Slides_Link"]=slides.strip(); trainings.at[idx,"Video_Link"]=video.strip(); trainings.at[idx,"Reference_Link"]=ref.strip(); trainings.at[idx,"Passing_Marks"]=int(passing); trainings.at[idx,"Status"]="Material Added"; trainings.at[idx,"Last_Updated"]=now(); write_sheet("Trainings",trainings)
                rec=read_sheet("Training_Records"); rec.loc[rec["Training_ID"]==tid,"Passing_Marks"]=int(passing); write_sheet("Training_Records",rec); log("Training Links Saved",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),training_id=tid); st.success("Saved.")
    with b:
        check_extractors_panel()
        uploaded_files=st.file_uploader("Upload Training Content",type=["txt","doc","docx","pdf","ppt","pptx"], accept_multiple_files=True, help="Upload one or more TXT, DOC, DOCX, PDF, PPT or PPTX files."); count=st.slider("Number of MCQs",5,20,10)
        if st.button("Generate MCQs"):
            text, report = read_uploads(uploaded_files)
            if not report.empty:
                st.subheader("File Extraction Report")
                table(report)
            if not text.strip(): st.error("No readable content found. Check file type or click Check / Install File Extraction Packages.")
            else:
                for uploaded in uploaded_files:
                    single_text = read_upload(uploaded)
                    if single_text.strip():
                        append_row("Training_Content",{"Content_ID":uid("CONTENT"),"Training_ID":tid,"File_Name":uploaded.name,"File_Type":uploaded.name.split(".")[-1].lower(),"Content_Text":single_text[:30000],"Uploaded_By":actor_get(actor,"Name"),"Uploaded_On":now()})
                new=generate_mcqs(tid,text,count)
                if new.empty: st.error("Could not generate MCQs. Use clear technical sentences.")
                else:
                    q=read_sheet("Question_Bank"); q=pd.concat([q,new],ignore_index=True); write_sheet("Question_Bank",q); log("MCQs Generated",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),training_id=tid,remarks=f"{len(new)} questions"); update_dashboard(); st.success(f"{len(new)} MCQs generated."); table(new)
        q=read_sheet("Question_Bank"); st.subheader("Existing MCQs"); table(q[q["Training_ID"]==tid] if not q.empty else q)
    with c:
        st.subheader("Schedule Training and Create MS Teams Link")
        sdate = st.date_input("Schedule Date")
        stime = st.text_input("Schedule Time", "10:00 AM")

        generated_link = meeting_link(tid, clean(tr["Training_Title"]), str(sdate), stime)
        existing_link = clean(tr.get("Meeting_Link", ""))
        session_key = f"teams_link_{tid}"

        st.caption("Step 1: Click the button below to open Microsoft Teams and create the meeting. Step 2: Copy the final Teams join link from Teams and paste it in the Meeting Link field. Step 3: Click Schedule and Notify.")

        col_a, col_b = st.columns([1, 2])
        with col_a:
            st.link_button("Open MS Teams to Create Meeting", generated_link)
        with col_b:
            if st.button("Use Generated Teams Draft Link"):
                safe_session_update({session_key: generated_link})
                st.success("Generated Teams draft link placed in the Meeting Link field.")

        default_link = st.session_state.get(session_key, existing_link or generated_link)
        link = st.text_input(
            "Meeting Link",
            value=default_link,
            help="Paste the final Microsoft Teams join link here. The generated draft link opens Teams to create the meeting; the final join link is produced by Teams.",
        )

        if st.button("Schedule and Notify"):
            if not url_ok(link):
                st.error("Meeting link must start with http:// or https://")
            else:
                trainings = read_sheet("Trainings")
                idx = trainings[trainings["Training_ID"] == tid].index[0]
                trainings.at[idx, "Schedule_Date"] = str(sdate)
                trainings.at[idx, "Schedule_Time"] = stime
                trainings.at[idx, "Meeting_Link"] = link.strip()
                trainings.at[idx, "Status"] = "Scheduled"
                trainings.at[idx, "Last_Updated"] = now()
                write_sheet("Trainings", trainings)
                notify_training(tid, actor)
                update_dashboard()
                st.success("Scheduled and notifications generated. The MS Teams meeting link has been saved.")

        n = read_sheet("Notifications")
        table(n[n["Training_ID"] == tid] if not n.empty else n)
    with d:
        rec=read_sheet("Training_Records"); trainees=rec[rec["Training_ID"]==tid]
        if trainees.empty: st.warning("No trainees assigned.")
        else:
            table(trainees[["User_ID","Name","Live_Attendance","Status","Progress_%"]]); selected=st.selectbox("Trainee",trainees["Name"].astype(str)+" — "+trainees["User_ID"].astype(str)); trainee_id=selected.split(" — ")[-1]; att=st.selectbox("Attendance",["Present","Absent"])
            if st.button("Mark Attendance"):
                rec=read_sheet("Training_Records"); mask=(rec["Training_ID"]==tid)&(rec["User_ID"]==trainee_id); rec.loc[mask,"Live_Attendance"]=att; rec.loc[mask,"Remarks"]=f"Attendance marked: {att}"; rec.loc[mask,"Last_Updated"]=now(); write_sheet("Training_Records",rec); log("Attendance Marked",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=trainee_id,training_id=tid,remarks=att); refresh_records(); st.success("Attendance saved.")
        recording=st.text_input("Recording Link")
        if st.button("Save Recording Link"):
            if not url_ok(recording): st.error("Recording link must start with http:// or https://")
            else:
                trainings=read_sheet("Trainings"); idx=trainings[trainings["Training_ID"]==tid].index[0]; trainings.at[idx,"Recording_Link"]=recording.strip(); trainings.at[idx,"Status"]="Recorded"; trainings.at[idx,"Last_Updated"]=now(); write_sheet("Trainings",trainings); log("Recording Saved",actor_get(actor,"User_ID"),actor_get(actor,"Name"),actor_get(actor,"Role"),training_id=tid); update_dashboard(); st.success("Recording saved.")
    with e:
        rec=read_sheet("Training_Records"); table(rec[rec["Training_ID"]==tid] if not rec.empty else rec)

def trainee_page(actor):
    user_id=actor_get(actor,"User_ID"); role=actor_get(actor,"Role"); st.header(f"{role} Training Portal")
    rec=read_sheet("Training_Records"); mine=rec[rec["User_ID"]==user_id]
    if mine.empty: st.warning("No training assigned."); return
    display=st.selectbox("My Training",mine["Training_Title"].astype(str)+" — "+mine["Training_ID"].astype(str)); tid=display.split(" — ")[-1]; record=mine[mine["Training_ID"]==tid].iloc[0]
    trainings=read_sheet("Trainings"); tr=trainings[trainings["Training_ID"]==tid].iloc[0]
    cols=st.columns(5); cols[0].metric("Progress",f"{record['Progress_%']}%"); cols[1].metric("Slides",record["Slides_Opened"]); cols[2].metric("Video",record["Video_Opened"]); cols[3].metric("Test",record["Test_Status"]); cols[4].metric("Certificate",record["Certificate_Status"])
    x,y,z=st.tabs(["Training Material","MCQ Test","My Record"])
    with x:
        c1,c2,c3=st.columns(3)
        with c1:
            st.subheader("Slides"); st.code(clean(tr["Slides_Link"]));
            if st.button("Mark Slides Complete"): trainee_activity(actor,tid,"Slides_Opened","Slides completed.")
        with c2:
            st.subheader("Video"); st.code(clean(tr["Video_Link"]));
            if st.button("Mark Video Complete"): trainee_activity(actor,tid,"Video_Opened","Video completed.")
        with c3:
            st.subheader("Recording"); st.code(clean(tr["Recording_Link"]));
            if st.button("Mark Recording Complete"):
                rec=read_sheet("Training_Records"); mask=(rec["User_ID"]==user_id)&(rec["Training_ID"]==tid); rec.loc[mask,"Recording_Opened"]="Yes"; rec.loc[mask,"Video_Opened"]="Yes"; rec.loc[mask,"Live_Attendance"]="Recording Viewed"; rec.loc[mask,"Remarks"]="Recording viewed."; rec.loc[mask,"Last_Updated"]=now(); write_sheet("Training_Records",rec); log("Recording Viewed",user_id,actor_get(actor,"Name"),role,user_id=user_id,training_id=tid); refresh_records()
    with y:
        qs=read_sheet("Question_Bank"); questions=qs[qs["Training_ID"]==tid] if not qs.empty else pd.DataFrame()
        if questions.empty:
            st.warning("MCQ test has not been generated yet.")
        else:
            already_submitted = str(record.get("Test_Status","")).strip().lower() not in ["", "not attempted"]
            if already_submitted:
                st.info(f"MCQ already submitted. Score: {record.get('Score', '')}% — {record.get('Test_Status', '')}")
                st.write("You cannot resubmit the MCQ test after submission. Contact Admin if you need a review.")
            else:
                with st.form("mcq"):
                    answers={}
                    for i,(_,q) in enumerate(questions.iterrows(),1):
                        st.markdown(f"**Q{i}. {q['Question']}**")
                        opts=[q["Option_A"],q["Option_B"],q["Option_C"],q["Option_D"]]
                        answers[q["Question_ID"]]=st.radio("Select answer",opts,key=f"q_{q['Question_ID']}",label_visibility="collapsed")
                    sub=st.form_submit_button("Submit Test")
                if sub:
                    total=len(questions)
                    correct=sum(1 for _,q in questions.iterrows() if answers.get(q["Question_ID"])==q["Correct_Answer"])
                    score=round(correct/total*100,2) if total else 0
                    passing=int(record["Passing_Marks"] or 75)
                    result="Passed" if score>=passing else "Failed"
                    rec=read_sheet("Training_Records")
                    mask=(rec["User_ID"]==user_id)&(rec["Training_ID"]==tid)
                    rec.loc[mask,"Score"]=score
                    rec.loc[mask,"Test_Status"]=result
                    rec.loc[mask,"Remarks"]=f"MCQ submitted. Correct {correct}/{total}"
                    rec.loc[mask,"Last_Updated"]=now()
                    write_sheet("Training_Records",rec)
                    log("MCQ Test Submitted",user_id,actor_get(actor,"Name"),role,user_id=user_id,training_id=tid,remarks=f"{score}% {result}")
                    if result=="Passed": issue_certificate(actor,tid,int(score))
                    refresh_records(); st.success(f"Score: {score}% — {result}")
    with z:
        rec=read_sheet("Training_Records"); table(rec[rec["User_ID"]==user_id])

def trainee_activity(actor, tid, field, remarks):
    user_id=actor_get(actor,"User_ID"); rec=read_sheet("Training_Records"); mask=(rec["User_ID"]==user_id)&(rec["Training_ID"]==tid); rec.loc[mask,field]="Yes"; rec.loc[mask,"Remarks"]=remarks; rec.loc[mask,"Last_Updated"]=now(); write_sheet("Training_Records",rec); log(field.replace("_"," "),user_id,actor_get(actor,"Name"),actor_get(actor,"Role"),user_id=user_id,training_id=tid,remarks=remarks); refresh_records()

def management_page(actor):
    st.header("Management View"); st.info("Management can monitor progress. Full Excel download is Admin-only.")
    for title,sheet in [("Dashboard","Dashboard"),("Trainings","Trainings"),("Training Records","Training_Records"),("Certificates","Certificates")]: st.subheader(title); table(read_sheet(sheet))

def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
        menu_items={
            "Get help": "https://psbureau.org/help",
            "Report a bug": "mailto:support@psbureau.org",
            "About": f"{APP_TITLE} — {APP_SUBTITLE}"
        }
    )
    if not db_valid(): create_db(reset=True)
    style(); header(); actor=require_login(); sidebar(actor)
    st.markdown(f"<div class='card'><b>{actor_get(actor,'Name')}</b> | {actor_get(actor,'Role')} | {actor_get(actor,'Department')} | {actor_get(actor,'Email')}</div>", unsafe_allow_html=True)
    role=actor_get(actor,"Role")
    if role=="Admin":
        page=st.sidebar.radio("Menu",["Dashboard","Admin Panel"]); dashboard(actor) if page=="Dashboard" else admin_page(actor)
    elif role=="Trainer":
        page=st.sidebar.radio("Menu",["Dashboard","Trainer Panel"]); dashboard(actor) if page=="Dashboard" else trainer_page(actor)
    elif role in ["Surveyor","Plan Appraiser","Rule Development Rep","Trainee","On Probation"]:
        page=st.sidebar.radio("Menu",["Dashboard","Training Portal"]); dashboard(actor) if page=="Dashboard" else trainee_page(actor)
    elif role=="Quality Management Representative":
        page=st.sidebar.radio("Menu",["Dashboard","Quality Review"]); dashboard(actor) if page=="Dashboard" else management_page(actor)
    elif role=="Management":
        page=st.sidebar.radio("Menu",["Dashboard","Management View"]); dashboard(actor) if page=="Dashboard" else management_page(actor)
    else: st.error("Unknown role. Contact Admin.")
    if role=="Admin":
        with st.expander("Admin Activity Log"): table(read_sheet("Activity_Log"))
    footer()

if __name__ == "__main__": main()