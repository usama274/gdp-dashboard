
# ============================================================
# CLASSIFICATION SOCIETY TRAINING PLATFORM
# Complete Standalone Python Streamlit App + Excel Database
# ============================================================
# HOW TO RUN:
# 1) Save this file as: app.py
# 2) Install required libraries:
#       pip install streamlit pandas openpyxl
# 3) Run:
#       streamlit run app.py
#
# The app will automatically create an Excel database file:
#       classification_society_training_platform.xlsx
#
# It supports:
# - Management dashboard
# - Admin role assignment
# - Admin training creation and trainee assignment
# - Trainer material submission, schedule, notifications, attendance, recording
# - Surveyor / Plan Appraiser training activity tracking
# - Test submission and automatic certification
# - Excel auto-update after every activity
# ============================================================

import os
import uuid
import smtplib
from email.message import EmailMessage
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st

DB_FILE = "classification_society_training_platform.xlsx"

# ============================================================
# MODULE 1: BASIC HELPERS
# ============================================================

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def today():
    return date.today().strftime("%Y-%m-%d")

def make_id(prefix):
    return f"{prefix}-{uuid.uuid4().hex[:8].upper()}"

def file_exists():
    return Path(DB_FILE).exists()

# Safe session-state helpers (mirrors app.py) to avoid StreamlitAPIException
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

def read_sheet(sheet_name):
    try:
        return pd.read_excel(DB_FILE, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def save_sheet(sheet_name, df):
    """Save the given DataFrame to the Excel database.

    If the database file already exists, replace the sheet with the same name.
    Otherwise, create a new workbook and write the sheet.
    """
    if Path(DB_FILE).exists():
        with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def append_row(sheet_name, row):
    df = read_sheet(sheet_name)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    save_sheet(sheet_name, df)

def safe_str(value):
    if pd.isna(value):
        return ""
    return str(value)

# ============================================================
# MODULE 2: ROLE PERMISSIONS
# ============================================================

ROLE_PERMISSIONS = {
    "Management": [
        "view_dashboard",
        "view_all_records",
        "view_certificates",
        "view_notifications",
        "download_database",
    ],
    "Admin": [
        "view_dashboard",
        "manage_staff",
        "assign_roles",
        "create_training",
        "assign_trainer",
        "assign_trainees",
        "view_all_records",
        "view_notifications",
        "view_certificates",
        "download_database",
    ],
    "Trainer": [
        "view_trainer_trainings",
        "update_training_fields",
        "submit_training_schedule",
        "generate_notifications",
        "mark_attendance",
        "upload_recording",
        "review_results",
    ],
    "Surveyor": [
        "view_own_training",
        "open_slides",
        "open_video",
        "open_recording",
        "take_test",
        "view_own_certificate",
    ],
    "Plan Appraiser": [
        "view_own_training",
        "open_slides",
        "open_video",
        "open_recording",
        "take_test",
        "view_own_certificate",
    ],
}

def allowed(role, action):
    return action in ROLE_PERMISSIONS.get(role, [])

def require_allowed(role, action):
    if not allowed(role, action):
        st.error(f"{role} is not allowed to perform: {action}")
        return False
    return True

# ============================================================
# MODULE 3: LOGGING, DASHBOARD, PROGRESS
# ============================================================

def log_activity(
    activity,
    actor_id="",
    actor_name="",
    actor_role="",
    staff_id="",
    training_id="",
    status="Success",
    remarks="",
):
    append_row("Activity_Log", {
        "Log_ID": make_id("LOG"),
        "Date_Time": now(),
        "Activity": activity,
        "Actor_ID": actor_id,
        "Actor_Name": actor_name,
        "Actor_Role": actor_role,
        "Staff_ID": staff_id,
        "Training_ID": training_id,
        "Status": status,
        "Remarks": remarks,
    })
    update_dashboard()

def calculate_progress(record):
    steps = [
        record.get("Slides_Opened") == "Yes",
        record.get("Video_Opened") == "Yes" or record.get("Recording_Opened") == "Yes",
        record.get("Live_Attendance") in ["Present", "Recording Viewed"],
        record.get("Test_Status") == "Passed",
        record.get("Certificate_Status") == "Issued",
    ]
    return int((sum(steps) / len(steps)) * 100)

def refresh_training_records():
    records = read_sheet("Training_Records")
    if records.empty:
        update_dashboard()
        return

    for idx, row in records.iterrows():
        progress = calculate_progress(row)
        records.at[idx, "Progress_%"] = progress

        if progress >= 100:
            records.at[idx, "Status"] = "Completed"
            if safe_str(row.get("Completed_On")).strip() == "":
                records.at[idx, "Completed_On"] = today()
        else:
            records.at[idx, "Status"] = "Pending"

        records.at[idx, "Last_Updated"] = now()

    save_sheet("Training_Records", records)
    update_dashboard()

def update_dashboard():
    staff = read_sheet("Staff")
    trainings = read_sheet("Trainings")
    records = read_sheet("Training_Records")
    certs = read_sheet("Certificates")
    logs = read_sheet("Activity_Log")

    if records.empty:
        completed = pending = avg_progress = 0
    else:
        completed = len(records[records["Status"] == "Completed"])
        pending = len(records[records["Status"] != "Completed"])
        avg_progress = round(records["Progress_%"].fillna(0).mean(), 2)

    dashboard = pd.DataFrame([
        ["Total Staff", len(staff)],
        ["Surveyors", len(staff[staff["Role"] == "Surveyor"]) if not staff.empty else 0],
        ["Plan Appraisers", len(staff[staff["Role"] == "Plan Appraiser"]) if not staff.empty else 0],
        ["Trainers", len(staff[staff["Role"] == "Trainer"]) if not staff.empty else 0],
        ["Total Trainings", len(trainings)],
        ["Completed Training Records", completed],
        ["Pending Training Records", pending],
        ["Average Completion %", avg_progress],
        ["Certificates Issued", len(certs)],
        ["Activities Logged", len(logs)],
        ["Last Updated", now()],
    ], columns=["Metric", "Value"])

    save_sheet("Dashboard", dashboard)

# ============================================================
# MODULE 4: DATABASE CREATION
# ============================================================

def create_database(reset=False):
    if file_exists() and not reset:
        return

    staff = pd.DataFrame([
        ["STF-ADMIN", "Admin User", "Admin", "Administration", "admin@psbureau.org", "Active"],
        ["STF-MGMT", "Management User", "Management", "Management", "management@psbureau.org", "Active"],
        ["STF-TRN-001", "Usama Saleem", "Trainer", "Training Department", "trainer@psbureau.org", "Active"],
        ["STF-SUR-001", "Muhammad Ali", "Surveyor", "Electrical Survey", "ali@psbureau.org", "Active"],
        ["STF-SUR-002", "Sara Malik", "Surveyor", "Hull & Machinery Survey", "sara@psbureau.org", "Active"],
        ["STF-APP-001", "Ahmed Khan", "Plan Appraiser", "Plan Appraisal", "ahmed@psbureau.org", "Active"],
    ], columns=["Staff_ID", "Name", "Role", "Department", "Email", "Status"])

    trainings = pd.DataFrame(columns=[
        "Training_ID",
        "Training_Title",
        "Category",
        "Target_Role",
        "Trainer_ID",
        "Trainer_Name",
        "Slides_Link",
        "Video_Link",
        "Reference_Link",
        "Schedule_Date",
        "Schedule_Time",
        "Meeting_Link",
        "Recording_Link",
        "Passing_Marks",
        "Status",
        "Submitted_By_Trainer",
        "Created_By",
        "Created_On",
        "Last_Updated",
    ])

    records = pd.DataFrame(columns=[
        "Record_ID",
        "Staff_ID",
        "Staff_Name",
        "Role",
        "Training_ID",
        "Training_Title",
        "Status",
        "Slides_Opened",
        "Video_Opened",
        "Live_Attendance",
        "Recording_Opened",
        "Test_Status",
        "Score",
        "Passing_Marks",
        "Certificate_Status",
        "Certificate_Link",
        "Due_Date",
        "Completed_On",
        "Progress_%",
        "Remarks",
        "Last_Updated",
    ])

    tests = pd.DataFrame(columns=[
        "Question_ID",
        "Training_ID",
        "Question",
        "Option_A",
        "Option_B",
        "Option_C",
        "Option_D",
        "Correct_Answer",
        "Marks",
    ])

    notifications = pd.DataFrame(columns=[
        "Notification_ID",
        "Training_ID",
        "Staff_ID",
        "Staff_Name",
        "Staff_Email",
        "Subject",
        "Message",
        "Status",
        "Generated_On",
        "Sent_On",
        "Generated_By",
    ])

    certificates = pd.DataFrame(columns=[
        "Certificate_ID",
        "Staff_ID",
        "Staff_Name",
        "Role",
        "Training_ID",
        "Training_Title",
        "Score",
        "Issued_On",
        "Certificate_Link",
        "Status",
        "Issued_By",
    ])

    logs = pd.DataFrame(columns=[
        "Log_ID",
        "Date_Time",
        "Activity",
        "Actor_ID",
        "Actor_Name",
        "Actor_Role",
        "Staff_ID",
        "Training_ID",
        "Status",
        "Remarks",
    ])

    role_permissions = []
    for role, actions in ROLE_PERMISSIONS.items():
        for action in actions:
            role_permissions.append([role, action, "Allowed"])
    role_permissions = pd.DataFrame(role_permissions, columns=["Role", "Action", "Permission"])

    dashboard = pd.DataFrame(columns=["Metric", "Value"])

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        staff.to_excel(writer, sheet_name="Staff", index=False)
        trainings.to_excel(writer, sheet_name="Trainings", index=False)
        records.to_excel(writer, sheet_name="Training_Records", index=False)
        tests.to_excel(writer, sheet_name="Tests", index=False)
        notifications.to_excel(writer, sheet_name="Notifications", index=False)
        certificates.to_excel(writer, sheet_name="Certificates", index=False)
        logs.to_excel(writer, sheet_name="Activity_Log", index=False)
        role_permissions.to_excel(writer, sheet_name="Role_Permissions", index=False)
        dashboard.to_excel(writer, sheet_name="Dashboard", index=False)

    log_activity("Database Created", remarks="Initial Excel training platform database created.")
    update_dashboard()

# ============================================================
# MODULE 5: BUSINESS LOGIC
# ============================================================

def generate_meeting_link(training_id):
    return f"https://teams.microsoft.com/l/meetup-join/{training_id}"

def generate_certificate_link(staff_id, training_id):
    return f"https://certificate.psbureau.org/{staff_id}/{training_id}"

def get_training_by_id(training_id):
    trainings = read_sheet("Trainings")
    if trainings.empty:
        return {}
    match = trainings[trainings["Training_ID"] == training_id]
    if match.empty:
        return {}
    return match.iloc[0].to_dict()

def generate_notifications_for_training(training_id, actor):
    trainings = read_sheet("Trainings")
    records = read_sheet("Training_Records")
    staff = read_sheet("Staff")

    if trainings.empty or records.empty:
        return

    training = trainings[trainings["Training_ID"] == training_id]
    if training.empty:
        return

    training_row = training.iloc[0]
    assigned = records[records["Training_ID"] == training_id]

    for _, rec in assigned.iterrows():
        staff_row = staff[staff["Staff_ID"] == rec["Staff_ID"]]
        if staff_row.empty:
            continue
        staff_row = staff_row.iloc[0]

        subject = f"Training Assigned: {training_row['Training_Title']}"
        message = (
            f"Dear {staff_row['Name']},\n\n"
            f"You have been assigned the following training:\n\n"
            f"Training: {training_row['Training_Title']}\n"
            f"Category: {training_row['Category']}\n"
            f"Trainer: {training_row['Trainer_Name']}\n"
            f"Schedule: {training_row['Schedule_Date']} at {training_row['Schedule_Time']}\n"
            f"Meeting Link: {training_row['Meeting_Link']}\n"
            f"Slides Link: {training_row['Slides_Link']}\n"
            f"Video Link: {training_row['Video_Link']}\n"
            f"Passing Marks: {training_row['Passing_Marks']}\n\n"
            f"Please complete the assigned activities and take the test.\n\n"
            f"Regards,\nTraining Department"
        )

        append_row("Notifications", {
            "Notification_ID": make_id("NOTIF"),
            "Training_ID": training_id,
            "Staff_ID": staff_row["Staff_ID"],
            "Staff_Name": staff_row["Name"],
            "Staff_Email": staff_row["Email"],
            "Subject": subject,
            "Message": message,
            "Status": "Generated",
            "Generated_On": now(),
            "Sent_On": "",
            "Generated_By": actor["Name"],
        })

    log_activity(
        "Notifications Generated",
        actor_id=actor["Staff_ID"],
        actor_name=actor["Name"],
        actor_role=actor["Role"],
        training_id=training_id,
        remarks="Email notification records generated in Excel.",
    )

def send_email_if_configured(to_email, subject, body):
    """
    Optional real email sending.
    Configure environment variables:
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_FROM
    If not configured, the app still generates email records in Excel.
    """
    host = os.getenv("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    sender = os.getenv("SMTP_FROM", user)

    if not host or not user or not password or not sender:
        return False, "SMTP not configured. Email generated in Excel only."

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    with smtplib.SMTP(host, port) as smtp:
        smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)

    return True, "Email sent."

# ============================================================
# MODULE 6: UI HELPERS
# ============================================================

def login_sidebar():
    st.sidebar.title("Login / Role")

    staff = read_sheet("Staff")
    if staff.empty:
        st.error("No staff database found.")
        st.stop()

    staff["Display"] = staff["Name"] + " — " + staff["Role"]
    selected = st.sidebar.selectbox("Select User", staff["Display"].tolist())
    row = staff[staff["Display"] == selected].iloc[0].to_dict()

    try:
        safe_session_update({"actor": row})
    except Exception:
        pass

    st.sidebar.success(f"Logged in as: {row['Role']}")
    st.sidebar.caption(row["Email"])
    return row

def show_download_button():
    if Path(DB_FILE).exists():
        with open(DB_FILE, "rb") as f:
            st.sidebar.download_button(
                "Download Excel Database",
                f,
                file_name=DB_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

def show_table(title, df):
    st.subheader(title)
    if df.empty:
        st.info("No records available.")
    else:
        st.dataframe(df, width="stretch")

# ============================================================
# MODULE 7: MANAGEMENT VIEW
# ============================================================

def management_view(actor):
    st.header("Management Dashboard")
    st.caption("Management can view complete training status, pending records, certificates, notifications, and activity log.")

    dashboard = read_sheet("Dashboard")
    if not dashboard.empty:
        cols = st.columns(4)
        for i, row in dashboard.iterrows():
            with cols[i % 4]:
                st.metric(row["Metric"], row["Value"])

    show_table("All Training Records", read_sheet("Training_Records"))
    show_table("Certificates", read_sheet("Certificates"))
    show_table("Notifications", read_sheet("Notifications"))
    show_table("Activity Log", read_sheet("Activity_Log"))

# ============================================================
# MODULE 8: ADMIN VIEW
# ============================================================

def admin_view(actor):
    st.header("Admin Panel")
    st.caption("Admin assigns roles, creates trainings, assigns trainers, and assigns surveyors/plan appraisers.")

    tab1, tab2, tab3, tab4 = st.tabs(["Staff & Roles", "Create Training", "Assign Trainees", "All Records"])

    with tab1:
        show_table("Current Staff", read_sheet("Staff"))

        with st.form("add_staff_form"):
            st.markdown("### Add / Update Staff Role")
            staff_id = st.text_input("Staff ID", value=make_id("STF"))
            name = st.text_input("Name")
            role = st.selectbox("Role", ["Admin", "Management", "Trainer", "Surveyor", "Plan Appraiser"])
            department = st.text_input("Department")
            email = st.text_input("Email")
            submitted = st.form_submit_button("Save Staff / Role")

            if submitted:
                if not require_allowed(actor["Role"], "manage_staff"):
                    return
                if not name or not email:
                    st.error("Name and email are required.")
                    return

                staff = read_sheet("Staff")
                existing = staff[staff["Staff_ID"] == staff_id]

                new_row = {
                    "Staff_ID": staff_id,
                    "Name": name,
                    "Role": role,
                    "Department": department,
                    "Email": email,
                    "Status": "Active",
                }

                if existing.empty:
                    staff = pd.concat([staff, pd.DataFrame([new_row])], ignore_index=True)
                    action = "Staff Added"
                else:
                    idx = existing.index[0]
                    for k, v in new_row.items():
                        staff.at[idx, k] = v
                    action = "Staff Role Updated"

                save_sheet("Staff", staff)
                log_activity(action, actor["Staff_ID"], actor["Name"], actor["Role"], staff_id=staff_id, remarks=f"{name} assigned as {role}")
                st.success("Staff saved and Excel updated.")
                st.rerun()

    with tab2:
        staff = read_sheet("Staff")
        trainers = staff[staff["Role"] == "Trainer"]

        with st.form("create_training_form"):
            st.markdown("### Create Training and Assign Trainer")
            training_id = st.text_input("Training ID", value=make_id("TRN"))
            title = st.text_input("Training Title")
            category = st.selectbox("Category", ["Basic Survey", "Electrical Survey", "Hull Survey", "Machinery Survey", "Plan Appraisal", "Statutory Survey", "Report Writing"])
            target_role = st.selectbox("Target Role", ["Surveyor", "Plan Appraiser"])
            trainer_options = (trainers["Name"] + " — " + trainers["Staff_ID"]).tolist() if not trainers.empty else []
            trainer_display = st.selectbox("Assign Trainer", trainer_options)
            passing_marks = st.number_input("Passing Marks", min_value=0, max_value=100, value=75)
            create_btn = st.form_submit_button("Create Training")

            if create_btn:
                if not require_allowed(actor["Role"], "create_training"):
                    return
                if not title:
                    st.error("Training title is required.")
                    return
                if not trainer_display:
                    st.error("Please create/assign at least one trainer first.")
                    return

                trainer_id = trainer_display.split(" — ")[-1]
                trainer_name = trainer_display.split(" — ")[0]

                trainings = read_sheet("Trainings")
                new = {
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
                    "Passing_Marks": passing_marks,
                    "Status": "Draft",
                    "Submitted_By_Trainer": "No",
                    "Created_By": actor["Name"],
                    "Created_On": now(),
                    "Last_Updated": now(),
                }

                trainings = pd.concat([trainings, pd.DataFrame([new])], ignore_index=True)
                save_sheet("Trainings", trainings)

                log_activity("Training Created", actor["Staff_ID"], actor["Name"], actor["Role"], training_id=training_id, remarks=f"Trainer assigned: {trainer_name}")
                st.success("Training created and trainer assigned.")
                st.rerun()

    with tab3:
        trainings = read_sheet("Trainings")
        staff = read_sheet("Staff")

        if trainings.empty:
            st.warning("No training created yet.")
            return

        training_display = st.selectbox("Select Training", trainings["Training_Title"] + " — " + trainings["Training_ID"])
        selected_training_id = training_display.split(" — ")[-1]
        training = trainings[trainings["Training_ID"] == selected_training_id].iloc[0]
        target_role = training["Target_Role"]

        eligible = staff[staff["Role"] == target_role]
        selected_staff = st.multiselect("Select Trainees", eligible["Name"] + " — " + eligible["Staff_ID"])
        due_date = st.date_input("Due Date")

        if st.button("Assign Selected Trainees"):
            if not require_allowed(actor["Role"], "assign_trainees"):
                return

            records = read_sheet("Training_Records")

            for item in selected_staff:
                staff_id = item.split(" — ")[-1]
                staff_row = staff[staff["Staff_ID"] == staff_id].iloc[0]

                exists = records[(records["Staff_ID"] == staff_id) & (records["Training_ID"] == selected_training_id)]
                if not exists.empty:
                    continue

                new_record = {
                    "Record_ID": make_id("REC"),
                    "Staff_ID": staff_id,
                    "Staff_Name": staff_row["Name"],
                    "Role": staff_row["Role"],
                    "Training_ID": selected_training_id,
                    "Training_Title": training["Training_Title"],
                    "Status": "Pending",
                    "Slides_Opened": "No",
                    "Video_Opened": "No",
                    "Live_Attendance": "Not Marked",
                    "Recording_Opened": "No",
                    "Test_Status": "Not Attempted",
                    "Score": "",
                    "Passing_Marks": training["Passing_Marks"],
                    "Certificate_Status": "Not Issued",
                    "Certificate_Link": "",
                    "Due_Date": str(due_date),
                    "Completed_On": "",
                    "Progress_%": 0,
                    "Remarks": "Assigned by Admin.",
                    "Last_Updated": now(),
                }

                records = pd.concat([records, pd.DataFrame([new_record])], ignore_index=True)
                log_activity("Trainee Assigned", actor["Staff_ID"], actor["Name"], actor["Role"], staff_id=staff_id, training_id=selected_training_id, remarks=f"Assigned to {training['Training_Title']}")

            save_sheet("Training_Records", records)
            refresh_training_records()
            st.success("Trainees assigned and Excel updated.")
            st.rerun()

    with tab4:
        show_table("All Training Records", read_sheet("Training_Records"))

# ============================================================
# MODULE 9: TRAINER VIEW
# ============================================================

def trainer_view(actor):
    st.header("Trainer Panel")
    st.caption("Trainer updates material, submits schedule, generates notifications, marks attendance, uploads recording, and reviews results.")

    trainings = read_sheet("Trainings")
    my_trainings = trainings[trainings["Trainer_ID"] == actor["Staff_ID"]] if not trainings.empty else pd.DataFrame()

    if my_trainings.empty:
        st.warning("No training assigned to this trainer yet.")
        return

    training_display = st.selectbox("Select Assigned Training", my_trainings["Training_Title"] + " — " + my_trainings["Training_ID"])
    training_id = training_display.split(" — ")[-1]
    training = my_trainings[my_trainings["Training_ID"] == training_id].iloc[0]

    tab1, tab2, tab3, tab4 = st.tabs(["Training Fields", "Schedule & Notify", "Attendance & Recording", "Results"])

    with tab1:
        with st.form("trainer_update_form"):
            st.markdown("### Update Training Material")
            slides_link = st.text_input("Slides Link", value=safe_str(training.get("Slides_Link", "")))
            video_link = st.text_input("Training Video Link", value=safe_str(training.get("Video_Link", "")))
            reference_link = st.text_input("Reference Material Link", value=safe_str(training.get("Reference_Link", "")))
            passing_marks = st.number_input("Passing Marks", 0, 100, int(training.get("Passing_Marks", 75) or 75))
            submitted = st.form_submit_button("Submit Training Material")

            if submitted:
                if not require_allowed(actor["Role"], "update_training_fields"):
                    return

                all_trainings = read_sheet("Trainings")
                idx = all_trainings[all_trainings["Training_ID"] == training_id].index[0]
                all_trainings.at[idx, "Slides_Link"] = slides_link
                all_trainings.at[idx, "Video_Link"] = video_link
                all_trainings.at[idx, "Reference_Link"] = reference_link
                all_trainings.at[idx, "Passing_Marks"] = passing_marks
                all_trainings.at[idx, "Submitted_By_Trainer"] = "Yes"
                all_trainings.at[idx, "Status"] = "Material Submitted"
                all_trainings.at[idx, "Last_Updated"] = now()
                save_sheet("Trainings", all_trainings)

                records = read_sheet("Training_Records")
                if not records.empty:
                    records.loc[records["Training_ID"] == training_id, "Passing_Marks"] = passing_marks
                    save_sheet("Training_Records", records)

                log_activity("Training Material Submitted", actor["Staff_ID"], actor["Name"], actor["Role"], training_id=training_id, remarks="Slides/video/reference/passing marks updated.")
                st.success("Training material submitted and Excel updated.")
                st.rerun()

    with tab2:
        st.markdown("### Schedule Training and Generate Notifications")
        schedule_date = st.date_input("Schedule Date")
        schedule_time = st.text_input("Schedule Time", value="10:00 AM")
        meeting_link = st.text_input("Meeting Link", value=generate_meeting_link(training_id))

        if st.button("Submit Schedule & Generate Notifications"):
            if not require_allowed(actor["Role"], "submit_training_schedule"):
                return

            all_trainings = read_sheet("Trainings")
            idx = all_trainings[all_trainings["Training_ID"] == training_id].index[0]
            all_trainings.at[idx, "Schedule_Date"] = str(schedule_date)
            all_trainings.at[idx, "Schedule_Time"] = schedule_time
            all_trainings.at[idx, "Meeting_Link"] = meeting_link
            all_trainings.at[idx, "Status"] = "Scheduled"
            all_trainings.at[idx, "Last_Updated"] = now()
            save_sheet("Trainings", all_trainings)

            generate_notifications_for_training(training_id, actor)
            log_activity("Training Scheduled", actor["Staff_ID"], actor["Name"], actor["Role"], training_id=training_id, remarks="Schedule submitted and notifications generated.")
            st.success("Schedule updated. Notifications generated in Excel.")
            st.rerun()

        notifs = read_sheet("Notifications")
        show_table("Generated Notifications", notifs[notifs["Training_ID"] == training_id] if not notifs.empty else pd.DataFrame())

        if st.button("Mark Generated Emails as Sent"):
            notifs = read_sheet("Notifications")
            if not notifs.empty:
                mask = notifs["Training_ID"] == training_id
                notifs.loc[mask, "Status"] = "Sent"
                notifs.loc[mask, "Sent_On"] = now()
                save_sheet("Notifications", notifs)
                log_activity("Emails Marked Sent", actor["Staff_ID"], actor["Name"], actor["Role"], training_id=training_id, remarks="Notification emails marked as sent.")
                st.success("Email status updated in Excel.")
                st.rerun()

    with tab3:
        records = read_sheet("Training_Records")
        training_records = records[records["Training_ID"] == training_id] if not records.empty else pd.DataFrame()

        show_table("Assigned Trainees", training_records)

        if not training_records.empty:
            selected_trainee = st.selectbox("Select Trainee for Attendance", training_records["Staff_Name"] + " — " + training_records["Staff_ID"])
            staff_id = selected_trainee.split(" — ")[-1]
            attendance = st.selectbox("Attendance", ["Present", "Absent"])

            if st.button("Mark Attendance"):
                if not require_allowed(actor["Role"], "mark_attendance"):
                    return

                records = read_sheet("Training_Records")
                mask = (records["Training_ID"] == training_id) & (records["Staff_ID"] == staff_id)
                records.loc[mask, "Live_Attendance"] = attendance
                records.loc[mask, "Remarks"] = f"Attendance marked by trainer: {attendance}"
                records.loc[mask, "Last_Updated"] = now()
                save_sheet("Training_Records", records)

                log_activity("Attendance Marked", actor["Staff_ID"], actor["Name"], actor["Role"], staff_id=staff_id, training_id=training_id, remarks=attendance)
                refresh_training_records()
                st.success("Attendance updated.")
                st.rerun()

        recording_link = st.text_input("Recording Link after Training Completion")
        if st.button("Save Recording Link"):
            if not require_allowed(actor["Role"], "upload_recording"):
                return

            all_trainings = read_sheet("Trainings")
            idx = all_trainings[all_trainings["Training_ID"] == training_id].index[0]
            all_trainings.at[idx, "Recording_Link"] = recording_link
            all_trainings.at[idx, "Status"] = "Recorded"
            all_trainings.at[idx, "Last_Updated"] = now()
            save_sheet("Trainings", all_trainings)

            records = read_sheet("Training_Records")
            if not records.empty:
                records.loc[records["Training_ID"] == training_id, "Remarks"] = "Recording available."
                records.loc[records["Training_ID"] == training_id, "Last_Updated"] = now()
                save_sheet("Training_Records", records)

            log_activity("Recording Uploaded", actor["Staff_ID"], actor["Name"], actor["Role"], training_id=training_id, remarks=recording_link)
            st.success("Recording link saved and Excel updated.")
            st.rerun()

    with tab4:
        records = read_sheet("Training_Records")
        show_table("Training Results", records[records["Training_ID"] == training_id] if not records.empty else pd.DataFrame())

# ============================================================
# MODULE 10: TRAINEE VIEW
# ============================================================

def update_trainee_activity(actor, training_id, field, value, remarks):
    records = read_sheet("Training_Records")
    mask = (records["Staff_ID"] == actor["Staff_ID"]) & (records["Training_ID"] == training_id)

    if records[mask].empty:
        st.error("Training record not found.")
        return

    records.loc[mask, field] = value
    records.loc[mask, "Remarks"] = remarks
    records.loc[mask, "Last_Updated"] = now()
    save_sheet("Training_Records", records)

    log_activity(
        field.replace("_", " "),
        actor["Staff_ID"],
        actor["Name"],
        actor["Role"],
        staff_id=actor["Staff_ID"],
        training_id=training_id,
        remarks=remarks,
    )

    refresh_training_records()
    st.success("Activity updated in Excel.")
    st.rerun()

def issue_certificate_auto(actor, training_id, score):
    records = read_sheet("Training_Records")
    certs = read_sheet("Certificates")

    mask = (records["Staff_ID"] == actor["Staff_ID"]) & (records["Training_ID"] == training_id)
    if records[mask].empty:
        return

    row = records[mask].iloc[0]
    cert_link = generate_certificate_link(actor["Staff_ID"], training_id)

    records.loc[mask, "Certificate_Status"] = "Issued"
    records.loc[mask, "Certificate_Link"] = cert_link
    records.loc[mask, "Remarks"] = "Certificate automatically issued after passing test."
    records.loc[mask, "Last_Updated"] = now()
    save_sheet("Training_Records", records)

    exists = certs[(certs["Staff_ID"] == actor["Staff_ID"]) & (certs["Training_ID"] == training_id)] if not certs.empty else pd.DataFrame()
    if exists.empty:
        certs = pd.concat([certs, pd.DataFrame([{
            "Certificate_ID": make_id("CERT"),
            "Staff_ID": actor["Staff_ID"],
            "Staff_Name": actor["Name"],
            "Role": actor["Role"],
            "Training_ID": training_id,
            "Training_Title": row["Training_Title"],
            "Score": score,
            "Issued_On": today(),
            "Certificate_Link": cert_link,
            "Status": "Issued",
            "Issued_By": "System Auto Certification",
        }])], ignore_index=True)
        save_sheet("Certificates", certs)

    log_activity("Certificate Auto Issued", actor["Staff_ID"], actor["Name"], actor["Role"], staff_id=actor["Staff_ID"], training_id=training_id, remarks=cert_link)
    refresh_training_records()

def trainee_view(actor):
    st.header(f"{actor['Role']} Training Portal")
    st.caption("Open slides/video/recording, complete activity, take test, and receive certificate after passing.")

    records = read_sheet("Training_Records")
    my_records = records[records["Staff_ID"] == actor["Staff_ID"]] if not records.empty else pd.DataFrame()

    if my_records.empty:
        st.warning("No training assigned yet.")
        return

    training_display = st.selectbox("Select Training", my_records["Training_Title"] + " — " + my_records["Training_ID"])
    training_id = training_display.split(" — ")[-1]

    records = read_sheet("Training_Records")
    record = records[(records["Staff_ID"] == actor["Staff_ID"]) & (records["Training_ID"] == training_id)].iloc[0]
    training = get_training_by_id(training_id)

    st.subheader("Training Status")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Progress", f"{record['Progress_%']}%")
    c2.metric("Slides", record["Slides_Opened"])
    c3.metric("Video", record["Video_Opened"])
    c4.metric("Test", record["Test_Status"])
    c5.metric("Certificate", record["Certificate_Status"])

    st.markdown("### Training Material")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.write("Slides")
        st.code(safe_str(training.get("Slides_Link", "")) or "No slides link submitted yet.")
        if st.button("Open Slides / Mark Done"):
            update_trainee_activity(actor, training_id, "Slides_Opened", "Yes", "Slides opened by trainee.")

    with col2:
        st.write("Video")
        st.code(safe_str(training.get("Video_Link", "")) or "No video link submitted yet.")
        if st.button("Open Video / Mark Done"):
            update_trainee_activity(actor, training_id, "Video_Opened", "Yes", "Video opened by trainee.")

    with col3:
        st.write("Recording")
        st.code(safe_str(training.get("Recording_Link", "")) or "No recording link available yet.")
        if st.button("Open Recording / Mark Done"):
            records = read_sheet("Training_Records")
            mask = (records["Staff_ID"] == actor["Staff_ID"]) & (records["Training_ID"] == training_id)
            records.loc[mask, "Recording_Opened"] = "Yes"
            records.loc[mask, "Video_Opened"] = "Yes"
            records.loc[mask, "Live_Attendance"] = "Recording Viewed"
            records.loc[mask, "Remarks"] = "Recording viewed by trainee."
            records.loc[mask, "Last_Updated"] = now()
            save_sheet("Training_Records", records)

            log_activity("Recording Opened", actor["Staff_ID"], actor["Name"], actor["Role"], staff_id=actor["Staff_ID"], training_id=training_id, remarks="Recording opened and marked complete.")
            refresh_training_records()
            st.success("Recording activity updated.")
            st.rerun()

    st.markdown("### Test / Assessment")
    passing_marks = int(record["Passing_Marks"]) if safe_str(record["Passing_Marks"]).strip() else 75
    score = st.number_input("Enter Test Score", min_value=0, max_value=100, value=passing_marks)

    if st.button("Submit Test"):
        records = read_sheet("Training_Records")
        mask = (records["Staff_ID"] == actor["Staff_ID"]) & (records["Training_ID"] == training_id)

        status = "Passed" if score >= passing_marks else "Failed"
        records.loc[mask, "Score"] = score
        records.loc[mask, "Test_Status"] = status
        records.loc[mask, "Remarks"] = f"Test submitted. Score {score}. Passing marks {passing_marks}."
        records.loc[mask, "Last_Updated"] = now()
        save_sheet("Training_Records", records)

        log_activity("Test Submitted", actor["Staff_ID"], actor["Name"], actor["Role"], staff_id=actor["Staff_ID"], training_id=training_id, remarks=f"Score: {score}, Result: {status}")
        refresh_training_records()

        if status == "Passed":
            issue_certificate_auto(actor, training_id, score)
            st.success("Test passed. Certificate issued.")
        else:
            st.error("Test failed. Certificate not issued.")

        st.rerun()

    show_table("My Training Record", my_records)

# ============================================================
# MODULE 11: NOTIFICATION / EMAIL CENTER
# ============================================================

def notification_center(actor):
    st.header("Notification / Email Center")
    st.caption("Emails are generated in Excel. If SMTP is configured, they can also be sent.")

    notifications = read_sheet("Notifications")
    show_table("Notifications", notifications)

    if notifications.empty:
        return

    if st.button("Send All Generated Emails"):
        notifications = read_sheet("Notifications")
        sent_count = 0

        for idx, row in notifications.iterrows():
            if row["Status"] not in ["Generated", "Failed"]:
                continue

            sent, message = send_email_if_configured(row["Staff_Email"], row["Subject"], row["Message"])
            notifications.at[idx, "Status"] = "Sent" if sent else "Generated"
            notifications.at[idx, "Sent_On"] = now() if sent else ""
            sent_count += 1 if sent else 0

        save_sheet("Notifications", notifications)
        log_activity("Email Sending Attempted", actor["Staff_ID"], actor["Name"], actor["Role"], remarks=f"Sent count: {sent_count}")
        st.success("Email process completed. Excel updated.")
        st.rerun()

# ============================================================
# MODULE 12: MAIN APP
# ============================================================

def main():
    st.set_page_config(page_title="Classification Society Training Platform", layout="wide")

    if not file_exists():
        create_database(reset=True)

    st.title("Classification Society Training Platform")
    st.caption("Python + Streamlit app with Excel database for training, notifications, schedules, tests, and certification.")

    actor = login_sidebar()

    st.sidebar.divider()

    if st.sidebar.button("Reset / Recreate Excel Database"):
        create_database(reset=True)
        st.sidebar.success("Database recreated.")
        st.rerun()

    show_download_button()

    role = actor["Role"]

    if role == "Management":
        page = st.sidebar.radio("Management Menu", ["Dashboard", "Notifications"])
        if page == "Dashboard":
            management_view(actor)
        else:
            notification_center(actor)

    elif role == "Admin":
        page = st.sidebar.radio("Admin Menu", ["Admin Panel", "Management View", "Notifications"])
        if page == "Admin Panel":
            admin_view(actor)
        elif page == "Management View":
            management_view(actor)
        else:
            notification_center(actor)

    elif role == "Trainer":
        page = st.sidebar.radio("Trainer Menu", ["Trainer Panel", "Notifications"])
        if page == "Trainer Panel":
            trainer_view(actor)
        else:
            notification_center(actor)

    elif role in ["Surveyor", "Plan Appraiser"]:
        trainee_view(actor)

    else:
        st.error("Unknown role.")

    st.divider()
    st.subheader("Excel Database Preview")
    with st.expander("Dashboard"):
        st.dataframe(read_sheet("Dashboard"), width="stretch")
    with st.expander("Activity Log"):
        st.dataframe(read_sheet("Activity_Log"), width="stretch")
    with st.expander("Role Permissions"):
        st.dataframe(read_sheet("Role_Permissions"), width="stretch")

if __name__ == "__main__":
    main()