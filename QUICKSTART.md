# 🎓 Classification Society Training Platform - Setup & User Guide

## ✅ Quick Start (Codespaces)

### Option 1: Auto-Start (Recommended)
The app will start automatically when you open Codespaces. Just wait a few seconds and look for the port notification.

### Option 2: Manual Start
If the app isn't running, open a terminal and run:
```bash
./run.sh
```

The app will start on **http://localhost:8502**

---

## 🔑 Default Login Credentials

Use any of these accounts to test the system:

| Role | Login ID | Password | Permissions |
|------|----------|----------|-------------|
| **Admin** | `admin` | `admin123` | Create users, trainings, assign trainees, manage database |
| **Trainer** | `trainer` | `trainer123` | Add training materials, schedule training, mark attendance |
| **Surveyor** | `surveyor` | `surveyor123` | Complete training, take tests, receive certificates |
| **Plan Appraiser** | `appraiser` | `appraiser123` | Complete training, take tests, receive certificates |
| **Management** | `management` | `mgmt123` | View dashboards, monitor progress |

---

## 📋 What You Can Do in the App

### Admin Panel
- ✅ Create new users with auto-generated Login IDs
- ✅ Reset user passwords
- ✅ Manage user status (Active/Inactive)
- ✅ Create training programs
- ✅ Assign trainees to training
- ✅ Download the Excel database

### Trainer Panel
- ✅ Add training material links (slides, videos, references)
- ✅ Upload training content
- ✅ Generate multiple-choice questions (MCQs)
- ✅ Schedule training sessions
- ✅ Send notifications to trainees
- ✅ Mark attendance
- ✅ Save recording links

### Trainee Portal (Surveyor / Plan Appraiser)
- ✅ View assigned training
- ✅ Access training materials (slides, videos, recordings)
- ✅ Complete MCQ tests
- ✅ Receive certificates after passing
- ✅ Track progress

### Management Dashboard
- ✅ Monitor system metrics
- ✅ View training statistics
- ✅ Track completion rates
- ✅ Download reports

---

## 🗄️ Database

The app automatically creates and manages an Excel database:
- **Main Database**: `classification_society_training_platform.xlsx`
- Contains all user data, training info, test records, and certificates
- Auto-updated after every action
- Download from Admin Dashboard anytime

---

## ⚠️ Troubleshooting

### "App won't start"
```bash
# Kill any existing process on port 8502
lsof -ti:8502 | xargs kill -9

# Then run:
./run.sh
```

### "I can't see the app in Codespaces"
1. Look for **Ports** panel at the bottom
2. Find port `8502` in the list
3. Click the world icon to open in browser
4. If it's not there, add it manually:
   - Click "+ Add Port"
   - Enter `8502`
   - Press Enter

### "Login doesn't work"
- Make sure you're using the exact credentials above
- Check that you haven't locked the account (5 failed attempts locks it)
- Use Admin to reset any user's password

### "Database file not found"
The app creates it automatically on first run. If missing:
1. Go to Admin Panel
2. Look for "Reset Database" button in sidebar
3. Click it to initialize

---

## 📂 Project Structure

```
gdp-dashboard/
├── app.py                    # Main Streamlit app
├── run.sh                    # Quick start script
├── requirements.txt           # Python dependencies
├── .devcontainer/
│   └── devcontainer.json      # Codespaces configuration
├── classification_society_training_platform.xlsx  # Main database
└── README.md                  # Original project README
```

---

## 🔧 How to Restart the App

### In Codespaces:
```bash
# If app is running in a terminal tab
# Press Ctrl+C to stop it

# Then restart:
./run.sh
```

### Important Notes
- Port: `8502` (already configured for this repo)
- Server requires CORS disabled for Codespaces compatibility
- Database persists between sessions
- All user data is stored in Excel format

---

## 📞 Support

- **Issue**: Streamlit warnings → Safe to ignore (deprecation notices)
- **Issue**: Slow first load → Normal, app initializes Excel sheets
- **Issue**: Port already in use → Run `lsof -ti:8502 | xargs kill -9`

---

## ✨ Key Features Enabled

✅ Complete CRUD operations for all entities  
✅ Excel database persistence  
✅ Role-based access control  
✅ Automatic MCQ generation from content  
✅ Test scoring and certification  
✅ Email notifications (configurable)  
✅ Activity logging  
✅ Dashboard analytics  

---

**Ready to use!** Just log in with one of the default credentials above. 🚀
