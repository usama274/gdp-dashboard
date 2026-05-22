# 🎓 Classification Society Training Platform

A complete **Streamlit-based training management system** with Excel database integration, designed for managing maritime classification society training programs.

## ✨ Features

- 👥 **User Management** - Create users with auto-generated login IDs
- 📚 **Training Management** - Create and manage training programs
- 📝 **Content Management** - Upload slides, documents, and create MCQ tests
- 👨‍🏫 **Trainer Tools** - Schedule training, mark attendance, manage recordings
- 📊 **Dashboard** - Monitor progress, view analytics
- 🏆 **Certification** - Auto-generate certificates upon passing tests
- 📱 **Excel Integration** - All data stored in Excel format
- 🔐 **Role-Based Access** - Admin, Trainer, Surveyor, Plan Appraiser, Management roles

## 🚀 Quick Start

### Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

### Run the Application

```bash
# Using the quick start script (recommended)
./run.sh

# OR run directly with Streamlit
streamlit run app.py --server.port 8502 --server.address 0.0.0.0 --server.enableCORS false
```

The app will be available at: **http://localhost:8502**

## 🔐 Default Credentials

| Role | Username | Password |
|------|----------|----------|
| **Admin** | `admin` | `admin123` |
| **Trainer** | `trainer` | `trainer123` |
| **Surveyor** | `surveyor` | `surveyor123` |
| **Plan Appraiser** | `appraiser` | `appraiser123` |
| **Management** | `management` | `mgmt123` |

## 📖 Documentation

- **[QUICKSTART.md](QUICKSTART.md)** - Getting started guide with all features
- **[TROUBLESHOOTING.md](TROUBLESHOOTING.md)** - Solutions to common issues
- **[SETUP_SUMMARY.md](SETUP_SUMMARY.md)** - Complete setup overview

## 🔧 Codespaces Setup

This project is fully configured for GitHub Codespaces:

1. Click "Code" → "Codespaces" → "Create codespace"
2. Wait for environment to initialize
3. App starts automatically on port 8502
4. Open the port in browser (8502)

All dependencies and configurations are pre-set in `.devcontainer/devcontainer.json`

## 📋 Project Structure

```
gdp-dashboard/
├── app.py                           # Main application
├── run.sh                           # Quick start script
├── verify-setup.sh                  # Verification script
├── requirements.txt                 # Python dependencies
├── QUICKSTART.md                    # User guide
├── TROUBLESHOOTING.md              # Troubleshooting
├── SETUP_SUMMARY.md                # Setup overview
├── .streamlit/config.toml          # Streamlit config
├── .devcontainer/                  # Codespaces config
└── classification_society_training_platform.xlsx  # Database
```

## 🛠️ Technology Stack

- **Framework**: Streamlit 1.57+
- **Database**: Excel (openpyxl 3.1+)
- **Data Processing**: Pandas 3.0+
- **Python**: 3.11
- **Hosting**: Docker, GitHub Container Registry, Render, VPS

## 🚢 Docker and Deployment

This repo includes a `Dockerfile` and GitHub Actions workflow to build and push the container to GitHub Container Registry (`ghcr.io`).

To build locally:
```bash
docker build -t gdp-dashboard:latest .
```

To run locally:
```bash
docker run -d --restart unless-stopped -p 8502:8502 gdp-dashboard:latest
```

To deploy via GitHub Actions and Render:
1. Add `RENDER_API_KEY` and `RENDER_SERVICE_ID` to GitHub Secrets.
2. Push to `main`.
3. The workflow will build the Docker image, push to `ghcr.io/<owner>/gdp-dashboard:latest`, and update the Render service image.

## 💡 Usage Examples

### As Admin
```
1. Create users → Set roles → Auto-generate login IDs
2. Create training programs → Select trainer
3. Assign trainees → Set due dates
4. Monitor dashboard → Download Excel database
```

### As Trainer
```
1. Add material links (slides, videos)
2. Upload content → Auto-generate MCQs
3. Schedule training → Send notifications
4. Mark attendance → Save recordings
```

### As Trainee (Surveyor/Appraiser)
```
1. View assigned training
2. Access all materials
3. Complete MCQ tests
4. Receive certificate after passing
```

## ⚙️ Configuration

### Streamlit Config
Located in `.streamlit/config.toml`:
- Port: 8502
- Server address: 0.0.0.0 (Codespaces compatible)
- CORS: Disabled (required for Codespaces)

### Codespaces Config
Located in `.devcontainer/devcontainer.json`:
- Python 3.11 image
- Auto-installs dependencies
- Port 8502 forwarded
- VS Code extensions included

## 📚 Database

The application uses Excel as its database:
- **File**: `classification_society_training_platform.xlsx`
- **Auto-created** on first run
- **Persists** between sessions
- **Downloadable** from Admin Dashboard
- **Resettable** if needed

Database tables include:
- Users, Trainings, Training_Content, Question_Bank
- Training_Records, Notifications, Certificates, Activity_Log
- Dashboard, System, Role_Permissions

## 🐛 Troubleshooting

### Port 8502 already in use?
```bash
lsof -ti:8502 | xargs kill -9
./run.sh
```

### Dependencies missing?
```bash
python3 -m pip install -r requirements.txt --upgrade
```

### Verify setup?
```bash
./verify-setup.sh
```

See **[TROUBLESHOOTING.md](TROUBLESHOOTING.md)** for more solutions.

## 📊 Available Roles & Permissions

| Action | Admin | Trainer | Surveyor | Appraiser | Management |
|--------|-------|---------|----------|-----------|-----------|
| Create Users | ✅ | ❌ | ❌ | ❌ | ❌ |
| Create Training | ✅ | ❌ | ❌ | ❌ | ❌ |
| Add Materials | ❌ | ✅ | ❌ | ❌ | ❌ |
| Schedule Training | ❌ | ✅ | ❌ | ❌ | ❌ |
| View Training | ❌ | ✅ | ✅ | ✅ | ✅ |
| Take Tests | ❌ | ❌ | ✅ | ✅ | ❌ |
| View Dashboard | ✅ | ✅ | ✅ | ✅ | ✅ |
| Download Database | ✅ | ❌ | ❌ | ❌ | ❌ |

## 📝 License

See [LICENSE](LICENSE) file for details.

## 🎯 Next Steps

1. **Read** [QUICKSTART.md](QUICKSTART.md) for complete feature guide
2. **Run** `./run.sh` to start the app
3. **Login** with admin credentials
4. **Create** a training program and get started!

---

**Version**: 1.0 | **Status**: ✅ Production Ready | **Updated**: May 21, 2026
