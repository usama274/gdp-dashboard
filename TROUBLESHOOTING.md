# 🔧 Troubleshooting Guide

## Common Issues and Solutions

### Issue 1: Port 8502 Already in Use
**Symptom**: "Address already in use" when starting the app

**Solution**:
```bash
# Find and kill the process using port 8502
lsof -ti:8502 | xargs kill -9

# Then restart
./run.sh
```

---

### Issue 2: Dependencies Missing
**Symptom**: "ModuleNotFoundError" when starting

**Solution**:
```bash
# Reinstall dependencies
python3 -m pip install --user -r requirements.txt --upgrade

# Try again
./run.sh
```

---

### Issue 3: Can't Access the App in Browser
**Symptom**: "Connection refused" or "Cannot reach localhost:8502"

**Solution for Codespaces**:
1. Check the **Ports** tab at the bottom of VS Code
2. Look for `8502` in the list
3. If missing:
   - Click **+ Add Port**
   - Type `8502`
   - Press Enter
4. Right-click the port and select "Open in Browser" (globe icon)

**Solution for Local Machine**:
1. Verify the app is running (look for Streamlit startup messages)
2. Open browser: `http://localhost:8502`
3. If still not working:
   ```bash
   # Check if port is listening
   netstat -tuln | grep 8502
   ```

---

### Issue 4: Login Failed or Account Locked
**Symptom**: "Invalid credentials" or account won't unlock

**Solution**:
1. Make sure you're using **exact** credentials:
   - Admin: `admin` / `admin123`
   - Trainer: `trainer` / `trainer123`
   - Surveyor: `surveyor` / `surveyor123`
   - Appraiser: `appraiser` / `appraiser123`

2. If account is locked (5 failed attempts):
   - Log in with **Admin** account
   - Go to Admin Panel → Password Reset / Status
   - Select the locked user
   - Click "Reset Password"
   - Use the new temporary password

3. Clear browser cache/cookies:
   - Press `Ctrl+Shift+Delete` (or `Cmd+Shift+Delete` on Mac)
   - Select "All time"
   - Clear cookies and cache
   - Refresh the page

---

### Issue 5: Database Not Found
**Symptom**: Database file missing or app shows empty data

**Solution**:
1. In Admin Panel, click "Reset Database" button (sidebar)
2. Or manually delete the database file and restart:
   ```bash
   rm classification_society_training_platform.xlsx
   ./run.sh
   ```

---

### Issue 6: Streamlit Warnings (use_container_width, Arrow serialization)
**Symptom**: Yellow/red warning messages in terminal

**Status**: ✅ **FIXED** - Updated to use `width="stretch"` instead

**If still seeing warnings**:
- These are deprecation notices and don't affect functionality
- Warnings have been suppressed in `.streamlit/config.toml`
- Safe to ignore

---

### Issue 7: App Runs But Shows Blank Page
**Symptom**: App loads but no login screen visible

**Solution**:
1. Hard refresh browser: `Ctrl+Shift+R` (or `Cmd+Shift+R` on Mac)
2. Clear browser storage:
   - Press F12 for Developer Tools
   - Go to Application → Storage
   - Click "Clear site data"
3. Restart the app:
   ```bash
   # Press Ctrl+C in the terminal running the app
   ./run.sh
   ```

---

### Issue 8: Slow Performance / Freezing
**Symptom**: App is very slow or freezes during operations

**Solution**:
1. **First time loading is slower** - normal, initializing database
2. Check available disk space:
   ```bash
   df -h
   ```
3. Restart the app and browser
4. Check system resources:
   ```bash
   top
   ```

---

### Issue 9: Excel File Keeps Resetting
**Symptom**: Data disappears or Excel file gets recreated

**Causes**:
- Clicking "Reset Database" button (intentional)
- Database corruption
- Insufficient disk space

**Solution**:
1. Check the Excel file is readable:
   ```bash
   ls -lh classification_society_training_platform.xlsx
   ```
2. Backup current database:
   ```bash
   cp classification_society_training_platform.xlsx backup.xlsx
   ```
3. Verify file permissions:
   ```bash
   chmod 666 classification_society_training_platform.xlsx
   ```

---

### Issue 10: CORS or XSRF Errors
**Symptom**: "CORS error" or "XSRF token missing"

**Status**: ✅ **ALREADY CONFIGURED** 

Configuration in `.streamlit/config.toml`:
```toml
enableCORS = false
enableXsrfProtection = false
```

This is necessary for Codespaces compatibility.

---

## Getting Help

If your issue isn't listed:

1. **Check logs**:
   ```bash
   # Restart app and watch full output
   ./run.sh
   ```

2. **Check configuration**:
   ```bash
   cat .streamlit/config.toml
   cat .devcontainer/devcontainer.json
   ```

3. **Verify environment**:
   ```bash
   python3 --version
   pip3 show streamlit pandas openpyxl
   ```

4. **Check database integrity**:
   ```bash
   python3 -c "import pandas as pd; df=pd.read_excel('classification_society_training_platform.xlsx'); print(df.head())"
   ```

---

## Quick Commands Reference

```bash
# Start the app
./run.sh

# Kill app process
lsof -ti:8502 | xargs kill -9

# Check if port is listening
lsof -i :8502

# Install/update dependencies
python3 -m pip install --user -r requirements.txt --upgrade

# Reset everything
rm classification_society_training_platform.xlsx
./run.sh

# Check Python version
python3 --version

# List installed packages
pip3 list | grep -E 'streamlit|pandas|openpyxl'
```

---

## Still Stuck?

1. Check the **QUICKSTART.md** for basic setup
2. Verify **Codespaces Settings**:
   - Python 3.11 enabled
   - Port 8502 forwarded
   - Extensions installed
3. Try recreating the Codespaces environment
4. Contact support with:
   - Full error message/output
   - Steps to reproduce
   - Current credentials being used
