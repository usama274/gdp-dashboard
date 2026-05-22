#!/bin/bash

# Setup Verification Script
# Checks that all components are properly configured

echo ""
echo "======================================"
echo "🔍 Classification Society Training Platform"
echo "   Setup Verification"
echo "======================================"
echo ""

PASS="✅"
FAIL="❌"

# Check Python
echo "Checking Python..."
if command -v python3 &> /dev/null; then
    PY_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
    echo "$PASS Python 3 installed: $PY_VERSION"
else
    echo "$FAIL Python 3 not found"
    exit 1
fi

echo ""
echo "Checking Python packages..."

# Check Streamlit
python3 -c "import streamlit; print('$PASS Streamlit:', streamlit.__version__)" 2>/dev/null || echo "$FAIL Streamlit not installed"

# Check Pandas
python3 -c "import pandas; print('$PASS Pandas:', pandas.__version__)" 2>/dev/null || echo "$FAIL Pandas not installed"

# Check openpyxl
python3 -c "import openpyxl; print('$PASS openpyxl:', openpyxl.__version__)" 2>/dev/null || echo "$FAIL openpyxl not installed"

echo ""
echo "Checking project files..."

# Check main app file
if [ -f "streamlit_app.py" ]; then
    echo "$PASS streamlit_app.py found"
else
    echo "$FAIL streamlit_app.py not found"
fi

# Check requirements
if [ -f "requirements.txt" ]; then
    echo "$PASS requirements.txt found"
else
    echo "$FAIL requirements.txt not found"
fi

# Check run script
if [ -f "run.sh" ]; then
    if [ -x "run.sh" ]; then
        echo "$PASS run.sh found and executable"
    else
        echo "$FAIL run.sh found but NOT executable (run: chmod +x run.sh)"
    fi
else
    echo "$FAIL run.sh not found"
fi

# Check Streamlit config
if [ -f ".streamlit/config.toml" ]; then
    echo "$PASS .streamlit/config.toml found"
else
    echo "$FAIL .streamlit/config.toml not found"
fi

# Check devcontainer
if [ -f ".devcontainer/devcontainer.json" ]; then
    echo "$PASS .devcontainer/devcontainer.json found"
else
    echo "$FAIL .devcontainer/devcontainer.json not found"
fi

# Check documentation
echo ""
echo "Checking documentation..."
[ -f "QUICKSTART.md" ] && echo "$PASS QUICKSTART.md found" || echo "$FAIL QUICKSTART.md not found"
[ -f "TROUBLESHOOTING.md" ] && echo "$PASS TROUBLESHOOTING.md found" || echo "$FAIL TROUBLESHOOTING.md not found"

# Check port availability
echo ""
echo "Checking port 8502..."
if ! lsof -i :8502 &> /dev/null; then
    echo "$PASS Port 8502 is available"
else
    echo "$FAIL Port 8502 is already in use"
    echo "   To free it: lsof -ti:8502 | xargs kill -9"
fi

# Check database
echo ""
echo "Checking database..."
if [ -f "classification_society_training_platform.xlsx" ]; then
    DB_SIZE=$(ls -lh classification_society_training_platform.xlsx | awk '{print $5}')
    echo "$PASS Database exists ($DB_SIZE)"
else
    echo "⚠️  Database not found (will be created on first run)"
fi

# Test Python syntax
echo ""
echo "Testing Python syntax..."
if python3 -m py_compile app.py 2>/dev/null; then
    echo "$PASS app.py has no syntax errors"
else
    echo "$FAIL app.py has syntax errors"
fi

echo ""
echo "======================================"
echo "✨ Verification Complete!"
echo "======================================"
echo ""
echo "🚀 To start the app, run:"
echo "   ./run.sh"
echo ""
echo "📍 Then open: http://localhost:8502"
echo ""
