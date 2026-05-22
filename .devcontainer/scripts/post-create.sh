#!/bin/bash

# This script runs after the Codespaces container is created
# It ensures all dependencies are ready and shows setup info

set -e

echo ""
echo "=================================="
echo "🏗️  Setting up Classification Society Training Platform"
echo "=================================="
echo ""

# Install Python dependencies
echo "📦 Installing dependencies..."
python3 -m pip install --user -q -r requirements.txt
echo "✅ Dependencies installed"

# Verify database files exist
echo ""
echo "🗄️  Checking database..."
if [ ! -f "classification_society_training_platform.xlsx" ]; then
  echo "⚠️  Database will be created on first app run"
else
  echo "✅ Database exists"
fi

echo ""
echo "=================================="
echo "✨ SETUP COMPLETE!"
echo "=================================="
echo ""
echo "🚀 To start the app:"
echo "   ./run.sh"
echo ""
echo "   OR use the Codespaces terminal and run:"
echo "   streamlit run app.py --server.port 8502 --server.address 0.0.0.0 --server.enableCORS false"
echo ""
echo "📍 App will be available at: http://localhost:8502"
echo ""
echo "🔐 Default Login: admin / admin123"
echo ""
echo "📚 For more info, see: QUICKSTART.md"
echo ""
