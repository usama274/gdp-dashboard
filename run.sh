#!/bin/bash

# Classification Society Training Platform - Quick Start Script
# Run this from the repo root to start the Streamlit app

python3 -m pip install --user -r requirements.txt > /dev/null 2>&1

echo ""
echo "🚀 Starting Classification Society Training Platform..."
PUBLIC_URL=${APP_PUBLIC_URL:-https://bug-free-space-doodle-64p7rvp6jrg245q7-8502.app.github.dev/}

echo "   Port: 8502"
echo "   Local URL: http://localhost:8502"
echo "   Public URL: ${PUBLIC_URL}"
echo ""
echo "📝 Default credentials:"
echo "   Admin:  admin / ${DEFAULT_ADMIN_PASSWORD:-Admin@1234}"
echo "   Trainer: trainer / ${DEFAULT_TRAINER_PASSWORD:-Trainer@1234}"
echo "   Surveyor: surveyor / ${DEFAULT_SURVEYOR_PASSWORD:-Surveyor@1234}"
echo ""

streamlit run app.py --server.port 8502 --server.address 0.0.0.0
