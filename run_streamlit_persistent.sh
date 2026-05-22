#!/usr/bin/env bash
# Run Streamlit app persistently on port 8502 using nohup.
# Usage: ./run_streamlit_persistent.sh &
APP=app.py
PORT=8502
LOG=streamlit-persistent.log
nohup streamlit run "$APP" --server.port $PORT --server.enableCORS false > "$LOG" 2>&1 &
echo "Streamlit started on port $PORT (logs: $LOG)"
