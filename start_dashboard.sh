#!/bin/bash
cd /home/effica/weekly_dashboard

# Check if Streamlit is already running so we don't start it twice
if pgrep -f "streamlit run app_streamlit.py" > /dev/null; then
    echo "Server is already running."
else
    # Start the server in the background
    nohup venv/bin/streamlit run app_streamlit.py --server.port 8501 > streamlit.log 2>&1 &
    # Give it 3 seconds to boot up completely
    sleep 3
fi

# Automatically open the default web browser when possible
if command -v xdg-open > /dev/null 2>&1; then
    xdg-open http://localhost:8501 > /dev/null 2>&1 || echo "Open http://localhost:8501 in your browser."
else
    echo "Open http://localhost:8501 in your browser."
fi


#./start_dashboard.sh
#or also use bash start_dashboard.sh or sh start_dashboard.sh
