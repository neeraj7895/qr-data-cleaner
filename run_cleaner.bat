@echo off
echo ==============================
echo   Starting QR Data Cleaner
echo ==============================
cd /d %~dp0
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)
call venv\Scripts\activate
echo Installing requirements...
pip install -r requirements.txt
echo Launching Streamlit App...
streamlit run app.py
pause
