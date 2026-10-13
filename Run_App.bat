@echo off
cd /d "%~dp0"
echo ======================================================
echo    Checking requirements and starting App...
echo ======================================================
python -m pip install --quiet streamlit pdfplumber pandas openpyxl pywin32
streamlit run app.py
pause