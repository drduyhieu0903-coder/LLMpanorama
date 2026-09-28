@echo off
title Dental AI Evaluator - Web Streamlit Launcher
echo ========================================================
echo   Dental AI Evaluation Tool - Web Streamlit v2.0
echo ========================================================
echo.
echo Dang khoi dong may chu web tren http://localhost:8501...
echo.

cd /d "%~dp0"
python -m streamlit run app.py --server.port 8501 --server.headless false

if errorlevel 1 (
    echo.
    echo [LOI] Khong the khoi dong Streamlit. Vui long kiem tra lai moi truong Python.
    pause
)
