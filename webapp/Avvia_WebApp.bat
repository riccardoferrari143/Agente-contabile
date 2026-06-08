@echo off
cd /d "%~dp0"
echo Avvio Agente Contabile Web App...
py -m streamlit run app.py
pause
