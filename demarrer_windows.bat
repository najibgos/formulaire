@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Installation des dependances (premiere fois seulement)...
python -m pip install -r requirements.txt --quiet
echo Lancement de l'application...
python -m streamlit run app.py
pause
