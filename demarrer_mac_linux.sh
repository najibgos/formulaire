#!/usr/bin/env bash
# Lancement du formulaire d'inscription (macOS / Linux)
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt --quiet
python3 -m streamlit run app.py
