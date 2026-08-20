@echo off
pip install -r requirements.txt --no-warn-script-location --disable-pip-version-check
python -m streamlit run app.py
pause