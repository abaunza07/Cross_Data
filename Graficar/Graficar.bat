@echo off
wsl.exe bash -c "cd /home/raczo07/Documents/Cross_Data && source .venv/bin/activate && python -m streamlit run Graficar/cruce_Report.py"
pause