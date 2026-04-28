@echo off
cd /d "%~dp0"
echo Starting SKU Agent UI...
python -m streamlit run sku_agent_ui.py
pause
