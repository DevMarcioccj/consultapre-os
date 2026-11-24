#!/usr/bin/env bash
# Gera dist/monitor_bebidas (binário para o SO em que você rodar: Linux ou Mac)
set -e
python3 -m pip install --upgrade pip
python3 -m pip install pyinstaller playwright beautifulsoup4 pandas openpyxl requests
python3 -m PyInstaller --onefile --name monitor_bebidas --collect-all playwright --collect-all openpyxl monitor_bebidas.py
echo "Pronto: dist/monitor_bebidas"
