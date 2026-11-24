@echo off
REM Gera dist\monitor_bebidas.exe (rode no Windows, com Python 3.10+ instalado)
python -m pip install --upgrade pip
python -m pip install pyinstaller playwright beautifulsoup4 pandas openpyxl requests
python -m PyInstaller --onefile --name monitor_bebidas --collect-all playwright --collect-all openpyxl monitor_bebidas.py
echo.
echo Pronto: dist\monitor_bebidas.exe
pause
