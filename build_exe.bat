@echo off
setlocal

python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --name glpi_app ^
  --hidden-import xlrd ^
  --hidden-import openpyxl ^
  --hidden-import odf ^
  main.py

echo.
echo Build finalizado. Executavel em dist\glpi_app.exe
pause
