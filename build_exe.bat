@echo off
setlocal

python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --onefile --windowed --name AtualizadorDeChamados main.py

echo.
echo Build finalizado. Executavel em dist\AtualizadorDeChamados.exe
pause
