@echo off
REM ============================================================
REM Construit TrafikSaveurs.exe (application bureau hors ligne)
REM ============================================================
cd /d "%~dp0"
echo [1/3] Installation des dependances...
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl pillow
echo [2/3] Generation de l'icone...
python make_icon.py
echo [3/3] Creation de l'executable...
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "TrafikSaveurs" ^
  --icon "icon.ico" ^
  --add-data "icon.ico;." ^
  --add-data "icon.png;." ^
  --add-data "index.html;." ^
  app.py
echo.
echo Termine ! Executable : dist\TrafikSaveurs.exe
pause
