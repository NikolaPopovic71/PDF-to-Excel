@echo off
REM Pravi PdfUExcel.exe (pokrenuti na Windows racunaru koji ima Python 3.10+)
python -m pip install -r requirements.txt
python -m PyInstaller --onefile --windowed --name PdfUExcel --collect-all tkinterdnd2 app.py
echo.
echo Gotovo! Fajl je u folderu dist\PdfUExcel.exe
pause
