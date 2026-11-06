@echo off
cd /d "%~dp0"
set "PY=C:\Python\Python312\python.exe"
set "PYW=C:\Python\Python312\pythonw.exe"
if not exist "%PY%" set "PY=python"
if not exist "%PYW%" set "PYW=%PY%"
"%PY%" -c "import openpyxl, serial, PIL" 2>nul
if errorlevel 1 (
  echo openpyxl kuruluyor...
  "%PY%" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo Kurulum basarisiz.
    pause
    exit /b 1
  )
)
start "" "%PYW%" "%~dp0telefon_excel.py"
