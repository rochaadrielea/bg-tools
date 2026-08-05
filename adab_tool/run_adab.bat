@echo off
REM Launch the ADAB Compare desktop app.
cd /d "%~dp0"
REM Activate your (quality) venv — adjust the path if yours is elsewhere.
call C:\Users\%USERNAME%\quality\Scripts\activate.bat
python adab_gui.py
pause
