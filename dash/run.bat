@echo off
REM ============================================================
REM run.bat - Launches the Quality BRM dashboard
REM Binds to all network interfaces so BG network users can access
REM ============================================================

REM Change to the folder where this .bat file lives
cd /d "%~dp0"

REM Activate the (quality) venv
REM Adjust the path below if your venv is elsewhere
call C:\Users\%USERNAME%\quality\Scripts\activate.bat

REM Start Streamlit
REM   --server.address=0.0.0.0     -> allow LAN access (not just localhost)
REM   --server.port=8501            -> fixed port
REM   --server.headless=true        -> don't auto-open browser
REM   --browser.gatherUsageStats=false
streamlit run app.py ^
  --server.address=0.0.0.0 ^
  --server.port=8501 ^
  --server.headless=true ^
  --browser.gatherUsageStats=false

REM If Streamlit crashes, this window will show the error.
REM When running as a background task, Windows will restart it.
pause