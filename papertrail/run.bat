@echo off
REM PaperTrail launcher (Windows)
cd /d "%~dp0"
python -c "import openpyxl" 2>nul || python -m pip install -r requirements.txt
python -m papertrail %*
pause
