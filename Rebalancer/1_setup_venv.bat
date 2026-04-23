@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1 && (set PY=py -3) || (set PY=python)
if not exist .venv\Scripts\python.exe (
  echo Creating .venv ...
  %PY% -m venv .venv
  if errorlevel 1 (
    echo Install Python 3.10+ from https://www.python.org/downloads/  ^(check "Add to PATH"^)
    exit /b 1
  )
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
if not exist mfd_pack.ini copy /Y mfd_pack.example.ini mfd_pack.ini
echo.
echo Next: add data\latestNAV_Reports.xlsx, put client files under data\clients\by_client\, edit mfd_pack.ini, then:
echo   python build_mfd_pack.py
echo.
echo Optional bootstrap ^(age/glide^):
echo   python build_mfd_pack.py --bootstrap-client-risk
echo.
endlocal
