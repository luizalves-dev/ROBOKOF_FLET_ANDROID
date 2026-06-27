@echo off
setlocal
cd /d "%~dp0"
python -m pip install --upgrade pip
if errorlevel 1 exit /b 1
python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1
flet build apk . --yes
if errorlevel 1 exit /b 1
echo.
echo APK gerado em: %CD%\build\apk
endlocal
