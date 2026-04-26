@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ==========================================
REM INSTELLINGEN
REM ==========================================
set "APP_DIR=C:\Users\patri\Documents\WHISKIES\klassement\Inschrijvingen"
set "APP_FILE=App.py"
set "PYTHON_EXE=%APP_DIR%\..\.venv\Scripts\python.exe"

REM ==========================================
REM NAAR JUISTE MAP
REM ==========================================
cd /d "%APP_DIR%" || (
    echo FOUT: map niet gevonden:
    echo %APP_DIR%
    pause
    exit /b 1
)

echo.
echo ==========================================
echo THE WHISKIES SCANNER START
echo ==========================================
echo Map: %CD%
echo.

REM ==========================================
REM CONTROLE PYTHON IN .VENV
REM ==========================================
if not exist "%PYTHON_EXE%" (
    echo FOUT: python.exe niet gevonden in:
    echo %PYTHON_EXE%
    echo.
    echo Controleer of je .venv bestaat op:
    echo %APP_DIR%\..\.venv
    pause
    exit /b 1
)

echo Python:
"%PYTHON_EXE%" --version
echo.

REM ==========================================
REM BESTAAT APP.PY?
REM ==========================================
if not exist "%APP_FILE%" (
    echo FOUT: %APP_FILE% niet gevonden in:
    echo %CD%
    pause
    exit /b 1
)

REM ==========================================
REM OUDE SERVER OP POORT 5000 STOPPEN
REM ==========================================
echo Oude scanner op poort 5000 stoppen indien nodig...
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
    echo   Proces PID %%P wordt afgesloten...
    taskkill /PID %%P /F >nul 2>&1
)

timeout /t 1 /nobreak >nul

REM ==========================================
REM BROWSER OPENEN
REM ==========================================
echo Browser openen...
start "" "http://127.0.0.1:5000"

REM ==========================================
REM SCANNER STARTEN
REM ==========================================
echo Scanner starten...
echo.
"%PYTHON_EXE%" "%APP_FILE%"

echo.
echo Scanner is gestopt.
pause
exit /b