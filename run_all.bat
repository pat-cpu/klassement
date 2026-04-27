@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM Ga naar de map waar dit .bat-bestand staat
cd /d "%~dp0"

REM Betrouwbaarder dan 'python'
set "PY_EXE=py -3"

REM Vraag jaartal (seizoenstart) en of er PDF's gegenereerd moeten worden
set "YEAR=%~1"
set "PDF=%~2"

if not defined YEAR set /p YEAR=Seizoen (startjaar, bv. 2025):
if not defined PDF set /p PDF=PDF's genereren? (Ja/Nee) [j/ja/n/nee]:

if "%YEAR%"=="" set "YEAR=2025"
if "%PDF%"=="" set "PDF=Ja"

REM Normaliseer PDF input (j/ja/y/yes -> Ja, n/nee -> Nee)
if /I "%PDF%"=="j" set "PDF=Ja"
if /I "%PDF%"=="ja" set "PDF=Ja"
if /I "%PDF%"=="y" set "PDF=Ja"
if /I "%PDF%"=="yes" set "PDF=Ja"
if /I "%PDF%"=="n" set "PDF=Nee"
if /I "%PDF%"=="nee" set "PDF=Nee"

echo === Gebruik Python: %PY_EXE%
echo Seizoen=%YEAR%  PDF=%PDF%
echo.

REM Check of de scripts bestaan
if not exist "scores_ophalen.py" (
  echo [FOUT] scores_ophalen.py ontbreekt.
  goto :END
)

if not exist "main.py" (
  echo [FOUT] main.py ontbreekt.
  goto :END
)

REM 1) CSV's ophalen
%PY_EXE% "scores_ophalen.py" --jaar %YEAR%
if errorlevel 1 (
  echo [FOUT] scores_ophalen.py mislukte.
  goto :END
)

REM Toon CSV's
echo ---- CSV's in data\%YEAR% ----
if exist "data\%YEAR%\*.csv" (
  dir /b "data\%YEAR%\*.csv"
) else (
  echo (geen CSV's gevonden)
)
echo.

REM Zorg dat outputmap bestaat
if not exist "output\%YEAR%" (
  mkdir "output\%YEAR%"
)

REM Maak timestamp aan voor log
for /f %%a in ('powershell -Command "Get-Date -Format yyyyMMdd_HHmm"') do set "TIMESTAMP=%%a"
set "LOGFILE=output\%YEAR%\run_%TIMESTAMP%.log"
echo Logbestand: %LOGFILE%
echo.

REM 2) HTML + PDF genereren: eerst LIVE tonen
echo === main.py (live output) ===
%PY_EXE% "main.py" --jaar %YEAR% --pdf %PDF%
set "RC=%ERRORLEVEL%"

REM En daarna dezelfde run nog eens naar log (stil)
%PY_EXE% "main.py" --jaar %YEAR% --pdf %PDF% 1>>"%LOGFILE%" 2>&1

if not "%RC%"=="0" (
  echo [FOUT] main.py mislukte. Zie log: %LOGFILE%
  goto :END
)

echo.
echo ✅ Klaar. HTML staat in: output\%YEAR%\html\
if /I "%PDF%"=="Ja" echo 📄 PDF's in: output\%YEAR%\pdf\



:END
pause
endlocal
