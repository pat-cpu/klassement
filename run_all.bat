



@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"

REM Gebruik altijd de venv in deze map
set "PY_EXE=%CD%\.venv\Scripts\python.exe"

if not exist "%PY_EXE%" (
  echo [FOUT] Python venv niet gevonden:
  echo %PY_EXE%
  echo.
  echo Maak eerst de venv aan of plaats dit .bat-bestand in de juiste projectmap.
  goto :END
)

set "YEAR=%~1"
set "PDF=%~2"

if not defined YEAR set /p YEAR=Seizoen (startjaar, bv. 2025): 
if not defined PDF set /p PDF=PDF's genereren? (Ja/Nee) [j/ja/n/nee]: 

if "%YEAR%"=="" set "YEAR=2025"
if "%PDF%"=="" set "PDF=Ja"

if /I "%PDF%"=="j" set "PDF=Ja"
if /I "%PDF%"=="ja" set "PDF=Ja"
if /I "%PDF%"=="y" set "PDF=Ja"
if /I "%PDF%"=="yes" set "PDF=Ja"
if /I "%PDF%"=="n" set "PDF=Nee"
if /I "%PDF%"=="nee" set "PDF=Nee"

echo === Gebruik Python: %PY_EXE%
echo Seizoen=%YEAR%  PDF=%PDF%
echo.

if not exist "scores_ophalen.py" (
  echo [FOUT] scores_ophalen.py ontbreekt.
  goto :END
)

if not exist "main.py" (
  echo [FOUT] main.py ontbreekt.
  goto :END
)

REM 1) CSV's ophalen
"%PY_EXE%" "scores_ophalen.py" --jaar %YEAR%
if errorlevel 1 (
  echo [FOUT] scores_ophalen.py mislukte.
  goto :END
)

echo ---- CSV's in data\%YEAR% ----
if exist "data\%YEAR%\*.csv" (
  dir /b "data\%YEAR%\*.csv"
) else (
  echo (geen CSV's gevonden)
)
echo.

if not exist "output\%YEAR%" (
  mkdir "output\%YEAR%"
)

for /f %%a in ('powershell -Command "Get-Date -Format yyyyMMdd_HHmm"') do set "TIMESTAMP=%%a"
set "LOGFILE=output\%YEAR%\run_%TIMESTAMP%.log"

echo Logbestand: %LOGFILE%
echo.

echo === main.py ===
"%PY_EXE%" "main.py" --jaar %YEAR% --pdf %PDF%
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
  echo [FOUT] main.py mislukte. Zie log: %LOGFILE%
  goto :END
)

echo.
echo ✅ Klaar. HTML staat in: output\%YEAR%\html\
if /I "%PDF%"=="Ja" echo 📄 PDF's in: output\%YEAR%\pdf\

if /I "%PDF%"=="Ja" (
  for %%f in ("output\%YEAR%\pdf\*.pdf") do set "LASTPDF=%%f"
  start "" "!LASTPDF!"
)

:END
pause
endlocal