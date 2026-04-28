echo === NIEUWE RUN_ALL 28-04 ===
pause
@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

REM Ga naar de map waar dit .bat-bestand staat
cd /d "%~dp0"

echo === THE WHISKIES - KLASSEMENT ===
echo Map: %CD%
echo.

REM -------------------------------------------------
REM Python kiezen:
REM - In Documenten: lokale .venv gebruiken
REM - In Dropbox: aparte veilige venv buiten Dropbox gebruiken
REM -------------------------------------------------

echo %CD% | find /I "Dropbox" >nul

if %errorlevel%==0 (
    echo [INFO] Running in Dropbox
    set "PY_EXE=C:\Users\patri\Documents\WHISKIES\venvs\klassement_dropbox\Scripts\python.exe"
) else (
    echo [INFO] Running in Documenten
    set "PY_EXE=%CD%\.venv\Scripts\python.exe"
)

echo Python: %PY_EXE%
echo.

REM -------------------------------------------------
REM Als Python/venv ontbreekt: alleen automatisch maken in Documenten
REM -------------------------------------------------

if not exist "%PY_EXE%" (

    echo [WAARSCHUWING] Python omgeving niet gevonden.

    echo %CD% | find /I "Dropbox" >nul
    if %errorlevel%==0 (
        echo [FOUT] Dropbox-venv ontbreekt.
        echo Maak eerst deze venv buiten Dropbox:
        echo C:\Users\patri\Documents\WHISKIES\venvs\klassement_dropbox
        goto :END
    )

    echo [INFO] Geen lokale venv gevonden. Nieuwe venv wordt aangemaakt...
    py -3 -m venv .venv

    if errorlevel 1 (
        echo [FOUT] Venv aanmaken mislukt.
        goto :END
    )

    set "PY_EXE=%CD%\.venv\Scripts\python.exe"
)

REM -------------------------------------------------
REM Packages installeren/controleren
REM -------------------------------------------------

REM -------------------------------------------------
REM Packages installeren/controleren
REM -------------------------------------------------

set "INSTALL_MARKER=%CD%\.venv\installed.ok"

echo Install-check bestand: %INSTALL_MARKER%

if exist "%INSTALL_MARKER%" (
    echo [INFO] Packages al geïnstalleerd.
) else (
    echo [INFO] Eerste keer: packages installeren...

    if not exist "requirements.txt" (
        echo [FOUT] requirements.txt niet gevonden.
        goto :END
    )

    "%PY_EXE%" -m pip install -r requirements.txt

    if errorlevel 1 (
        echo [FOUT] Installatie van requirements mislukt.
        goto :END
    )

    echo OK>"%INSTALL_MARKER%"
)

REM -------------------------------------------------
REM Vraag seizoen en PDF-keuze
REM -------------------------------------------------

set "YEAR=%~1"
set "PDF=%~2"

if not defined YEAR set /p YEAR=Seizoen startjaar bv. 2025: 
if not defined PDF set /p PDF=PDF's genereren? Ja/Nee [j/ja/n/nee]: 

if "%YEAR%"=="" set "YEAR=2025"
if "%PDF%"=="" set "PDF=Ja"

if /I "%PDF%"=="j" set "PDF=Ja"
if /I "%PDF%"=="ja" set "PDF=Ja"
if /I "%PDF%"=="y" set "PDF=Ja"
if /I "%PDF%"=="yes" set "PDF=Ja"
if /I "%PDF%"=="n" set "PDF=Nee"
if /I "%PDF%"=="nee" set "PDF=Nee"

echo.
echo Seizoen: %YEAR%
echo PDF: %PDF%
echo.

REM -------------------------------------------------
REM Controle scripts
REM -------------------------------------------------

if not exist "scores_ophalen.py" (
    echo [FOUT] scores_ophalen.py ontbreekt.
    goto :END
)

if not exist "main.py" (
    echo [FOUT] main.py ontbreekt.
    goto :END
)

REM -------------------------------------------------
REM 1. CSV's ophalen
REM -------------------------------------------------

echo === CSV's ophalen ===
"%PY_EXE%" "scores_ophalen.py" --jaar %YEAR%

if errorlevel 1 (
    echo [FOUT] scores_ophalen.py mislukte.
    goto :END
)

echo.
echo ---- CSV's in data\%YEAR% ----

if exist "data\%YEAR%\*.csv" (
    dir /b "data\%YEAR%\*.csv"
) else (
    echo Geen CSV's gevonden.
)

echo.

REM -------------------------------------------------
REM Outputmap en logbestand
REM -------------------------------------------------

if not exist "output\%YEAR%" (
    mkdir "output\%YEAR%"
)

for /f %%a in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmm"') do set "TIMESTAMP=%%a"

set "LOGFILE=output\%YEAR%\run_%TIMESTAMP%.log"

echo Logbestand: %LOGFILE%
echo.

REM -------------------------------------------------
REM 2. HTML + PDF genereren
REM -------------------------------------------------

echo === HTML/PDF genereren ===
"%PY_EXE%" "main.py" --jaar %YEAR% --pdf %PDF%
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo [FOUT] main.py mislukte. Zie log: %LOGFILE%
    goto :END
)

REM Log achteraf ook bewaren
"%PY_EXE%" "main.py" --jaar %YEAR% --pdf %PDF% 1>>"%LOGFILE%" 2>&1

echo.
echo ✅ Klaar. HTML staat in: output\%YEAR%\html\

if /I "%PDF%"=="Ja" (
    echo 📄 PDF's staan in: output\%YEAR%\pdf\

REM Laatste maand-PDF openen volgens seizoenvolgorde
set "LASTPDF="

for %%m in (September Oktober November December Januari Februari Maart April Mei Juni Juli Augustus) do (
    if exist "output\%YEAR%\pdf\%%m.pdf" (
        set "LASTPDF=output\%YEAR%\pdf\%%m.pdf"
    )
)

if defined LASTPDF (
    echo PDF openen: !LASTPDF!
    start "" "!LASTPDF!"
) else (
    echo [WAARSCHUWING] Geen maand-PDF gevonden om te openen.
)
)

:END
echo.
pause
endlocal
