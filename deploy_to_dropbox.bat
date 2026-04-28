@echo off
setlocal EnableExtensions
chcp 65001 >nul

set "SOURCE=C:\Users\patri\Documents\WHISKIES\klassement"
set "DEST=C:\Users\patri\Dropbox\THE WHISKIES\klassement"
set "BACKUP=C:\Users\patri\Dropbox\THE WHISKIES\BACKUP_klassement"

echo === VEILIGE DEPLOY NAAR DROPBOX ===
echo Bron: %SOURCE%
echo Doel: %DEST%
echo Backup: %BACKUP%
echo.

REM Check bron
if not exist "%SOURCE%\main.py" (
  echo [FOUT] Bronmap klopt niet. main.py niet gevonden.
  goto :END
)

REM Backup maken van huidige Dropbox-versie
if exist "%DEST%" (
  echo [INFO] Backup maken van huidige Dropbox-versie...
  robocopy "%DEST%" "%BACKUP%" /MIR /XD .venv .git __pycache__ /XF *.pyc
)

echo.
echo [INFO] Bestanden kopieren naar Dropbox...

REM Veilige kopie: GEEN .venv, GEEN .git, GEEN cache
robocopy "%SOURCE%" "%DEST%" /E /XD .venv .git __pycache__ /XF *.pyc

echo.
echo ✅ Deploy klaar.
echo Dropbox-versie bijgewerkt.
echo Backup staat in:
echo %BACKUP%
echo.

:END
pause
endlocal