@echo off
cd /d "%~dp0"

echo === Git upload bezig... ===

git add .
git commit -m "update %date% %time%"
git push origin main

echo.
echo ✅ Klaar!
pause