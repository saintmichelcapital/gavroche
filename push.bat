@echo off
cd /d "%~dp0"
echo === Gavroche : push vers GitHub ===
echo.
git status --short
echo.
echo --- push en cours ---
git push
echo.
echo --- terminé ---
pause
