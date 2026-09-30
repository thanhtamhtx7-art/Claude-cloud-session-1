@echo off
rem Bam dup de gom file sang thu muc moi (chi doc va chep, khong sua gi thu muc goc).
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Gom_file_chuyen_may.ps1" %*
echo.
pause
