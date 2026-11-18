@echo off
setlocal
cd /d "%~dp0"
title DiskGuard
echo ==========================================
echo  DiskGuard - DiskGenius report analyzer
echo ==========================================
echo.
where powershell.exe >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Windows PowerShell was not found.
  echo This environment may not include PowerShell.
  pause
  exit /b 9009
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0DiskGuard.ps1"
set "DG_EXIT=%ERRORLEVEL%"
echo.
if "%DG_EXIT%"=="0" (
  echo [DiskGuard] Finished successfully. Exit code: 0
) else (
  echo [DiskGuard] Finished with an error. Exit code: %DG_EXIT%
  echo Check the Output folder for DiskGuard_failure_*.log.
)
echo.
pause
exit /b %DG_EXIT%
