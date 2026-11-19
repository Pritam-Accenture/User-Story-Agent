@echo off
set PORT=8000
set FOUND=
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  taskkill /PID %%p /F >nul 2>&1
  set FOUND=1
)
if defined FOUND ( echo User Story Agent stopped. ) else ( echo The agent was not running. )
timeout /t 3 /nobreak >nul
