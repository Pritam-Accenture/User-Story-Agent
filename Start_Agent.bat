@echo off
setlocal
cd /d "%~dp0"
set PORT=8000

rem ---- find Python
set PY=
where py >nul 2>&1 && set PY=py
if not defined PY ( where python >nul 2>&1 && python --version >nul 2>&1 && set PY=python )
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY (
  echo Python 3.10 or newer is required. Install it from https://www.python.org/downloads/ and run this again.
  pause & exit /b 1
)

if not exist "agent\workspace" mkdir "agent\workspace"

rem ---- first run: install Python packages
if not exist "agent\workspace\.deps_ok" (
  echo First run - installing Python packages, this takes a minute...
  %PY% -m pip install -r requirements.txt || ( echo Package install failed. & pause & exit /b 1 )
  echo ok> "agent\workspace\.deps_ok"
)

rem ---- first run: build the user interface
if not exist "agent\ui\dist\index.html" (
  where npm >nul 2>&1 || ( echo Node.js is required to build the UI. Install it from https://nodejs.org and run this again. & pause & exit /b 1 )
  echo First run - building the user interface...
  pushd agent\ui
  call npm install || ( echo npm install failed. & popd & pause & exit /b 1 )
  call npm run build || ( echo UI build failed. & popd & pause & exit /b 1 )
  popd
)

rem ---- already running?
netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul && (
  echo The agent is already running - opening it.
  start "" http://127.0.0.1:%PORT%
  exit /b 0
)

rem ---- start the server in a minimised window and open the browser
start "User Story Agent server" /min %PY% agent\server.py %PORT%
echo Starting... the browser opens in a few seconds.
timeout /t 5 /nobreak >nul
start "" http://127.0.0.1:%PORT%
echo.
echo The agent is running. Use Stop_Agent.bat to close it.
timeout /t 4 /nobreak >nul
