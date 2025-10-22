@echo off
setlocal

REM Start both backend and frontend dev servers for the Excel AI Add-in.
cd /d "%~dp0"

if not exist node_modules ( 
  echo [INFO] node_modules not found. Installing dependencies...
  call npm install || goto :error
)

if not defined OPENAI_API_KEY (
  echo [WARN] OPENAI_API_KEY is not set. Chat features will return an error until it is configured.
)

REM Ensure CRA dev server accepts requests from localhost.
set "WDS_ALLOWED_HOSTS=all"
set "DANGEROUSLY_DISABLE_HOST_CHECK=true"

echo [INFO] Starting development servers... (Ctrl+C to stop)
call npm run dev
goto :eof

:error
echo [ERROR] Failed to start the development environment.
exit /b 1
