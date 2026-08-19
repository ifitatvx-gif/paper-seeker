@echo off
cd /d "%~dp0"
if not exist node_modules (
  echo [1/2] First run: installing dependencies, please wait...
  npm install --no-audit --no-fund --cache .\.npm-cache
)
echo [2/2] Starting server, opening http://localhost:3000 ...
start "" "http://localhost:3000"
node server.js
pause
