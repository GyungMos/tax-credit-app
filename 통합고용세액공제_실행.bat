@echo off
echo 통합고용세액공제 자동화 서비스를 시작합니다...
cd /d "%~dp0"
start /min cmd /c "node server/index.js"
cd client
start /min cmd /c "npm run dev -- --port 3000"
timeout /t 5
start http://localhost:3000
echo 서비스가 시작되었습니다. 브라우저를 확인해주세요.
pause
