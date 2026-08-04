@echo off

set GITBASH="C:\Program Files\Git\git-bash.exe"


start "" %GITBASH% -c "ngrok http https://localhost:3000; exec bash"

start "" %GITBASH% -c "cd /c/Users/alqud/Desktop/2026/Onequiz-addin/Onequiz && npm run dev-server; exec bash"

start "" %GITBASH% -c "cd /c/Users/alqud/Desktop/2026/Onequiz-addin/server && npx nodemon server.js; exec bash

