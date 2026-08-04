#!/bin/bash

GITBASH="/c/Program Files/Git/git-bash.exe"

cmd.exe /c start "" "$GITBASH" -c "cd /c/Users/alqud/Desktop/2026/Onequiz-addin && npm start; exec bash"

cmd.exe /c start "" "$GITBASH" -c "cd /c/Users/alqud/Desktop/2026/Onequiz-addin/server && npx nodemon server.js; exec bash"

cmd.exe /c start "" "$GITBASH" -c "ngrok http 3000; exec bash"