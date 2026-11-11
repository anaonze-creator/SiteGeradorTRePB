@echo off
cd /d "%~dp0"
start "Servidor Gerador de TR e PB" /min cmd /c "python -m http.server 8000"
timeout /t 2 /nobreak >nul
start "" "http://localhost:8000"
