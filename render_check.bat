@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul && (py render_check.py %*) || (python render_check.py %*)
