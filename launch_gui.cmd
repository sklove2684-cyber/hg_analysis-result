@echo off & cd /d "%~dp0" & call ".venv\Scripts\honyu-automation.exe" & if errorlevel 1 (echo. & echo [ERROR] honyu-automation failed. & pause)
