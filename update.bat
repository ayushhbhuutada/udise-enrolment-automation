@echo off & title UDISE Automation - Update & cd /d "%~dp0" & powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\update.ps1" & echo. & pause & exit /b
