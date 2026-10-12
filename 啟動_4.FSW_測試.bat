@echo off
title FSW Semi-Auto Test
cd /d "%~dp0"

echo ==============================================
echo       Running FSW Semi-Auto Test 
echo ==============================================
echo.

.venv\Scripts\python.exe general_auto_engine.py "4.FSW_POL"

echo.
pause
