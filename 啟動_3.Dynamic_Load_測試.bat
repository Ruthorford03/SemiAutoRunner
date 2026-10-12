@echo off
title Dynamic Load Semi-Auto Test
cd /d "%~dp0"

echo ==============================================
echo       Running Dynamic Load Semi-Auto Test 
echo ==============================================
echo.

.venv\Scripts\python.exe general_auto_engine.py "3.Dynamic_Load_POL"

echo.
pause
