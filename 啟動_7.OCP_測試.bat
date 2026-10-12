@echo off
title OCP Semi-Auto Test
cd /d "%~dp0"

echo =======================================================
echo       Running OCP Semi-Auto Test 
echo =======================================================
echo.

echo [*] Starting OCP Measurement...
.venv\Scripts\python.exe general_auto_engine.py "7.Overcurrent_protection_POL"
echo.

echo =======================================================
echo       OCP Test Completed!
echo =======================================================
echo.
pause
