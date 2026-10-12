@echo off
title Jitter Semi-Auto Test
cd /d "%~dp0"

echo =======================================================
echo       Running Jitter Semi-Auto Test 
echo =======================================================
echo.

echo [*] Starting Jitter Measurement...
.venv\Scripts\python.exe general_auto_engine.py "6.Jitter_POL"
echo.

echo =======================================================
echo       Jitter Test Completed!
echo =======================================================
echo.
pause
