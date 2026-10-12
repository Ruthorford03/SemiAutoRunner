@echo off
title Switching Waveforms Semi-Auto Test
cd /d "%~dp0"

echo =======================================================
echo       Running Switching Waveforms Semi-Auto Test 
echo =======================================================
echo.

echo [*] Phase 1: High Side Rising Edge...
.venv\Scripts\python.exe general_auto_engine.py "5.Switching_HS_Rise"
echo.

echo [*] Phase 2: High Side Falling Edge...
.venv\Scripts\python.exe general_auto_engine.py "5.Switching_HS_Fall"
echo.

echo [*] Phase 3: Low Side Falling Edge...
.venv\Scripts\python.exe general_auto_engine.py "5.Switching_LS_Fall"
echo.

echo [*] Phase 4: Low Side Rising Edge...
.venv\Scripts\python.exe general_auto_engine.py "5.Switching_LS_Rise"
echo.

echo =======================================================
echo       Switching Waveforms Test Completed!
echo =======================================================
echo.
pause
