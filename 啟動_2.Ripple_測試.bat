@echo off
title Ripple Semi-Auto Test
cd /d "%~dp0"

echo ==============================================
echo       Running Ripple Semi-Auto Test 
echo ==============================================
echo.

.venv\Scripts\python.exe general_auto_engine.py "2.Ripple_and_Noise_POL"

echo.
pause
