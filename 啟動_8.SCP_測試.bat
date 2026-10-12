@echo off
title SCP Semi-Auto Test
cd /d "%~dp0"

echo =======================================================
echo       Running SCP Semi-Auto Test 
echo =======================================================
echo.

echo [*] Phase 1: Short-Circuit Protection After Power Up...
.venv\Scripts\python.exe general_auto_engine.py "8.Short-Circuit_protection_After"
echo.

echo [*] Phase 2: Short-Circuit Protection Before Power Up...
.venv\Scripts\python.exe general_auto_engine.py "8.Short-Circuit_protection_Before"
echo.

echo =======================================================
echo       SCP Test Completed!
echo =======================================================
echo.
pause
