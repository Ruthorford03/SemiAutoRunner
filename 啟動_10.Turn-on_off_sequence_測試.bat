@echo off
title Turn-on ^& off sequence Semi-Auto Test
cd /d "%~dp0"

echo =======================================================
echo       Running Turn-on ^& off sequence Semi-Auto Test 
echo =======================================================
echo.

echo [*] Phase 1: Turn-ON Sequence...
.venv\Scripts\python.exe general_auto_engine.py "10.Turn-on & off sequence_Turn_ON"
echo.

echo [*] Phase 2: Turn-OFF Sequence...
.venv\Scripts\python.exe general_auto_engine.py "10.Turn-on & off sequence_Turn_OFF"
echo.

echo =======================================================
echo       Turn-on ^& off sequence Test Completed!
echo =======================================================
echo.
pause
