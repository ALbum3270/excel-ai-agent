@echo off
REM Build ShareRibbon + WordAi + ExcelAi + PowerPointAi only (no installer / vdproj).
REM Default: Debug. Pass Release as first arg if needed:
REM   build-code.bat
REM   build-code.bat Release

set "CONFIG=Debug"
if /I "%~1"=="Release" set "CONFIG=Release"
if /I "%~1"=="Debug" set "CONFIG=Debug"

echo Building code projects [%CONFIG%] ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build-code-projects.ps1" -Configuration %CONFIG%
exit /b %ERRORLEVEL%
