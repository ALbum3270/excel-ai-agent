@echo off
REM Alias: code projects only (Debug). Does NOT build OfficeAgent.vdproj / MSI.
REM Prefer build-code.bat for clarity. See docs\build-and-installer.md
call "%~dp0build-code.bat" Debug
exit /b %ERRORLEVEL%
