@echo off
REM Prepare MSI inputs: Release code build + OfficeAgent.vdproj SourcePath audit.
REM Does NOT build the MSI itself (requires VS Installer Projects).
REM
REM   build-installer-prep.bat
REM   build-installer-prep.bat -SkipBuild

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build-installer-prep.ps1" %*
exit /b %ERRORLEVEL%
