@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build-code-projects.ps1" -Configuration Debug -Projects "ExcelAi\ExcelAi.vbproj","PowerPointAi\PowerPointAi.vbproj"
exit /b %ERRORLEVEL%
