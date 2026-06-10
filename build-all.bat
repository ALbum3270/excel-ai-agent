@echo off
echo ================================
echo Building WordAi...
echo ================================
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "F:\ai\code\AiHelper\WordAi\WordAi.vbproj" /t:Build /p:Configuration=Debug /v:minimal /nologo
if %ERRORLEVEL% NEQ 0 (
    echo WordAi build FAILED
    exit /b 1
)
echo WordAi build SUCCESS
echo.

echo ================================
echo Building ExcelAi...
echo ================================
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "F:\ai\code\AiHelper\ExcelAi\ExcelAi.vbproj" /t:Build /p:Configuration=Debug /v:minimal /nologo
if %ERRORLEVEL% NEQ 0 (
    echo ExcelAi build FAILED
    exit /b 1
)
echo ExcelAi build SUCCESS
echo.

echo ================================
echo Building PowerPointAi...
echo ================================
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "F:\ai\code\AiHelper\PowerPointAi\PowerPointAi.vbproj" /t:Build /p:Configuration=Debug /v:minimal /nologo
if %ERRORLEVEL% NEQ 0 (
    echo PowerPointAi build FAILED
    exit /b 1
)
echo PowerPointAi build SUCCESS
echo.

echo ================================
echo All projects built successfully!
echo ================================
