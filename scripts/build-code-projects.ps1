param(
    [string]$Configuration = "Debug",
    [string]$Platform = "AnyCPU",
    [string[]]$Projects = @(
        "ShareRibbon\ShareRibbon.vbproj",
        "WordAi\WordAi.vbproj",
        "ExcelAi\ExcelAi.vbproj",
        "PowerPointAi\PowerPointAi.vbproj"
    )
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

function Get-MSBuildPath {
    $candidates = @(
        "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vswhere) {
        $installPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if (-not [string]::IsNullOrWhiteSpace($installPath)) {
            $msbuild = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path -LiteralPath $msbuild) {
                return $msbuild
            }
        }
    }

    $cmd = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw "MSBuild.exe not found. Install Visual Studio with MSBuild or add MSBuild.exe to PATH."
}

function Invoke-ProjectBuild {
    param(
        [string]$ProjectPath
    )

    $fullPath = Join-Path $repoRoot $ProjectPath
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Project not found: $ProjectPath"
    }

    Write-Host ""
    Write-Host "==> Build $ProjectPath [$Configuration|$Platform]"
    & $script:MSBuild $fullPath /t:Build /p:Configuration=$Configuration /p:Platform=$Platform /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed: $ProjectPath"
    }
    Write-Host "PASS: $ProjectPath"
}

$script:MSBuild = Get-MSBuildPath
Write-Host "MSBuild: $script:MSBuild"

Push-Location $repoRoot
try {
    foreach ($project in $Projects) {
        Invoke-ProjectBuild -ProjectPath $project
    }

    Write-Host ""
    Write-Host "Code projects built successfully."
}
finally {
    Pop-Location
}

