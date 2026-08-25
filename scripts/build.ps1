param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$NoRestore,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))

function Get-MSBuildPath {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vswhere) {
        $candidate = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" |
            Select-Object -First 1
        if (-not [string]::IsNullOrWhiteSpace($candidate)) { return $candidate }
    }
    $command = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) { return $command.Source }
    throw "MSBuild was not found. Install Visual Studio Build Tools with .NET desktop and Office/VSTO components."
}

$msbuild = Get-MSBuildPath
$solution = Join-Path $repoRoot "ExcelAiAgent.sln"
$target = if ($Clean) { "Clean;Build" } else { "Build" }
$arguments = @($solution, "/t:$target", "/p:Configuration=$Configuration", "/p:RestorePackagesConfig=true", "/p:BuildCodeOnly=true", "/m", "/v:minimal")
if (-not $NoRestore) { $arguments += "/restore" }

& $msbuild @arguments
if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }
Write-Host "PASS: ExcelAiAgent.sln ($Configuration)"
