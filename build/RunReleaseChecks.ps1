param(
    [switch]$SkipBuild,
    [switch]$SkipSmoke,
    [switch]$VerifySignatures
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Invoke-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Host ""
    Write-Host "==> $Name"
    & $Action
    Write-Host "PASS: $Name"
}

function Get-MSBuildPath {
    $knownPath = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe'
    if (Test-Path -LiteralPath $knownPath) {
        return $knownPath
    }

    $cmd = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw 'MSBuild.exe not found. Install Visual Studio Build Tools or add MSBuild.exe to PATH.'
}

Set-Location $repoRoot

Invoke-Step 'Version audit' {
    & (Join-Path $repoRoot 'build\AuditVersion.ps1')
}

Invoke-Step 'Installer Release audit' {
    & (Join-Path $repoRoot 'build\AuditInstallerRelease.ps1')
}

if (-not $SkipBuild) {
    $msbuild = Get-MSBuildPath
    $projects = @(
        'WordAi\WordAi.vbproj',
        'ExcelAi\ExcelAi.vbproj',
        'PowerPointAi\PowerPointAi.vbproj'
    )

    foreach ($project in $projects) {
        Invoke-Step "Build Release $project" {
            & $msbuild (Join-Path $repoRoot $project) /p:Configuration=Release /p:Platform=AnyCPU
            if ($LASTEXITCODE -ne 0) {
                throw "MSBuild failed for $project"
            }
        }
    }
}

if (-not $SkipSmoke) {
    $smokeScripts = @(
        'scripts\smoke-memory-pipeline.ps1',
        'scripts\smoke-skills-registry.ps1',
        'scripts\smoke-db-schema-drift.ps1'
    )

    foreach ($script in $smokeScripts) {
        Invoke-Step "Smoke $script" {
            & (Join-Path $repoRoot $script)
        }
    }
}

Invoke-Step 'Git diff whitespace check' {
    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($null -eq $git) {
        $git = Get-Command git -ErrorAction SilentlyContinue
    }
    if ($null -eq $git) {
        throw 'git not found.'
    }

    & $git.Source -C $repoRoot diff --check
    if ($LASTEXITCODE -ne 0) {
        throw 'git diff --check failed.'
    }
}

if ($VerifySignatures) {
    Invoke-Step 'Signature verification' {
        & (Join-Path $repoRoot 'build\VerifySignatures.ps1')
    }
}

Write-Host ""
Write-Host 'Release checks complete.'
