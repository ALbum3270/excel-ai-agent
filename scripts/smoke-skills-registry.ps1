param(
    [switch]$Build,
    [string]$Configuration = "Debug",
    [string]$Platform = "AnyCPU"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

if ($Build) {
    if (-not (Test-Path $msbuild)) {
        throw "MSBuild not found: $msbuild"
    }
    & $msbuild (Join-Path $repoRoot "WordAi\WordAi.vbproj") /p:Configuration=$Configuration /p:Platform=$Platform
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed, exit code: $LASTEXITCODE"
    }
}

$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
if (-not (Test-Path (Join-Path $bin "ShareRibbon.dll"))) {
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-skills-registry.ps1 -Build"
}

Push-Location $bin
try {
    Add-Type -Path (Join-Path $bin "ShareRibbon.dll")

    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()
    $dbPath = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()
    $skillName = "smoke-skill-registry"

    [ShareRibbon.SkillsService]::RecordSkillUsage($skillName, $true, 42)

    $registry = [ShareRibbon.AgentMemoryRepository]::GetSkillRegistryByName($skillName)
    $legacy = [ShareRibbon.MemoryRepository]::GetSkillUsage($skillName)

    $result = [pscustomobject]@{
        Database = $dbPath
        SkillName = $skillName
        RegistryUsage = $(if ($null -ne $registry) { $registry.UsageCount } else { 0 })
        RegistrySuccess = $(if ($null -ne $registry) { $registry.SuccessCount } else { 0 })
        LegacyUsage = $(if ($null -ne $legacy) { $legacy.UsageCount } else { 0 })
        LegacySuccess = $(if ($null -ne $legacy) { $legacy.SuccessCount } else { 0 })
    }

    $result | Format-List

    if ($null -eq $registry -or $registry.UsageCount -lt 1 -or $registry.SuccessCount -lt 1) {
        throw "Smoke failed: skills_registry was not updated."
    }

    if ($null -eq $legacy -or $legacy.UsageCount -lt 1 -or $legacy.SuccessCount -lt 1) {
        throw "Smoke failed: legacy skills_usage compatibility write was not updated."
    }
}
finally {
    Pop-Location
}
