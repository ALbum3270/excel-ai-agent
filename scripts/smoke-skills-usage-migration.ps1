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
$dll = Join-Path $bin "ShareRibbon.dll"
if (-not (Test-Path $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-skills-usage-migration.ps1 -Build"
}

Push-Location $bin
try {
    Add-Type -Path $dll

    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()
    $dbPath = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()
    $connString = [ShareRibbon.OfficeAiDatabase]::GetConnectionString()
    $skillName = "smoke-skill-migrate-$([Guid]::NewGuid().ToString('N'))"
    $markerKey = "skills_usage_table_to_registry_v1"

    $conn = New-Object System.Data.SQLite.SQLiteConnection($connString)
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
DELETE FROM data_migration_marker WHERE migration_key = @marker;
DELETE FROM skills_registry WHERE skill_name = @skill;
DELETE FROM skills_usage WHERE skill_name = @skill;
INSERT INTO skills_usage (skill_name, usage_count, success_count, total_tokens, last_used_at)
VALUES (@skill, 7, 5, 101, '2026-06-07 23:00:00');
"@
        [void]$cmd.Parameters.AddWithValue("@marker", $markerKey)
        [void]$cmd.Parameters.AddWithValue("@skill", $skillName)
        [void]$cmd.ExecuteNonQuery()
    }
    finally {
        $conn.Dispose()
    }

    $flags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
    $field = [ShareRibbon.OfficeAiDatabase].GetField("_initialized", $flags)
    if ($null -eq $field) {
        throw "Could not access OfficeAiDatabase._initialized"
    }
    $field.SetValue($null, $false)

    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()
    $registry = [ShareRibbon.AgentMemoryRepository]::GetSkillRegistryByName($skillName)
    if ($null -eq $registry) {
        throw "Smoke failed: migrated skill was not found in skills_registry."
    }
    if ($registry.UsageCount -ne 7 -or $registry.SuccessCount -ne 5) {
        throw "Smoke failed: expected usage/success 7/5, got $($registry.UsageCount)/$($registry.SuccessCount)."
    }

    $field.SetValue($null, $false)
    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()
    $again = [ShareRibbon.AgentMemoryRepository]::GetSkillRegistryByName($skillName)
    if ($again.UsageCount -ne 7 -or $again.SuccessCount -ne 5) {
        throw "Smoke failed: migration was not idempotent, got $($again.UsageCount)/$($again.SuccessCount)."
    }

    [pscustomobject]@{
        Database = $dbPath
        SkillName = $skillName
        RegistryUsage = $again.UsageCount
        RegistrySuccess = $again.SuccessCount
        Idempotent = $true
    } | Format-List
}
finally {
    Pop-Location
}
