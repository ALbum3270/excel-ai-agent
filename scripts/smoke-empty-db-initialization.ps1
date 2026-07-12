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
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-empty-db-initialization.ps1 -Build"
}

$documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
$dbDir = Join-Path $documents "OfficeAiAppData-Debug"
$dbPath = Join-Path $dbDir "office_ai.db"
if (-not (Test-Path $dbDir)) {
    New-Item -ItemType Directory -Path $dbDir | Out-Null
}

$backupDir = Join-Path $dbDir ("empty-db-smoke-backup-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $backupDir | Out-Null

$dbFiles = @($dbPath, "$dbPath-wal", "$dbPath-shm")
$movedOriginals = @()

try {
    foreach ($file in $dbFiles) {
        if (Test-Path $file) {
            $target = Join-Path $backupDir (Split-Path -Leaf $file)
            Move-Item -LiteralPath $file -Destination $target
            $movedOriginals += [pscustomobject]@{ Source = $file; Backup = $target }
        }
    }

    $childScript = Join-Path $backupDir "verify-empty-db.ps1"
    $childContent = @'
param(
    [string]$Bin,
    [string]$Dll
)

$ErrorActionPreference = "Stop"

Push-Location $Bin
try {
    Add-Type -Path $Dll
    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()

    $conn = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT version FROM schema_version LIMIT 1"
        $version = [int]$cmd.ExecuteScalar()
        if ($version -ne 11) {
            throw "Expected schema version 11, got $version"
        }

        $requiredTables = @(
            "atomic_memory",
            "user_profile",
            "conversation_event",
            "memory_item",
            "memory_embedding",
            "memory_job",
            "skills_registry",
            "agent_run",
            "agent_run_step",
            "data_migration_marker"
        )

        foreach ($table in $requiredTables) {
            $check = $conn.CreateCommand()
            $check.CommandText = "SELECT COUNT(1) FROM sqlite_master WHERE type = 'table' AND name = @name"
            [void]$check.Parameters.AddWithValue("@name", $table)
            if ([int]$check.ExecuteScalar() -ne 1) {
                throw "Expected table not found: $table"
            }
        }
    }
    finally {
        if ($conn) {
            $conn.Dispose()
        }
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }

    [pscustomobject]@{
        Database = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()
        SchemaVersion = 11
        RequiredTables = 10
        EmptyInitialization = $true
    } | Format-List
}
finally {
    Pop-Location
}
'@
    Set-Content -LiteralPath $childScript -Value $childContent -Encoding UTF8

    & powershell -NoProfile -ExecutionPolicy Bypass -File $childScript -Bin $bin -Dll $dll
    if ($LASTEXITCODE -ne 0) {
        throw "Empty DB initialization child process failed, exit code: $LASTEXITCODE"
    }
}
finally {
    $generatedDir = Join-Path $backupDir "generated"
    New-Item -ItemType Directory -Path $generatedDir -Force | Out-Null

    foreach ($file in $dbFiles) {
        if (Test-Path $file) {
            Move-Item -LiteralPath $file -Destination (Join-Path $generatedDir (Split-Path -Leaf $file))
        }
    }

    foreach ($item in $movedOriginals) {
        if (Test-Path $item.Backup) {
            Move-Item -LiteralPath $item.Backup -Destination $item.Source
        }
    }
}
