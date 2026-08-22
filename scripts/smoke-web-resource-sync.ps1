param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$assemblyPath = Join-Path $repoRoot "ShareRibbon\bin\$Configuration\ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $assemblyPath)) {
    throw "ShareRibbon assembly not found: $assemblyPath. Build the code projects first."
}

Add-Type -Path $assemblyPath

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("office-ai-resource-sync-" + [Guid]::NewGuid().ToString("N"))
try {
    New-Item -ItemType Directory -Path (Join-Path $testRoot "js") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $testRoot "css") -Force | Out-Null

    # Reproduce the production failure: the old extractor trusted this shared marker and
    # skipped every file even when an embedded frontend resource had changed.
    Set-Content -LiteralPath (Join-Path $testRoot ".version") -Value "2026.07.18.2" -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $testRoot "js\agent-card.js") -Value "stale-agent-card" -Encoding UTF8

    $resolvedRoot = [ShareRibbon.ResourceExtractor]::ExtractResources($testRoot)
    if ([string]::IsNullOrWhiteSpace($resolvedRoot)) {
        throw "Resource extraction returned an empty root"
    }

    $expected = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Resources\js\agent-card.js") -Raw -Encoding UTF8
    $actualPath = Join-Path $resolvedRoot "js\agent-card.js"
    $actual = Get-Content -LiteralPath $actualPath -Raw -Encoding UTF8
    if (-not [string]::Equals($expected, $actual, [StringComparison]::Ordinal)) {
        throw "A matching global version marker still leaves an individual frontend resource stale"
    }

    # Synchronisation should be content-addressed: a second pass must not rewrite a file
    # whose bytes already match the embedded resource.
    $firstWrite = (Get-Item -LiteralPath $actualPath).LastWriteTimeUtc
    Start-Sleep -Milliseconds 30
    [void][ShareRibbon.ResourceExtractor]::ExtractResources($testRoot)
    $secondWrite = (Get-Item -LiteralPath $actualPath).LastWriteTimeUtc
    if ($secondWrite -ne $firstWrite) {
        throw "Unchanged frontend resources are rewritten on every startup"
    }

    Write-Host "PASS: embedded web resources replace stale files and leave matching files untouched"
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
