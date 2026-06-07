$ErrorActionPreference = 'Stop'
$vdproj = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')) 'OfficeAgent\OfficeAgent.vdproj'
$content = Get-Content -LiteralPath $vdproj -Raw -Encoding Default

$issues = New-Object System.Collections.Generic.List[string]

if ($content -match '\\bin\\Debug\\') {
    $issues.Add('OfficeAgent.vdproj still references bin\Debug outputs.')
}

if ($content -match '"AllowLaterVersions"\s*=\s*"11:FALSE"') {
    $issues.Add('OfficeAgent.vdproj still blocks later .NET Framework versions.')
}

$debugSymbols = Select-String -Path $vdproj -Pattern '"TargetName"\s*=\s*"8:.*\.(pdb|xml)"'
if ($debugSymbols.Count -gt 0) {
    $issues.Add("OfficeAgent.vdproj still includes $($debugSymbols.Count) pdb/xml target entries. Remove these through the setup project UI or a vdproj-aware cleanup.")
}

if ($issues.Count -gt 0) {
    foreach ($issue in $issues) {
        Write-Warning $issue
    }
    exit 1
}

Write-Host 'Installer Release audit passed.'
