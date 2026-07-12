param(
    [string]$VdprojPath = "",
    [switch]$IncludePlainFiles
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
if ([string]::IsNullOrWhiteSpace($VdprojPath)) {
    $VdprojPath = Join-Path $repoRoot "OfficeAgent\OfficeAgent.vdproj"
}

if (-not (Test-Path -LiteralPath $VdprojPath)) {
    throw "Installer project not found: $VdprojPath"
}

$installerDir = Split-Path -Parent (Resolve-Path $VdprojPath)
$content = Get-Content -LiteralPath $VdprojPath -Encoding Default
$missing = New-Object System.Collections.Generic.List[object]
$checked = 0

foreach ($line in $content) {
    if ($line -notmatch '^\s*"SourcePath"\s*=\s*"8:(.*)"') {
        continue
    }

    $sourcePath = $Matches[1]
    if ([string]::IsNullOrWhiteSpace($sourcePath)) {
        continue
    }

    if ($sourcePath.StartsWith("<")) {
        continue
    }

    $hasDirectoryPart = $sourcePath.Contains("\") -or $sourcePath.Contains("/")
    if (-not $hasDirectoryPart -and -not $IncludePlainFiles) {
        continue
    }

    if ([System.IO.Path]::IsPathRooted($sourcePath)) {
        $resolved = $sourcePath
    }
    else {
        $resolved = Join-Path $installerDir $sourcePath
    }

    $checked += 1
    if (-not (Test-Path -LiteralPath $resolved)) {
        $missing.Add([pscustomobject]@{
            SourcePath = $sourcePath
            ResolvedPath = $resolved
        })
    }
}

if ($missing.Count -gt 0) {
    $uniqueMissing = @($missing | Sort-Object SourcePath -Unique)
    Write-Host "Installer input audit failed. Missing files: $($uniqueMissing.Count)"
    $uniqueMissing |
        Sort-Object SourcePath |
        Format-Table SourcePath, ResolvedPath -AutoSize |
        Out-String -Width 260 |
        Write-Host
    exit 1
}

Write-Host "Installer input audit passed. Checked $checked file references."
exit 0
