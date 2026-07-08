param(
    [string]$VdprojPath = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($VdprojPath)) {
    $VdprojPath = Join-Path $repoRoot "OfficeAgent\OfficeAgent.vdproj"
}

if (-not (Test-Path -LiteralPath $VdprojPath)) {
    throw "Installer project not found: $VdprojPath"
}

$lines = Get-Content -LiteralPath $VdprojPath
$items = New-Object System.Collections.Generic.List[object]
$currentName = $null
$currentSource = $null
$currentFolder = $null

foreach ($line in $lines) {
    if ($line -match '^\s*"Name"\s*=\s*"8:(.+\.dll)"') {
        $currentName = $Matches[1]
        continue
    }

    if ($line -match '^\s*"SourcePath"\s*=\s*"8:(.+\.dll)"') {
        $currentSource = $Matches[1]
        continue
    }

    if ($line -match '^\s*"Folder"\s*=\s*"8:(.+)"') {
        $currentFolder = $Matches[1]
        if (-not [string]::IsNullOrWhiteSpace($currentName)) {
            $items.Add([pscustomobject]@{
                Name = $currentName
                SourcePath = $currentSource
                Folder = $currentFolder
            })
        }
        $currentName = $null
        $currentSource = $null
        $currentFolder = $null
    }
}

$duplicates = $items |
    Group-Object Name |
    Where-Object { $_.Count -gt 1 } |
    Sort-Object Count -Descending

if ($duplicates.Count -eq 0) {
    "No duplicate DLL entries found."
    exit 0
}

foreach ($group in $duplicates) {
    ""
    "== $($group.Name) ($($group.Count)x) =="
    $group.Group |
        Sort-Object SourcePath, Folder |
        Select-Object SourcePath, Folder |
        Format-Table -AutoSize | Out-String -Width 220
}

