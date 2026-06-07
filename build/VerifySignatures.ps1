param(
    [switch]$AllowMissing,
    [string[]]$Artifacts = @(
        'OfficeAgent\Release\OfficeAgent.msi',
        'ExcelAi\bin\Release\ExcelAi.dll',
        'WordAi\bin\Release\WordAi.dll',
        'PowerPointAi\bin\Release\PowerPointAi.dll',
        'ShareRibbon\bin\Release\ShareRibbon.dll'
    )
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$signtool = Get-Command signtool.exe -ErrorAction SilentlyContinue

if ($null -eq $signtool) {
    $kitRoot = "${env:ProgramFiles(x86)}\Windows Kits\10\bin"
    if (Test-Path $kitRoot) {
        $signtoolPath = Get-ChildItem -Path $kitRoot -Recurse -Filter signtool.exe |
            Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($signtoolPath) {
            $signtool = $signtoolPath
        }
    }
}

if ($null -eq $signtool) {
    throw 'signtool.exe not found. Install Windows SDK or add signtool.exe to PATH.'
}

foreach ($artifact in $Artifacts) {
    $path = Join-Path $repoRoot $artifact
    if (-not (Test-Path -LiteralPath $path)) {
        if ($AllowMissing) {
            Write-Warning "Skip missing artifact: $artifact"
            continue
        }
        throw "Missing artifact: $artifact"
    }

    Write-Host "Verifying signature: $artifact"
    & $signtool.Source verify /pa /tw $path
    if ($LASTEXITCODE -ne 0) {
        throw "Signature verification failed for $artifact"
    }
}

Write-Host 'Signature verification complete.'
