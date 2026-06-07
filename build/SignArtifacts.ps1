param(
    [string]$CertificateThumbprint = $env:OFFICE_AI_SIGN_CERT_THUMBPRINT,
    [string]$PfxPath = $env:OFFICE_AI_SIGN_PFX,
    [string]$PfxPassword = $env:OFFICE_AI_SIGN_PFX_PASSWORD,
    [string]$TimestampUrl = 'http://timestamp.digicert.com',
    [switch]$AllowMissing,
    [switch]$SkipVerify,
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

$signArgs = @('sign', '/fd', 'SHA256', '/tr', $TimestampUrl, '/td', 'SHA256')

if (-not [string]::IsNullOrWhiteSpace($PfxPath)) {
    $resolvedPfx = Resolve-Path $PfxPath
    $signArgs += @('/f', $resolvedPfx)
    if (-not [string]::IsNullOrEmpty($PfxPassword)) {
        $signArgs += @('/p', $PfxPassword)
    }
} elseif (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    $signArgs += @('/sha1', $CertificateThumbprint)
} else {
    throw 'Set OFFICE_AI_SIGN_CERT_THUMBPRINT or OFFICE_AI_SIGN_PFX before signing.'
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

    Write-Host "Signing $artifact"
    & $signtool.Source @signArgs $path
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed for $artifact"
    }

    if (-not $SkipVerify) {
        Write-Host "Verifying $artifact"
        & $signtool.Source verify /pa /tw $path
        if ($LASTEXITCODE -ne 0) {
            throw "signtool verify failed for $artifact"
        }
    }
}

Write-Host 'Signing complete.'
