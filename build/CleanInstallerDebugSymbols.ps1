$ErrorActionPreference = 'Stop'
$vdproj = Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..')) 'OfficeAgent\OfficeAgent.vdproj'
$lines = [System.Collections.Generic.List[string]](Get-Content -LiteralPath $vdproj -Encoding Default)
$output = New-Object System.Collections.Generic.List[string]
$removed = 0
$i = 0

while ($i -lt $lines.Count) {
    $line = $lines[$i]
    $isBlockHeader = $line -match '^\s*"\{[0-9A-Fa-f-]+\}:[^"]+"\s*$'
    $hasOpenBrace = ($i + 1 -lt $lines.Count) -and ($lines[$i + 1] -match '^\s*\{\s*$')

    if ($isBlockHeader -and $hasOpenBrace) {
        $block = New-Object System.Collections.Generic.List[string]
        $block.Add($lines[$i])
        $i++
        $block.Add($lines[$i])
        $i++

        while ($i -lt $lines.Count) {
            $block.Add($lines[$i])
            if ($lines[$i] -match '^\s*\}\s*$') {
                $i++
                break
            }
            $i++
        }

        $blockText = [string]::Join("`n", $block)
        if ($blockText -match '"TargetName"\s*=\s*"8:[^"]+\.(pdb|xml)"') {
            $removed++
            continue
        }

        foreach ($blockLine in $block) {
            $output.Add($blockLine)
        }
        continue
    }

    $output.Add($line)
    $i++
}

Set-Content -LiteralPath $vdproj -Value $output -Encoding Default
Write-Host "Removed $removed pdb/xml installer file blocks."
