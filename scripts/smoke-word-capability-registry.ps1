$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$registryPath = Join-Path $repoRoot 'WordAi\Services\WordCapabilityRegistry.vb'
$harnessPath = Join-Path $repoRoot 'WordAi\Services\WordActionHarness.vb'
$projectPath = Join-Path $repoRoot 'WordAi\WordAi.vbproj'

if (-not (Test-Path -LiteralPath $registryPath)) {
    throw "Word capability registry not found: $registryPath"
}

$registry = Get-Content -LiteralPath $registryPath -Raw
$harness = Get-Content -LiteralPath $harnessPath -Raw
$project = Get-Content -LiteralPath $projectPath -Raw

if ($project -notmatch 'Services\\WordCapabilityRegistry\.vb') {
    throw 'WordCapabilityRegistry.vb is not registered in WordAi.vbproj.'
}

$requiredCapabilities = @(
    'word.proofread',
    'word.direct-formatting',
    'word.numbering',
    'word.semantic-reformat'
)

foreach ($capability in $requiredCapabilities) {
    if ($registry -notmatch [regex]::Escape($capability)) {
        throw "Missing Word capability descriptor: $capability"
    }
}

$requiredContracts = @(
    'InputSchema',
    'RiskLevel',
    'SupportsPreview',
    'SupportsUndo',
    'ObserveContract',
    'RepairContract',
    'ExplainContract',
    'ExampleRequests'
)

foreach ($contract in $requiredContracts) {
    if ($registry -notmatch [regex]::Escape($contract)) {
        throw "Missing capability contract field: $contract"
    }
}

if ($harness -notmatch 'WordCapabilityRegistry\.Require') {
    throw 'WordActionHarness does not attach capability descriptors to plans.'
}

if ($harness -notmatch 'ProofreadIntentCompiler\.LooksLikeProofreadCommand') {
    throw 'WordActionHarness does not route explicit proofread requests through ProofreadIntentCompiler.'
}

Write-Host 'Word capability registry smoke passed.'
