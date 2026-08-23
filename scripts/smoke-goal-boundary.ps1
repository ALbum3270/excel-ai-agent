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

$instanceNonPublic = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
$captureRawMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod("CaptureRawUserRequest", $instanceNonPublic)
$setGoalMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod("SetGoalContractOnce", $instanceNonPublic)
if ($null -eq $captureRawMethod -or $null -eq $setGoalMethod) {
    throw "AgentTaskSpec goal attachment seam is missing."
}

function Set-RawRequest {
    param($Spec, [string]$Value)
    $captureRawMethod.Invoke($Spec, @($Value)) | Out-Null
}

function Set-FrozenGoal {
    param($Spec, $Goal)
    $setGoalMethod.Invoke($Spec, @($Goal)) | Out-Null
}

$compilerReturnsFrozenGoal = @([ShareRibbon.Agent.Goals.GoalCompiler].GetMethods() | Where-Object {
    $_.ReturnType -eq [ShareRibbon.Agent.Goals.GoalContract]
})
if ($compilerReturnsFrozenGoal.Count -ne 0) {
    throw "GoalCompiler bypasses validation by returning a frozen GoalContract."
}

function Decode-JsonString {
    param([string]$Json)
    return ($Json | ConvertFrom-Json)
}

function New-SourceClause {
    param(
        [string]$Id,
        [string]$Text,
        [string]$RequiredCapability = ""
    )

    $clause = [ShareRibbon.Agent.Goals.CandidateGoalSourceClause]::new()
    $clause.Id = $Id
    $clause.Text = $Text
    $clause.IsExplicit = $true
    $clause.RequiredCapability = $RequiredCapability
    return $clause
}

function New-Criterion {
    param(
        [string]$Id,
        [string]$Statement,
        [string]$Kind,
        [string[]]$SourceClauseIds,
        [string]$CapabilityId = ""
    )

    $criterion = [ShareRibbon.Agent.Goals.CandidateGoalCriterion]::new()
    $criterion.Id = $Id
    $criterion.Statement = $Statement
    $criterion.Kind = $Kind
    $criterion.Required = $true
    $criterion.CapabilityId = $CapabilityId
    foreach ($sourceClauseId in $SourceClauseIds) {
        $criterion.SourceClauseIds.Add($sourceClauseId)
    }
    return $criterion
}

$computeAndWriteText = Decode-JsonString '"\u8ba1\u7b97\u6bcf\u4e2a\u5730\u533a\u5e73\u5747\u9500\u552e\u989d\u5e76\u5199\u5230 Summary"'
$computeText = Decode-JsonString '"\u8ba1\u7b97\u6bcf\u4e2a\u5730\u533a\u5e73\u5747\u9500\u552e\u989d"'
$writeText = Decode-JsonString '"\u5199\u5230 Summary"'
$pythonText = Decode-JsonString '"\u7528 Python \u8ba1\u7b97\u6bcf\u4e2a\u5730\u533a\u5e73\u5747\u9500\u552e\u989d"'
$semanticText = Decode-JsonString '"\u7ed3\u679c\u8981\u4e13\u4e1a\u4e00\u4e9b\uff0c\u9002\u5408\u5411\u7ba1\u7406\u5c42\u5c55\u793a"'
$differentGoalText = Decode-JsonString '"\u53e6\u4e00\u4e2a\u7528\u6237\u76ee\u6807"'

# Faithfulness: a destination-only assertion cannot cover an omitted computation clause.
$omittedCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$omittedCandidate.RawUserRequest = $computeAndWriteText
$omittedCandidate.SourceClauses.Add((New-SourceClause "clause-compute" $computeText))
$omittedCandidate.SourceClauses.Add((New-SourceClause "clause-write" $writeText))
$omittedCandidate.Criteria.Add((New-Criterion "criterion-summary" "Summary exists" "outcome" @("clause-write")))
$omittedCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($omittedCandidate)
$omittedValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($omittedCompilation)
if ($omittedValidation.Succeeded -or
    -not (($omittedValidation.Errors -join " ") -match "clause-compute")) {
    throw "Goal coverage accepted a candidate that omitted the computation clause."
}

# Faithfulness: an explicitly required capability must survive as policy and criterion.
$capabilityCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$capabilityCandidate.RawUserRequest = $pythonText
$capabilityCandidate.SourceClauses.Add((New-SourceClause "clause-python" $pythonText "PythonCompute"))
$capabilityCandidate.Criteria.Add((New-Criterion "criterion-python-text" $pythonText "semantic" @("clause-python")))
$capabilityCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($capabilityCandidate)
$capabilityValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($capabilityCompilation)
if ($capabilityValidation.Succeeded -or
    -not (($capabilityValidation.Errors -join " ") -match "PythonCompute")) {
    throw "Goal coverage accepted a candidate that omitted an explicitly required capability."
}

# Preservation: unstructured language survives verbatim instead of being guessed into style.
$semanticCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($semanticText)
$semanticValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($semanticCompilation)
if (-not $semanticValidation.Succeeded) {
    throw "Verbatim semantic preservation failed coverage: $($semanticValidation.Errors -join '; ')"
}
$semanticContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $semanticCompilation,
    $semanticValidation)
$preserved = @($semanticContract.Criteria | Where-Object {
    $_.Kind -eq "semantic" -and $_.Statement -ceq $semanticText
})
if ($preserved.Count -ne 1 -or $semanticContract.RawUserRequest -cne $semanticText) {
    throw "Raw semantic criterion was not preserved verbatim in the frozen goal."
}

# Immutability: the authoritative interface has no setters or goal-relaxation methods.
$writableProperties = @([ShareRibbon.Agent.Goals.GoalContract].GetProperties() | Where-Object { $_.CanWrite })
$mutationMethods = @([ShareRibbon.Agent.Goals.GoalContract].GetMethods() | Where-Object {
    $_.Name -match "^(Remove|Replace|Relax|Set|Add)"
})
if ($writableProperties.Count -ne 0 -or $mutationMethods.Count -ne 0) {
    throw "Frozen GoalContract still exposes a mutation interface."
}

$criteriaList = [System.Collections.IList]$semanticContract.Criteria
if (-not $criteriaList.IsReadOnly -or -not $criteriaList.IsFixedSize) {
    throw "Frozen GoalContract collection is not exposed as read-only/fixed-size."
}

# Freeze is a defensive copy and hashing is canonical rather than list-order dependent.
$hashBeforeCandidateMutation = $semanticContract.ContractHash
$semanticCompilation.Candidate.Criteria[0].Statement = "weakened after freeze"
if ($semanticContract.Criteria[0].Statement -cne $semanticText -or
    $semanticContract.ContractHash -ne $hashBeforeCandidateMutation) {
    throw "Frozen GoalContract changed when its candidate was mutated."
}

$semanticCompilationAgain = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($semanticText)
$semanticValidationAgain = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($semanticCompilationAgain)
$semanticContractAgain = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $semanticCompilationAgain,
    $semanticValidationAgain)
if ($semanticContractAgain.ContractHash -ne $hashBeforeCandidateMutation -or
    $semanticContractAgain.GoalId -ne $semanticContract.GoalId) {
    throw "Equivalent goal compilations produced unstable identity/hash values."
}

# AgentTaskSpec permits one idempotent attachment, never replacement.
$taskSpecWritableAuthority = @([ShareRibbon.Agent.AgentTaskSpec].GetProperties() | Where-Object {
    ($_.Name -eq "RawUserRequest" -or $_.Name -eq "GoalContract") -and $_.CanWrite
})
if ($taskSpecWritableAuthority.Count -ne 0) {
    throw "AgentTaskSpec exposes a public setter for authoritative goal state."
}
$publicAuthorityMethods = @([ShareRibbon.Agent.AgentTaskSpec].GetMethods() | Where-Object {
    $_.Name -eq "CaptureRawUserRequest" -or $_.Name -eq "SetGoalContractOnce"
})
if ($publicAuthorityMethods.Count -ne 0) {
    throw "AgentTaskSpec exposes public goal attachment methods."
}

$spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
Set-RawRequest $spec $semanticText
Set-FrozenGoal $spec $semanticContract
Set-FrozenGoal $spec $semanticContractAgain

$differentCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($differentGoalText)
$differentValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($differentCompilation)
$differentContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $differentCompilation,
    $differentValidation)
$replacementRejected = $false
try {
    Set-FrozenGoal $spec $differentContract
}
catch {
    $replacementRejected = $_.Exception.ToString() -match "cannot be replaced"
}
if (-not $replacementRejected -or $spec.GoalContract.ContractHash -ne $hashBeforeCandidateMutation) {
    throw "AgentTaskSpec allowed a frozen goal to be replaced."
}

$mismatchSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
Set-RawRequest $mismatchSpec $semanticText
$mismatchRejected = $false
try {
    Set-FrozenGoal $mismatchSpec $differentContract
}
catch {
    $mismatchRejected = $_.Exception.ToString() -match "does not represent"
}
if (-not $mismatchRejected -or $null -ne $mismatchSpec.GoalContract) {
    throw "AgentTaskSpec accepted a GoalContract compiled from a different raw request."
}

$uncapturedSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$uncapturedRejected = $false
try {
    Set-FrozenGoal $uncapturedSpec $semanticContract
}
catch {
    $uncapturedRejected = $_.Exception.ToString() -match "must be captured"
}
if (-not $uncapturedRejected -or $null -ne $uncapturedSpec.GoalContract) {
    throw "AgentTaskSpec accepted a GoalContract before capturing RawUserRequest."
}

# Legacy OutcomeContract is not accepted by any Goal compiler/freezer method.
$legacyInputMethods = @(
    [ShareRibbon.Agent.Goals.GoalCompiler].GetMethods(),
    [ShareRibbon.Agent.Goals.GoalContractFreezer].GetMethods()
) | ForEach-Object { $_ } | Where-Object {
    @($_.GetParameters() | Where-Object { $_.ParameterType -eq [ShareRibbon.Agent.OutcomeContract] }).Count -gt 0
}
if ($legacyInputMethods.Count -ne 0) {
    throw "Legacy OutcomeContract can still flow back into authoritative goal semantics."
}

Write-Host "PASS: immutable, traceable Goal Boundary contracts"
