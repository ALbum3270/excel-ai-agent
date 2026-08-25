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
$setCompilationMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod("SetGoalCompilationOnce", $instanceNonPublic)
if ($null -eq $captureRawMethod -or $null -eq $setGoalMethod -or $null -eq $setCompilationMethod) {
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

function Set-GoalCompilation {
    param($Spec, $Compilation)
    $setCompilationMethod.Invoke($Spec, @($Compilation)) | Out-Null
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
        [string]$Text
    )

    $clause = [ShareRibbon.Agent.Goals.CandidateGoalSourceClause]::new()
    $clause.Id = $Id
    $clause.Text = $Text
    $clause.IsExplicit = $true
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

# Faithfulness: every explicit clause survives even when the structured interpretation omitted it.
$omittedCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$omittedCandidate.RawUserRequest = $computeAndWriteText
$omittedCandidate.SourceClauses.Add((New-SourceClause "clause-compute" $computeText))
$omittedCandidate.SourceClauses.Add((New-SourceClause "clause-write" $writeText))
$summaryHint = New-Criterion "criterion-summary" "Summary exists" "outcome" @("clause-write")
$summaryHint.Required = $false
$omittedCandidate.Criteria.Add($summaryHint)
$sourceCountBeforeCompile = $omittedCandidate.SourceClauses.Count
$criterionCountBeforeCompile = $omittedCandidate.Criteria.Count
$omittedCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($omittedCandidate)
$omittedValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($omittedCompilation)
if (-not $omittedValidation.Succeeded) {
    throw "Clause-level verbatim fallback failed: $($omittedValidation.Errors -join '; ')"
}
$preservedCompute = @($omittedCompilation.Candidate.Criteria | Where-Object {
    $_.Required -and $_.Kind -eq "semantic" -and $_.Statement -ceq $computeText -and $_.SourceClauseIds.Contains("clause-compute")
})
if ($preservedCompute.Count -ne 1) {
    throw "An omitted structured computation clause was not retained as exact semantic authority."
}
if ($omittedCandidate.SourceClauses.Count -ne $sourceCountBeforeCompile -or
    $omittedCandidate.Criteria.Count -ne $criterionCountBeforeCompile) {
    throw "GoalCompiler mutated the caller-owned candidate while adding preservation criteria."
}

# Assumptions are diagnostic only and cannot become required goal semantics.
$assumptionText = Decode-JsonString '"\u53ea\u7edf\u8ba1\u9500\u552e\u989d>1000"'
$averageText = Decode-JsonString '"\u8ba1\u7b97\u5e73\u5747\u9500\u552e\u989d"'
$assumptionCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$assumptionCandidate.RawUserRequest = $averageText
$assumptionCandidate.SourceClauses.Add((New-SourceClause "clause-average" $averageText))
$assumptionCandidate.Criteria.Add((New-Criterion "criterion-invented" $assumptionText "compute" @()))
$assumptionCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile(
    $assumptionCandidate,
    [string[]]@(),
    [string[]]@($assumptionText),
    [string[]]@(),
    $false)
$assumptionValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($assumptionCompilation)
if (-not $assumptionValidation.Succeeded -or
    @($assumptionCompilation.Candidate.Criteria | Where-Object { $_.Id -eq "criterion-invented" -and $_.Required }).Count -ne 0) {
    throw "A non-authoritative model assumption was not demoted to a planning hint: $($assumptionValidation.Errors -join '; ')"
}

# A paraphrased assumption cannot become required merely by referencing a genuine clause.
$paraphraseCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$paraphraseCandidate.RawUserRequest = $averageText
$paraphraseCandidate.SourceClauses.Add((New-SourceClause "clause-average" $averageText))
$paraphraseCandidate.Criteria.Add((New-Criterion "criterion-paraphrase" "only include sales above 1000" "compute" @("clause-average")))
$paraphraseCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile(
    $paraphraseCandidate,
    [string[]]@(),
    [string[]]@("only include sales > 1000"),
    [string[]]@(),
    $false)
$paraphraseValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($paraphraseCompilation)
if (-not $paraphraseValidation.Succeeded -or
    @($paraphraseCompilation.Candidate.Criteria | Where-Object { $_.Id -eq "criterion-paraphrase" -and $_.Required }).Count -ne 0) {
    throw "A paraphrased assumption was not demoted before freeze."
}

# Verbatim text does not authorize model-invented type/verifier metadata. The compiler keeps
# the exact user statement while deterministically normalizing every executable label.
$metadataCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$metadataCandidate.RawUserRequest = $averageText
$metadataCandidate.SourceClauses.Add((New-SourceClause "metadata-source" $averageText))
$metadataCriterion = New-Criterion "metadata-criterion" $averageText "delete-everything" @("metadata-source") "invented"
$metadataCriterion.VerificationCapability = "ExecuteVBA"
$metadataCandidate.Criteria.Add($metadataCriterion)
$metadataConstraint = [ShareRibbon.Agent.Goals.CandidateGoalConstraint]::new()
$metadataConstraint.Id = "metadata-constraint"
$metadataConstraint.Statement = $averageText
$metadataConstraint.Kind = "run-arbitrary-code"
$metadataConstraint.Required = $true
$metadataConstraint.SourceClauseIds.Add("metadata-source")
$metadataCandidate.Constraints.Add($metadataConstraint)
$metadataCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($metadataCandidate)
$metadataValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($metadataCompilation)
$metadataContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($metadataCompilation, $metadataValidation)
if (-not $metadataValidation.Succeeded -or
    $metadataContract.Criteria[0].Kind -ne "semantic" -or
    $metadataContract.Criteria[0].VerificationCapability -ne "semantic" -or
    -not [string]::IsNullOrEmpty($metadataContract.Criteria[0].CapabilityId) -or
    $metadataContract.Constraints[0].Kind -ne "semantic") {
    throw "Untrusted criterion/constraint metadata entered the frozen goal authority."
}

# The public validation seam independently rejects callers that bypass the compiler demotion.
$bypassCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$bypassCandidate.RawUserRequest = $averageText
$bypassSource = New-SourceClause "bypass-source" $averageText
$bypassSource.SourceStart = 0
$bypassCandidate.SourceClauses.Add($bypassSource)
$bypassCandidate.Criteria.Add((New-Criterion "bypass-semantic" $averageText "semantic" @("bypass-source")))
$bypassCandidate.Criteria.Add((New-Criterion "bypass-paraphrase" "only include sales above 1000" "compute" @("bypass-source")))
$bypassCompilation = [ShareRibbon.Agent.Goals.GoalCompilationResult]::new(
    $bypassCandidate,
    [ShareRibbon.Agent.Goals.GoalCoverageMapEntry[]]@(),
    [string[]]@(), [string[]]@("only include sales > 1000"), [string[]]@(), $false)
$bypassValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($bypassCompilation)
if ($bypassValidation.Succeeded -or
    -not (($bypassValidation.Errors -join " ") -match "model paraphrase")) {
    throw "The validator accepted a paraphrased assumption through a genuine source reference."
}

# Public compilation inputs with a null raw request fail as validation data, never as a NullReferenceException.
$nullRawCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$nullRawCandidate.RawUserRequest = $null
$nullRawCompilation = [ShareRibbon.Agent.Goals.GoalCompilationResult]::new(
    $nullRawCandidate,
    [ShareRibbon.Agent.Goals.GoalCoverageMapEntry[]]@(),
    [string[]]@(),
    [string[]]@(),
    [string[]]@(),
    $false)
$nullRawValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($nullRawCompilation)
if ($nullRawValidation.Succeeded -or
    -not (($nullRawValidation.Errors -join " ") -match "RawUserRequest")) {
    throw "Null RawUserRequest did not fail closed as structured validation data."
}

# A model cannot fabricate the raw evidence that allegedly authorized a clause.
$fabricatedCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$fabricatedCandidate.RawUserRequest = $averageText
$fabricatedCandidate.SourceClauses.Add((New-SourceClause "clause-fabricated" $assumptionText))
$fabricatedCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($fabricatedCandidate)
$fabricatedValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($fabricatedCompilation)
if ($fabricatedValidation.Succeeded -or
    -not (($fabricatedValidation.Errors -join " ") -match "verified exact occurrence")) {
    throw "Goal coverage accepted a fabricated source clause that is absent from RawUserRequest."
}

# Exact substring origin is not enough when the selected span drops a governing modifier.
$negativeDeleteText = Decode-JsonString '"\u4e0d\u8981\u5220\u9664\u9500\u552e\u6570\u636e\u5de5\u4f5c\u8868"'
$positiveDeleteFragment = Decode-JsonString '"\u5220\u9664\u9500\u552e\u6570\u636e\u5de5\u4f5c\u8868"'
$polarityCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$polarityCandidate.RawUserRequest = $negativeDeleteText
$polarityCandidate.SourceClauses.Add((New-SourceClause "clause-inverted" $positiveDeleteFragment))
$polarityCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($polarityCandidate)
$polarityValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($polarityCompilation)
if ($polarityValidation.Succeeded -or
    -not (($polarityValidation.Errors -join " ") -match "incomplete semantic span")) {
    throw "Goal coverage promoted an exact fragment after dropping its governing negation."
}

# Faithfulness: explicit method language is independently resolved even if the model omits it.
$capabilityCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$capabilityCandidate.RawUserRequest = $pythonText
$capabilityCandidate.SourceClauses.Add((New-SourceClause "clause-python" $pythonText))
$capabilityCandidate.Criteria.Add((New-Criterion "criterion-python-text" $pythonText "semantic" @("clause-python")))
$capabilityCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($capabilityCandidate)
$capabilityValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($capabilityCompilation)
if (-not $capabilityValidation.Succeeded -or
    -not $capabilityCompilation.Candidate.RequiredCapabilities.Contains("PythonCompute")) {
    throw "Independent capability evidence failed to recover an explicit Python requirement: $($capabilityValidation.Errors -join '; ')"
}
$capabilityContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($capabilityCompilation, $capabilityValidation)

# Method capabilities are hard constraints only when the user affirmatively requires them.
# Negated, optional, meta-question and identifier-only mentions must never be promoted into
# RequiredCapabilities merely because they contain an engine name and an invocation verb.
$nonRequiredMethodRequests = @(
    (Decode-JsonString '"\u4e0d\u8981\u4f7f\u7528 Python \u8ba1\u7b97\uff0c\u6539\u7528\u516c\u5f0f\u3002"'),
    (Decode-JsonString '"\u5982\u679c\u53ef\u4ee5\uff0c\u8003\u8651\u4f7f\u7528 Python \u8ba1\u7b97\u3002"'),
    (Decode-JsonString '"\u4e3a\u4ec0\u4e48\u4f7f\u7528 Python \u8ba1\u7b97\u8fd9\u4e48\u6162\uff1f"'),
    (Decode-JsonString '"\u65b0\u5efa\u5de5\u4f5c\u8868 Python\u5e73\u5747\u503c0821"'),
    (Decode-JsonString '"\u4e0d\u8981\u6267\u884c VBA\uff0c\u6539\u7528\u516c\u5f0f\u3002"'),
    "don't use Python to calculate; use a formula instead",
    'write the text "use Python to calculate" into A1'
)
foreach ($requestText in $nonRequiredMethodRequests) {
    $methodCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile([string]$requestText)
    if ($methodCompilation.Candidate.RequiredCapabilities.Count -ne 0) {
        throw "A non-affirmative method mention became a hard capability constraint: [$requestText] -> $($methodCompilation.Candidate.RequiredCapabilities -join ',')"
    }
}
$requiredVbaText = Decode-JsonString '"\u4f7f\u7528 VBA \u81ea\u52a8\u5316\u5904\u7406\u5f53\u524d\u5de5\u4f5c\u7c3f"'
$requiredVbaCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($requiredVbaText)
if (-not $requiredVbaCompilation.Candidate.RequiredCapabilities.Contains("ExecuteVBA")) {
    throw "Affirmative VBA execution was not preserved as a hard method constraint."
}
$multipleMethodsText = Decode-JsonString '"\u4f7f\u7528 Python \u8ba1\u7b97\uff0c\u5e76\u4f7f\u7528 VBA \u5199\u5165\u7ed3\u679c"'
$multipleMethodsCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($multipleMethodsText)
$multipleMethodsValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($multipleMethodsCompilation)
if (-not $multipleMethodsValidation.Succeeded -or
    -not $multipleMethodsCompilation.Candidate.RequiredCapabilities.Contains("PythonCompute") -or
    -not $multipleMethodsCompilation.Candidate.RequiredCapabilities.Contains("ExecuteVBA")) {
    throw "Independent capability evidence edges did not preserve multiple explicit method constraints: $($multipleMethodsValidation.Errors -join '; ')"
}

# Once frozen, GoalContract is the only capability authority; legacy projections cannot add policy.
$capabilitySpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
Set-RawRequest $capabilitySpec $pythonText
$capabilitySpec.RequiredCapabilities.Add("CreateSheet")
$capabilitySpec.MandatoryTools.Add("WriteData")
Set-FrozenGoal $capabilitySpec $capabilityContract
if (-not [ShareRibbon.Agent.AgentExecutionContract]::IsRequiredCapability($capabilitySpec, "PythonCompute") -or
    [ShareRibbon.Agent.AgentExecutionContract]::IsRequiredCapability($capabilitySpec, "CreateSheet") -or
    [ShareRibbon.Agent.AgentExecutionContract]::IsRequiredCapability($capabilitySpec, "WriteData")) {
    throw "Legacy capability projections still override the frozen GoalContract."
}
$legacyCapabilitySpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$legacyCapabilitySpec.RequiredCapabilities.Add("CreateSheet")
if (-not [ShareRibbon.Agent.AgentExecutionContract]::IsRequiredCapability($legacyCapabilitySpec, "CreateSheet")) {
    throw "Pre-GoalContract compatibility no longer reads persisted legacy capabilities."
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

# Freeze is a defensive copy; exact and semantic hash identities remain stable after freezing.
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

# ContractHash protects exact frozen representation; SemanticHash ignores only model labels and order.
$semanticRaw = Decode-JsonString '"\u8ba1\u7b97\u5e73\u5747\u9500\u552e\u989d\u5e76\u5199\u5165\u6c47\u603b"'
$semanticCompute = Decode-JsonString '"\u8ba1\u7b97\u5e73\u5747\u9500\u552e\u989d"'
$semanticWrite = Decode-JsonString '"\u5199\u5165\u6c47\u603b"'

function New-RenamedSemanticCandidate {
    param([bool]$Reverse)
    $candidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
    $candidate.RawUserRequest = $semanticRaw
    if (-not $Reverse) {
        $candidate.SourceClauses.Add((New-SourceClause "compute-a" $semanticCompute))
        $candidate.SourceClauses.Add((New-SourceClause "write-a" $semanticWrite))
        $candidate.Criteria.Add((New-Criterion "compute-criterion-a" $semanticCompute "compute" @("compute-a")))
        $candidate.Criteria.Add((New-Criterion "write-criterion-a" $semanticWrite "state" @("write-a")))
    }
    else {
        $candidate.SourceClauses.Add((New-SourceClause "output-z" $semanticWrite))
        $candidate.SourceClauses.Add((New-SourceClause "calculation-z" $semanticCompute))
        $candidate.Criteria.Add((New-Criterion "output-check-z" $semanticWrite "state" @("output-z")))
        $candidate.Criteria.Add((New-Criterion "calculation-check-z" $semanticCompute "compute" @("calculation-z")))
    }
    return $candidate
}

$semanticACompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile((New-RenamedSemanticCandidate $false))
$semanticAValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($semanticACompilation)
$semanticAContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($semanticACompilation, $semanticAValidation)
$semanticBCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile((New-RenamedSemanticCandidate $true))
$semanticBValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($semanticBCompilation)
$semanticBContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($semanticBCompilation, $semanticBValidation)
if ($semanticAContract.ContractHash -eq $semanticBContract.ContractHash -or
    $semanticAContract.SemanticHash -ne $semanticBContract.SemanticHash -or
    $semanticAContract.GoalId -ne $semanticBContract.GoalId) {
    throw "ContractHash and SemanticHash do not preserve their distinct representation/semantic meanings."
}

# Candidate validation is a true TOCTOU snapshot: list reordering invalidates stale authorization.
$staleCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile((New-RenamedSemanticCandidate $false))
$staleValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($staleCompilation)
$temporaryClause = $staleCompilation.Candidate.SourceClauses[0]
$staleCompilation.Candidate.SourceClauses[0] = $staleCompilation.Candidate.SourceClauses[1]
$staleCompilation.Candidate.SourceClauses[1] = $temporaryClause
$staleRejected = $false
try {
    [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($staleCompilation, $staleValidation) | Out-Null
}
catch {
    $staleRejected = $_.Exception.ToString() -match "does not match"
}
if (-not $staleRejected) {
    throw "Freeze accepted a candidate whose list order changed after validation."
}

# Line-ending representation changes ContractHash, but not semantic graph identity.
$lfText = "alpha`nbeta"
$crlfText = "alpha`r`nbeta"
$lfCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($lfText)
$lfContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($lfCompilation, [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($lfCompilation))
$crlfCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($crlfText)
$crlfContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($crlfCompilation, [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($crlfCompilation))
if ($lfContract.ContractHash -eq $crlfContract.ContractHash -or $lfContract.SemanticHash -ne $crlfContract.SemanticHash) {
    throw "Line-ending normalization leaked between ContractHash and SemanticHash."
}

# Duplicate clause text must not erase graph topology when model ids are ignored.
function New-DuplicateTopologyCandidate {
    param([bool]$SecondTopology)
    $duplicateRaw = Decode-JsonString '"\u7532|\u7532"'
    $duplicateText = Decode-JsonString '"\u7532"'
    $candidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
    $candidate.RawUserRequest = $duplicateRaw
    $firstClause = New-SourceClause "duplicate-1" $duplicateText
    $firstClause.SourceStart = 0
    $secondClause = New-SourceClause "duplicate-2" $duplicateText
    $secondClause.SourceStart = 2
    $candidate.SourceClauses.Add($firstClause)
    $candidate.SourceClauses.Add($secondClause)
    $candidate.Criteria.Add((New-Criterion "edge-a" $duplicateText "compute" @("duplicate-1")))
    $candidate.Criteria.Add((New-Criterion "edge-b" $duplicateText "compute" @($(if ($SecondTopology) { "duplicate-2" } else { "duplicate-1" }))))
    $candidate.Criteria.Add((New-Criterion "edge-c" $duplicateText "compute" @("duplicate-2")))
    return $candidate
}
$topologyOneCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile((New-DuplicateTopologyCandidate $false))
$topologyOneContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $topologyOneCompilation,
    [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($topologyOneCompilation))
$topologyTwoCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile((New-DuplicateTopologyCandidate $true))
$topologyTwoContract = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $topologyTwoCompilation,
    [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($topologyTwoCompilation))
if ($topologyOneContract.SemanticHash -eq $topologyTwoContract.SemanticHash) {
    throw "SemanticHash collapsed non-isomorphic references between duplicate source clauses."
}

# Canonical semantic identity must remain bounded for many repeated clauses.
$symmetricCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$symmetricCandidate.RawUserRequest = ((1..64 | ForEach-Object { "a" }) -join "|")
for ($sourceIndex = 0; $sourceIndex -lt 64; $sourceIndex++) {
    $source = New-SourceClause "repeat-$sourceIndex" "a"
    $source.SourceStart = $sourceIndex * 2
    $symmetricCandidate.SourceClauses.Add($source)
    $symmetricCandidate.Criteria.Add((New-Criterion "repeat-criterion-$sourceIndex" "a" "semantic" @("repeat-$sourceIndex")))
}
$symmetricCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($symmetricCandidate)
$symmetricValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($symmetricCompilation)
$symmetricTimer = [System.Diagnostics.Stopwatch]::StartNew()
[ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze($symmetricCompilation, $symmetricValidation) | Out-Null
$symmetricTimer.Stop()
if ($symmetricTimer.ElapsedMilliseconds -gt 1500) {
    throw "Semantic identity regressed to combinatorial search: $($symmetricTimer.ElapsedMilliseconds) ms"
}

# AgentTaskSpec permits one idempotent attachment, never replacement.
$taskSpecWritableAuthority = @([ShareRibbon.Agent.AgentTaskSpec].GetProperties() | Where-Object {
    ($_.Name -eq "RawUserRequest" -or $_.Name -eq "GoalContract") -and $_.CanWrite
})
if ($taskSpecWritableAuthority.Count -ne 0) {
    throw "AgentTaskSpec exposes a public setter for authoritative goal state."
}
$publicAuthorityMethods = @([ShareRibbon.Agent.AgentTaskSpec].GetMethods() | Where-Object {
    $_.Name -eq "CaptureRawUserRequest" -or
    $_.Name -eq "SetGoalCompilationOnce" -or
    $_.Name -eq "SetGoalContractOnce"
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

# Idempotence compares the whole compilation, including clarification/assumption state.
$compilationSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
Set-RawRequest $compilationSpec $averageText
$compilationA = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile(
    ([ShareRibbon.Agent.Goals.GoalCompiler]::Compile($averageText)).Candidate,
    [string[]]@(), [string[]]@("assumption-a"), [string[]]@(), $false)
$compilationB = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile(
    ([ShareRibbon.Agent.Goals.GoalCompiler]::Compile($averageText)).Candidate,
    [string[]]@(), [string[]]@("assumption-b"), [string[]]@(), $false)
Set-GoalCompilation $compilationSpec $compilationA
$compilationReplacementRejected = $false
try {
    Set-GoalCompilation $compilationSpec $compilationB
}
catch {
    $compilationReplacementRejected = $_.Exception.ToString() -match "cannot be replaced"
}
if (-not $compilationReplacementRejected) {
    throw "AgentTaskSpec treated different compilation governance state as idempotent."
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

# Offline production path: intent JSON -> merge -> BuildTaskSpec -> LoopEngine Validate/Freeze.
$rawProductionRequest = Decode-JsonString '"\u4f7f\u7528Python\u8ba1\u7b97\u6bcf\u4e2a\u533a\u57df\u7684\u5e73\u5747\u9500\u552e\u989d\uff0c\u5e76\u5199\u5165\u201cPython\u5e73\u5747\u503c0821\u201d"'
$computeClauseText = Decode-JsonString '"\u4f7f\u7528Python\u8ba1\u7b97\u6bcf\u4e2a\u533a\u57df\u7684\u5e73\u5747\u9500\u552e\u989d"'
$outputClauseText = Decode-JsonString '"\u5199\u5165\u201cPython\u5e73\u5747\u503c0821\u201d"'
$modelAssumption = Decode-JsonString '"\u53ea\u7edf\u8ba1\u9500\u552e\u989d>1000"'
$intentPayloadText = @'
{
  "intentType": "DATA_ANALYSIS",
  "confidence": 0.95,
  "description": "structured goal",
  "interactionMode": "execute",
  "requestedOutputs": ["worksheet"],
  "goalInterpretation": {
    "rawUserRequest": "MODEL_FORGED_RAW",
    "sourceClauses": [
      {"id":"model-compute","text":"\u4f7f\u7528Python\u8ba1\u7b97\u6bcf\u4e2a\u533a\u57df\u7684\u5e73\u5747\u9500\u552e\u989d","isExplicit":true,"requiredCapability":"PythonCompute"},
      {"id":"model-output","text":"\u5199\u5165\u201cPython\u5e73\u5747\u503c0821\u201d","isExplicit":true,"requiredCapability":""}
    ],
    "criteria": [
      {"id":"semantic-compute","statement":"\u4f7f\u7528Python\u8ba1\u7b97\u6bcf\u4e2a\u533a\u57df\u7684\u5e73\u5747\u9500\u552e\u989d","kind":"semantic","sourceClauseIds":["model-compute"],"required":true,"verificationCapability":"semantic","capabilityId":""},
      {"id":"semantic-output","statement":"\u5199\u5165\u201cPython\u5e73\u5747\u503c0821\u201d","kind":"semantic","sourceClauseIds":["model-output"],"required":true,"verificationCapability":"semantic","capabilityId":""},
      {"id":"python-required","statement":"Python must perform the calculation","kind":"capability","sourceClauseIds":["model-compute"],"required":true,"verificationCapability":"PythonCompute","capabilityId":"PythonCompute"},
      {"id":"compute-average","statement":"\u4f7f\u7528Python\u8ba1\u7b97\u6bcf\u4e2a\u533a\u57df\u7684\u5e73\u5747\u9500\u552e\u989d","kind":"compute","sourceClauseIds":["model-compute"],"required":true,"verificationCapability":"data","capabilityId":""},
      {"id":"write-result","statement":"\u5199\u5165\u201cPython\u5e73\u5747\u503c0821\u201d","kind":"state","sourceClauseIds":["model-output"],"required":true,"verificationCapability":"worksheet","capabilityId":""}
    ],
    "constraints": [],
    "requiredCapabilities": ["PythonCompute"],
    "unresolvedClauses": [],
    "assumptions": ["\u53ea\u7edf\u8ba1\u9500\u552e\u989d>1000"],
    "requiresClarification": false
  }
}
'@

$intentService = [ShareRibbon.IntentRecognitionService]::new("Excel")
$parseIntentMethod = [ShareRibbon.IntentRecognitionService].GetMethod("ParseLLMIntentResponse", $instanceNonPublic)
$mergeIntentMethod = [ShareRibbon.IntentRecognitionService].GetMethod(
    "MergeLlmIntentResult",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $parseIntentMethod -or $null -eq $mergeIntentMethod) {
    throw "Production intent goal parser/merge seam is missing."
}

# sourceStart must survive the real intent JSON parser. Repeated source text is accepted only
# when each semantic occurrence has an exact UTF-16 offset; omission or a wrong offset fails.
$duplicateIntentRaw = Decode-JsonString '"\u7532|\u7532"'
$duplicateIntentJson = @'
{
  "intentType":"DATA_ANALYSIS","confidence":0.95,"interactionMode":"execute","requestedOutputs":[],
  "goalInterpretation":{
    "sourceClauses":[
      {"id":"first","text":"\u7532","isExplicit":true,"sourceStart":0},
      {"id":"second","text":"\u7532","isExplicit":true,"sourceStart":2}
    ],
    "criteria":[
      {"id":"first-state","statement":"\u7532","kind":"state","sourceClauseIds":["first"],"required":true},
      {"id":"second-state","statement":"\u7532","kind":"state","sourceClauseIds":["second"],"required":true}
    ],
    "constraints":[],"requiredCapabilities":[],"unresolvedClauses":[],"assumptions":[],"requiresClarification":false
  }
}
'@
$duplicateParsed = [ShareRibbon.IntentResult]$parseIntentMethod.Invoke(
    $intentService,
    [object[]]@($duplicateIntentJson, [string]$duplicateIntentRaw))
$duplicateParsedCandidate = $duplicateParsed.GoalInterpretation.Candidate
if ($duplicateParsedCandidate.SourceClauses[0].SourceStart -ne 0 -or
    $duplicateParsedCandidate.SourceClauses[1].SourceStart -ne 2) {
    throw "ParseGoalInterpretation discarded sourceStart from the production JSON path."
}
$duplicateParsedCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($duplicateParsedCandidate)
$duplicateParsedValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($duplicateParsedCompilation)
if (-not $duplicateParsedValidation.Succeeded) {
    throw "Repeated source clauses with exact occurrence offsets failed validation: $($duplicateParsedValidation.Errors -join '; ')"
}

$missingOffsetJson = $duplicateIntentJson.Replace('"sourceStart":0', '"sourceStart":-1').Replace('"sourceStart":2', '"sourceStart":-1')
$missingOffsetParsed = [ShareRibbon.IntentResult]$parseIntentMethod.Invoke(
    $intentService,
    [object[]]@($missingOffsetJson, [string]$duplicateIntentRaw))
$missingOffsetCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($missingOffsetParsed.GoalInterpretation.Candidate)
if ([ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($missingOffsetCompilation).Succeeded) {
    throw "Repeated source text without disambiguating sourceStart did not fail closed."
}

$emojiIntentRaw = Decode-JsonString '"\ud83d\ude00|\u7532"'
$emojiIntentJson = $duplicateIntentJson.Replace('"first","text":"\u7532"', '"emoji","text":"\ud83d\ude00"').Replace('"first-state","statement":"\u7532"', '"emoji-state","statement":"\ud83d\ude00"').Replace('["first"]', '["emoji"]').Replace('"second","text":"\u7532","isExplicit":true,"sourceStart":2', '"second","text":"\u7532","isExplicit":true,"sourceStart":3')
$emojiParsed = [ShareRibbon.IntentResult]$parseIntentMethod.Invoke(
    $intentService,
    [object[]]@($emojiIntentJson, [string]$emojiIntentRaw))
$emojiCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($emojiParsed.GoalInterpretation.Candidate)
if (-not [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($emojiCompilation).Succeeded -or
    $emojiCompilation.Candidate.SourceClauses[1].SourceStart -ne 3) {
    throw "sourceStart was not interpreted as a zero-based UTF-16 offset."
}

$parsedIntent = [ShareRibbon.IntentResult]$parseIntentMethod.Invoke(
    $intentService,
    [object[]]@($intentPayloadText, [string]$rawProductionRequest))
$mergedIntent = [ShareRibbon.IntentResult]::new()
$mergeIntentMethod.Invoke($null, @($mergedIntent, $parsedIntent)) | Out-Null
if ($null -eq $mergedIntent.GoalInterpretation) {
    throw "Structured GoalInterpretation disappeared during the production intent merge."
}

$runtime = [ShareRibbon.Agent.AiNativeRuntime]::new([ShareRibbon.Agent.ToolRegistry]::new($null))
$buildTaskSpecMethod = [ShareRibbon.Agent.AiNativeRuntime].GetMethod("BuildTaskSpec", $instanceNonPublic)
$runtimeRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
$runtimeRequest.UserInput = $rawProductionRequest
$runtimeRequest.AppType = "Excel"
$emptyTools = [System.Collections.Generic.List[ShareRibbon.Agent.ToolDescriptor]]::new()
$emptySkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
$productionSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpecMethod.Invoke(
    $runtime,
    @($runtimeRequest, $mergedIntent, $emptyTools, $emptySkills, "Excel"))
$goalCompilationProperty = [ShareRibbon.Agent.AgentTaskSpec].GetProperty("GoalCompilation", $instanceNonPublic)
$attachedCompilation = [ShareRibbon.Agent.Goals.GoalCompilationResult]$goalCompilationProperty.GetValue($productionSpec)
if ($null -eq $attachedCompilation -or
    $attachedCompilation.Candidate.RawUserRequest -cne $rawProductionRequest -or
    ($attachedCompilation.Diagnostics -join " ") -notmatch "structured goal interpretation") {
    throw "BuildTaskSpec did not carry the structured interpretation with captured raw authority."
}
$attachedValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($attachedCompilation)
if (-not $attachedValidation.Succeeded) {
    throw "Structured production interpretation failed deterministic validation before freeze: $($attachedValidation.Errors -join '; '); raw=[$rawProductionRequest]; clauses=$(@($attachedCompilation.Candidate.SourceClauses | ForEach-Object { '[' + $_.Text + '] len=' + $_.Text.Length }) -join ' | ')"
}

$productionSession = [ShareRibbon.Agent.AgentSession]::new($rawProductionRequest, "Excel", "")
$productionSession.Spec = $productionSpec
$freezeMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "EstablishFrozenGoalContract",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
$freezeError = [string]$freezeMethod.Invoke($null, @($productionSession))
if (-not [string]::IsNullOrEmpty($freezeError) -or $null -eq $productionSpec.GoalContract) {
    throw "Production Goal Interpretation path did not validate/freeze: $freezeError"
}
$productionGoal = $productionSpec.GoalContract
$computeSource = @($productionGoal.SourceClauses | Where-Object { $_.Text -ceq $computeClauseText })
$outputSource = @($productionGoal.SourceClauses | Where-Object { $_.Text -ceq $outputClauseText })
$requiredAssumptionLeak = @($productionGoal.Criteria + $productionGoal.Constraints | Where-Object {
    $_.Required -and $_.Statement -ceq $modelAssumption
})
if ($productionGoal.RawUserRequest -cne $rawProductionRequest -or
    -not $productionGoal.RequiredCapabilities.Contains("PythonCompute") -or
    $computeSource.Count -ne 1 -or $outputSource.Count -ne 1 -or
    @($productionGoal.Criteria | Where-Object { $_.Kind -eq "compute" -and $_.SourceClauseIds.Contains($computeSource[0].Id) }).Count -ne 1 -or
    @($productionGoal.Criteria | Where-Object { $_.Kind -eq "state" -and $_.SourceClauseIds.Contains($outputSource[0].Id) }).Count -ne 1 -or
    $requiredAssumptionLeak.Count -ne 0) {
    throw "Frozen production goal lost raw authority, clause traceability, capability policy, or assumption governance. raw=$($productionGoal.RawUserRequest); caps=$($productionGoal.RequiredCapabilities -join ','); computeSources=$($computeSource.Count); outputSources=$($outputSource.Count); criteria=$(@($productionGoal.Criteria | ForEach-Object { $_.Kind + ':' + $_.Statement + ':' + ($_.SourceClauseIds -join ',') }) -join ' | '); assumptionLeaks=$($requiredAssumptionLeak.Count)"
}

# Invalid structured semantics may fall back to exact text, but the downgrade is observable and
# independently evidenced method constraints survive it.
$invalidStructuredCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$invalidStructuredCandidate.RawUserRequest = $pythonText
$invalidStructuredCandidate.SourceClauses.Add((New-SourceClause "invalid-python-source" $pythonText))
$invalidStructuredCandidate.SourceClauses.Add((New-SourceClause "fabricated-source" "only include sales above 1000"))
$invalidStructuredCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($invalidStructuredCandidate)
$invalidStructuredSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
Set-RawRequest $invalidStructuredSpec $pythonText
Set-GoalCompilation $invalidStructuredSpec $invalidStructuredCompilation
$invalidStructuredSession = [ShareRibbon.Agent.AgentSession]::new($pythonText, "Excel", "")
$invalidStructuredSession.Spec = $invalidStructuredSpec
$invalidStructuredError = [string]$freezeMethod.Invoke($null, @($invalidStructuredSession))
if (-not [string]::IsNullOrEmpty($invalidStructuredError) -or
    $null -eq $invalidStructuredSpec.GoalContract -or
    -not $invalidStructuredSpec.GoalContract.RequiredCapabilities.Contains("PythonCompute") -or
    [string]::IsNullOrWhiteSpace($invalidStructuredSpec.GoalInterpretationFallbackReason)) {
    throw "Exact-text fallback hid its provenance or dropped an explicit method constraint: $invalidStructuredError"
}
$fallbackPollution = @($invalidStructuredSpec.GoalContract.SourceClauses | Where-Object {
    $_.Id -eq "fabricated-source" -or $_.Text -ceq "only include sales above 1000"
})
$fallbackCriterionPollution = @($invalidStructuredSpec.GoalContract.Criteria | Where-Object {
    $_.Statement -ceq "only include sales above 1000"
})
if ($fallbackPollution.Count -ne 0 -or $fallbackCriterionPollution.Count -ne 0) {
    throw "Exact-text fallback retained fabricated model source or criterion semantics."
}

# A model that explicitly requires clarification cannot be downgraded to raw fallback and
# executed. The frozen authority remains absent until the user supplies the missing fact.
$clarificationCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$clarificationCandidate.RawUserRequest = $averageText
$clarificationCandidate.SourceClauses.Add((New-SourceClause "clarification-source" $averageText))
$clarificationCandidate.Criteria.Add((New-Criterion "clarification-criterion" $averageText "semantic" @("clarification-source")))
$clarificationCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile(
    $clarificationCandidate,
    [string[]]@("clarification-source"),
    [string[]]@(),
    [string[]]@(),
    $true)
$clarificationSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
Set-RawRequest $clarificationSpec $averageText
Set-GoalCompilation $clarificationSpec $clarificationCompilation
$clarificationSession = [ShareRibbon.Agent.AgentSession]::new($averageText, "Excel", "")
$clarificationSession.Spec = $clarificationSpec
$clarificationError = [string]$freezeMethod.Invoke($null, @($clarificationSession))
if ([string]::IsNullOrWhiteSpace($clarificationError) -or
    $clarificationError -notmatch "clarification" -or
    $null -ne $clarificationSpec.GoalContract) {
    throw "RequiresClarification was silently downgraded into an executable raw goal: $clarificationError"
}

# Cross-turn destination corrections preserve the complete transcript and deterministic method
# constraint without importing mutable legacy TaskSpec projections as semantic authority.
$correctionText = Decode-JsonString '"\u6211\u521a\u521a\u8bf4\u7684\u65b0\u5de5\u4f5c\u8868\u6539\u4e3a\u4e4b\u524d\u521b\u5efa\u7684\u9a8c\u8bc1\u6c47\u603b0821"'
$correctionIntent = [ShareRibbon.IntentResult]::new()
$correctionIntent.ResponseMode = "execute"
$correctionRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
$correctionRequest.UserInput = $correctionText
$correctionRequest.AppType = "Excel"
$correctionRequest.PreviousTaskSpec = $productionSpec
$correctionSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpecMethod.Invoke(
    $runtime,
    @($correctionRequest, $correctionIntent, $emptyTools, $emptySkills, "Excel"))
$expectedCorrectionRaw = $rawProductionRequest + "`r`n" + $correctionText
if ($correctionSpec.RawUserRequest -cne $expectedCorrectionRaw -or
    [string]::IsNullOrWhiteSpace($correctionSpec.GoalInterpretationFallbackReason)) {
    throw "Cross-turn correction lost exact transcript authority or hid raw-fallback provenance."
}
$correctionSession = [ShareRibbon.Agent.AgentSession]::new($correctionText, "Excel", "")
$correctionSession.Spec = $correctionSpec
$correctionFreezeError = [string]$freezeMethod.Invoke($null, @($correctionSession))
if (-not [string]::IsNullOrEmpty($correctionFreezeError) -or
    $null -eq $correctionSpec.GoalContract -or
    -not $correctionSpec.GoalContract.RequiredCapabilities.Contains("PythonCompute") -or
    $correctionSpec.GoalContract.RawUserRequest -cne $expectedCorrectionRaw) {
    throw "Cross-turn correction lost the explicit Python method or exact combined request: $correctionFreezeError"
}

Write-Host "PASS: immutable, traceable Goal Boundary contracts"
