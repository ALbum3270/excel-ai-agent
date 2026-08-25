param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$assemblyPath = Join-Path $repoRoot "ShareRibbon\bin\$Configuration\ShareRibbon.dll"
$toolsPath = Join-Path $repoRoot "ShareRibbon\Tools"

if (-not (Test-Path -LiteralPath $assemblyPath)) {
    throw "ShareRibbon assembly not found: $assemblyPath. Build the code projects first."
}

Add-Type -Path $assemblyPath

$registry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$registry.LoadFromDirectory($toolsPath)

$goalOutcomeProjectionType = [ShareRibbon.Agent.AgentSession].Assembly.GetType(
    "ShareRibbon.Agent.Goals.GoalOutcomeProjection",
    $true)
$sealOutcomeContractMethod = $goalOutcomeProjectionType.GetMethod(
    "Seal",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
$requiredHostCriterionIdsMethod = $goalOutcomeProjectionType.GetMethod(
    "RequiredHostCriterionIds",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $sealOutcomeContractMethod -or $null -eq $requiredHostCriterionIdsMethod) {
    throw "GoalOutcomeProjection.Seal was not found"
}
$captureRawGoalMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod(
    "CaptureRawUserRequest",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$setFrozenGoalMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod(
    "SetGoalContractOnce",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$setGoalCompilationMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod(
    "SetGoalCompilationOnce",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$recordGoalFallbackMethod = [ShareRibbon.Agent.AgentTaskSpec].GetMethod(
    "RecordGoalInterpretationFallback",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$goalExecutionAdmissionType = [ShareRibbon.Agent.AgentSession].Assembly.GetType(
    "ShareRibbon.Agent.Goals.GoalExecutionAdmission",
    $true)
$validateGoalExecutionAdmissionMethod = $goalExecutionAdmissionType.GetMethod(
    "Validate",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $captureRawGoalMethod -or $null -eq $setFrozenGoalMethod -or
    $null -eq $setGoalCompilationMethod -or $null -eq $recordGoalFallbackMethod -or
    $null -eq $validateGoalExecutionAdmissionMethod) {
    throw "AgentTaskSpec Goal boundary methods were not found"
}

function New-TestGoalContract {
    param(
        [string]$RawRequest,
        [string[]]$ClauseTexts,
        [string[]]$CriterionIds,
        [string[]]$CriterionKinds
    )

    if ($ClauseTexts.Count -ne $CriterionIds.Count -or
        $ClauseTexts.Count -ne $CriterionKinds.Count) {
        throw "Goal test vectors must have matching lengths"
    }

    $candidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
    $candidate.RawUserRequest = $RawRequest
    $searchStart = 0
    for ($index = 0; $index -lt $ClauseTexts.Count; $index++) {
        $text = $ClauseTexts[$index]
        $sourceStart = $RawRequest.IndexOf(
            $text,
            $searchStart,
            [System.StringComparison]::Ordinal)
        if ($sourceStart -lt 0) {
            throw "Goal test clause is not an exact raw-request span: $text"
        }
        $clauseId = "source-$($index + 1)"
        $clause = [ShareRibbon.Agent.Goals.CandidateGoalSourceClause]::new()
        $clause.Id = $clauseId
        $clause.Text = $text
        $clause.IsExplicit = $true
        $clause.SourceStart = $sourceStart
        $candidate.SourceClauses.Add($clause)

        $criterion = [ShareRibbon.Agent.Goals.CandidateGoalCriterion]::new()
        $criterion.Id = $CriterionIds[$index]
        $criterion.Statement = $text
        $criterion.Kind = $CriterionKinds[$index]
        $criterion.Required = $true
        $criterion.SourceClauseIds.Add($clauseId)
        $candidate.Criteria.Add($criterion)
        $searchStart = $sourceStart + $text.Length
    }

    $compilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($candidate)
    $validation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($compilation)
    if (-not $validation.Succeeded) {
        throw "Test GoalContract failed validation: $($validation.Errors -join '; ')"
    }
    return [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
        $compilation,
        $validation)
}

function New-TestGoalSession {
    param([ShareRibbon.Agent.Goals.GoalContract]$GoalContract)

    $session = [ShareRibbon.Agent.AgentSession]::new(
        $GoalContract.RawUserRequest,
        "Excel",
        "")
    $session.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
    $session.Spec.MutationPolicy = "allow"
    $session.Spec.Goal = $GoalContract.RawUserRequest
    $null = $captureRawGoalMethod.Invoke(
        $session.Spec,
        @($GoalContract.RawUserRequest))
    $null = $setFrozenGoalMethod.Invoke($session.Spec, @($GoalContract))
    return $session
}

# A raw-preserving interpretation fallback must remain executable. It preserves the exact
# user request; per-action SafetyGate and final host verification still guard mutations.
$fallbackRawRequest = "format the current worksheet"
$fallbackCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($fallbackRawRequest)
$fallbackValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($fallbackCompilation)
if (-not $fallbackValidation.Succeeded) {
    throw "Raw-preserving Goal fixture failed validation: $($fallbackValidation.Errors -join '; ')"
}
$fallbackGoal = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $fallbackCompilation,
    $fallbackValidation)
$fallbackSession = New-TestGoalSession $fallbackGoal
$null = $setGoalCompilationMethod.Invoke($fallbackSession.Spec, @($fallbackCompilation))
$null = $recordGoalFallbackMethod.Invoke(
    $fallbackSession.Spec,
    @("structured interpretation unavailable"))
$fallbackSession.Spec.MutationPolicy = "allow"
$fallbackMutationAdmissionError = [string]$validateGoalExecutionAdmissionMethod.Invoke(
    $null,
    @($fallbackSession.Spec))
if ([string]::IsNullOrWhiteSpace($fallbackMutationAdmissionError)) {
    # expected
} else {
    throw "Exact-text interpretation fallback was blocked before adaptive execution: $fallbackMutationAdmissionError"
}
$fallbackSession.Spec.MutationPolicy = "read_only"
$fallbackReadAdmissionError = [string]$validateGoalExecutionAdmissionMethod.Invoke(
    $null,
    @($fallbackSession.Spec))
if (-not [string]::IsNullOrWhiteSpace($fallbackReadAdmissionError)) {
    throw "Exact-text interpretation fallback blocked a read-only task: $fallbackReadAdmissionError"
}

# Required constraints are preserved as Goal-bound host obligations. They may reach the
# adaptive loop, but completion cannot omit their ids from the verification projection.
$constraintRawRequest = "write result without changing source"
$constraintCandidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
$constraintCandidate.RawUserRequest = $constraintRawRequest
$constraintClause = [ShareRibbon.Agent.Goals.CandidateGoalSourceClause]::new()
$constraintClause.Id = "source-constraint"
$constraintClause.Text = $constraintRawRequest
$constraintClause.IsExplicit = $true
$constraintClause.SourceStart = 0
$constraintCandidate.SourceClauses.Add($constraintClause)
$constraintCriterion = [ShareRibbon.Agent.Goals.CandidateGoalCriterion]::new()
$constraintCriterion.Id = "goal-constrained-write"
$constraintCriterion.Statement = $constraintRawRequest
$constraintCriterion.Kind = "semantic"
$constraintCriterion.Required = $true
$constraintCriterion.SourceClauseIds.Add("source-constraint")
$constraintCandidate.Criteria.Add($constraintCriterion)
$requiredConstraint = [ShareRibbon.Agent.Goals.CandidateGoalConstraint]::new()
$requiredConstraint.Id = "constraint-source"
$requiredConstraint.Statement = $constraintRawRequest
$requiredConstraint.Required = $true
$requiredConstraint.SourceClauseIds.Add("source-constraint")
$constraintCandidate.Constraints.Add($requiredConstraint)
$constraintCompilation = [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($constraintCandidate)
$constraintValidation = [ShareRibbon.Agent.Goals.GoalCoverageValidator]::Validate($constraintCompilation)
if (-not $constraintValidation.Succeeded) {
    throw "Required-constraint Goal fixture failed validation: $($constraintValidation.Errors -join '; ')"
}
$constraintGoal = [ShareRibbon.Agent.Goals.GoalContractFreezer]::Freeze(
    $constraintCompilation,
    $constraintValidation)
$constraintSession = New-TestGoalSession $constraintGoal
$null = $setGoalCompilationMethod.Invoke($constraintSession.Spec, @($constraintCompilation))
$constraintAdmissionError = [string]$validateGoalExecutionAdmissionMethod.Invoke(
    $null,
    @($constraintSession.Spec))
if (-not [string]::IsNullOrWhiteSpace($constraintAdmissionError)) {
    throw "A required Goal constraint was blocked before it could be verified: $constraintAdmissionError"
}
$requiredConstraintProjectionIds = $requiredHostCriterionIdsMethod.Invoke($null, @($constraintSession.Spec))
if (-not $requiredConstraintProjectionIds.Contains("constraint-source")) {
    throw "A required Goal constraint disappeared from host-verifiable obligations"
}

# Read-only is a semantic safety invariant. A stale or overly conservative risk label
# must not turn a non-mutating lookup into an approval prompt.
$readTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$readTool.Id = "TestRead"
$readTool.AppType = "excel"
$readTool.RiskLevel = "risky"
$readTool.AccessMode = "read"
$readDecision = ([ShareRibbon.Agent.Execution.SafetyGate]::new()).Evaluate(
    $readTool,
    [Newtonsoft.Json.Linq.JObject]::new())
if ($readDecision.Action -ne [ShareRibbon.Agent.Execution.SafetyAction]::Allow -or
    $readDecision.RiskLevel -ne "safe") {
    throw "Read-only tools are not approval-free by default"
}

# TaskSpec.RequiredTools describes the anticipated plan. It must not hide any supporting
# tool that the selected Skill explicitly authorizes. Assert this as a set invariant over
# the real Excel registry/Skill rather than checking one prompt or one named tool.
$skillDefinition = [ShareRibbon.SkillsDirectoryService]::GetAllSkills($true) |
    Where-Object Name -eq "Excel Table Agent" |
    Select-Object -First 1
if ($null -eq $skillDefinition) {
    throw "Excel Table Agent skill was not loaded"
}

$skill = [ShareRibbon.Agent.AgentSkill]::new()
foreach ($toolId in $skillDefinition.AllowedTools) {
    $skill.RequiredTools.Add($toolId)
}

$availableExcelToolIds = @($registry.GetAvailableTools("Excel") | ForEach-Object Id)
$missingFromSkill = @($availableExcelToolIds | Where-Object { -not $skill.RequiredTools.Contains($_) })
if ($missingFromSkill.Count -gt 0) {
    throw "Registered Excel tools are missing from Excel Table Agent: $($missingFromSkill -join ', ')"
}

# A specialised Skill is useful retrieval context, but it must not become a smaller
# capability universe than the semantic task contract. Select a narrow real Skill by
# set properties (not by prompt/name) and require promotion to a covering baseline.
$narrowSkill = [ShareRibbon.SkillsDirectoryService]::GetAllSkills($true) |
    Where-Object {
        $_.AllowedTools -contains "DataAnalysis" -and
        $_.AllowedTools -notcontains "ReadRange"
    } |
    Select-Object -First 1
if ($null -eq $narrowSkill) {
    throw "No narrow Excel analysis Skill available for capability-resolution smoke"
}
$selectedSkillList = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
$selectedSkillList.Add($narrowSkill)
$pipelineTools = [string[]]@("ReadRange", "PythonCompute", "CreateSheet", "WriteData")
$resolvedSkills = [ShareRibbon.Agent.SkillCapabilityResolver]::ResolvePrimarySkill(
    $selectedSkillList,
    $pipelineTools,
    "Excel")
$uncoveredPipelineTools = @($pipelineTools | Where-Object {
    $resolvedSkills.Count -eq 0 -or $resolvedSkills[0].AllowedTools -notcontains $_
})
if ($uncoveredPipelineTools.Count -gt 0) {
    throw "Skill capability resolver still hides task dependencies: $($uncoveredPipelineTools -join ', ')"
}

# A specialised Skill is retrieval/advice, not the authorization universe for built-in
# host tools. BuildTaskSpec must not copy its whole allowed-tools list into semantic task
# requirements, and execution must retain the application's registered host capabilities.
$syntheticNarrowSkill = [ShareRibbon.SkillFileDefinition]::new()
$syntheticNarrowSkill.Name = "Synthetic analysis advisor"
$syntheticNarrowSkill.Application = "excel"
$syntheticNarrowSkill.AllowedTools = [System.Collections.Generic.List[string]]::new()
$syntheticNarrowSkill.AllowedTools.Add("DataAnalysis")
$syntheticNarrowSkill.AllowedTools.Add("AutoFit")
$syntheticSkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
$syntheticSkills.Add($syntheticNarrowSkill)
$syntheticRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
$syntheticRequest.UserInput = "aggregate the observed table and place the result in a destination worksheet"
$syntheticRequest.AppType = "Excel"
$syntheticIntent = [ShareRibbon.IntentResult]::new()
$syntheticIntent.ResponseMode = "execute"
$syntheticIntent.Confidence = 0.95
$syntheticIntent.UserFriendlyDescription = "aggregate observed workbook data into a destination"
$syntheticIntent.GoalInterpretation = [ShareRibbon.Agent.Goals.GoalInterpretationPayload]::new()
$syntheticIntent.GoalInterpretation.Candidate =
    [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($syntheticRequest.UserInput).Candidate
$runtime = [ShareRibbon.Agent.AiNativeRuntime]::new($registry)
$buildTaskSpecMethod = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
    "BuildTaskSpec",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $buildTaskSpecMethod) {
    throw "AiNativeRuntime.BuildTaskSpec was not found"
}
$syntheticSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpecMethod.Invoke(
    $runtime,
    @($syntheticRequest, $syntheticIntent, $registry.GetAvailableTools("Excel"), $syntheticSkills, "Excel"))
if ($syntheticSpec.RequiredTools.Contains("DataAnalysis") -or
    $syntheticSpec.RequiredTools.Contains("AutoFit")) {
    throw "A specialised Skill allowed-tools list was copied into TaskSpec.RequiredTools"
}

# Workbook-dependent questions remain executable read-only tasks even when a provider labels
# them GENERAL_QUERY/answer. The structured Goal payload says the answer depends on Office
# state; plain chat keeps that payload absent and must remain on the conversational path.
$readOnlyContext = [ShareRibbon.Agent.Context.OfficeContext]::new()
$readOnlyContext.AppType = "Excel"
$readOnlyContext.HostData["tables"] = [Newtonsoft.Json.Linq.JArray]::Parse(
    '[{"sheet":"SalesData","address":"A1:D25"}]')
$readOnlyRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
$readOnlyRequest.UserInput = "How many records does each region have? Answer only; do not write to the workbook."
$readOnlyRequest.AppType = "Excel"
$readOnlyRequest.OfficeContext = $readOnlyContext
$readOnlyIntent = [ShareRibbon.IntentResult]::new()
$readOnlyIntent.ResponseMode = "answer"
$readOnlyIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::GENERAL_QUERY
$readOnlyIntent.IntentType = [ShareRibbon.ExcelIntentType]::GENERAL_QUERY
$readOnlyIntent.GoalInterpretation = [ShareRibbon.Agent.Goals.GoalInterpretationPayload]::new()
$readOnlyIntent.GoalInterpretation.Candidate =
    [ShareRibbon.Agent.Goals.GoalCompiler]::Compile($readOnlyRequest.UserInput).Candidate
$readOnlySpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpecMethod.Invoke(
    $runtime,
    @($readOnlyRequest, $readOnlyIntent, $registry.GetAvailableTools("Excel"), $syntheticSkills, "Excel"))
if ($readOnlySpec.MutationPolicy -ne "read_only" -or
    -not $readOnlySpec.RequiredTools.Contains("ReadRange")) {
    throw "A structured workbook question classified GENERAL_QUERY was still routed as plain chat"
}
$plainChatIntent = [ShareRibbon.IntentResult]::new()
$plainChatIntent.ResponseMode = "answer"
$plainChatRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
$plainChatRequest.UserInput = "What is a pivot table?"
$plainChatRequest.AppType = "Excel"
$plainChatRequest.OfficeContext = $readOnlyContext
$plainChatSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpecMethod.Invoke(
    $runtime,
    @($plainChatRequest, $plainChatIntent, $registry.GetAvailableTools("Excel"), $syntheticSkills, "Excel"))
if ($plainChatSpec.RequiredTools.Contains("ReadRange") -or
    $plainChatSpec.RequiresHostExecution -or
    $plainChatSpec.MutationPolicy -ne "read_only") {
    throw "Pure Excel help text did not remain a no-host, no-mutation turn inside the unified Agent"
}

$syntheticSession = [ShareRibbon.Agent.AgentSession]::new($syntheticRequest.UserInput, "Excel", "")
$syntheticSession.Spec = $syntheticSpec
$syntheticAgentSkill = [ShareRibbon.Agent.AgentSkill]::new()
$syntheticAgentSkill.Name = $syntheticNarrowSkill.Name
foreach ($toolId in $syntheticNarrowSkill.AllowedTools) {
    $syntheticAgentSkill.RequiredTools.Add($toolId)
}
$selectedPrivateTool = "skill_script.Synthetic analysis advisor.local.py"
$selectedMcpTool = "mcp.selected-external-tool"
$syntheticAgentSkill.RequiredTools.Add($selectedPrivateTool)
$syntheticAgentSkill.RequiredTools.Add($selectedMcpTool)
$syntheticContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession(
    $syntheticSession,
    $syntheticAgentSkill)
$hiddenHostTools = @($registry.GetAvailableTools("Excel") | Where-Object {
    -not $syntheticContext.IsToolAllowed($_)
})
if ($hiddenHostTools.Count -gt 0) {
    throw "A specialised advisor still hides registered Excel host tools: $($hiddenHostTools.Id -join ', ')"
}

# The host baseline must come from trusted registration provenance, not a tool-id prefix or
# the mere fact that a descriptor is registered.  Internal writes and arbitrary custom tools
# stay fail-closed unless the run explicitly grants them.
$unselectedMemoryPromotion = $registry.GetTool("memory.promote")
if ($null -eq $unselectedMemoryPromotion -or
    $syntheticContext.IsToolAllowed($unselectedMemoryPromotion)) {
    throw "An unselected Agent-internal write tool leaked into the Office host baseline"
}
$unselectedDirectCall = [ShareRibbon.Agent.ToolCall]::new()
$unselectedDirectCall.ToolId = "memory.promote"
$unselectedDirectCall.Parameters = [Newtonsoft.Json.Linq.JObject]::new()
$normalizationMessage = ""
if ($registry.TryNormalizeToolCall(
        "Excel",
        $unselectedDirectCall,
        $syntheticContext,
        [ref]$normalizationMessage)) {
    throw "Direct-id normalization exposed a tool denied by descriptor authorization"
}
$customRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$unselectedCustomTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$unselectedCustomTool.Id = "UnselectedCustomWrite"
$unselectedCustomTool.Name = "Unselected custom write"
$unselectedCustomTool.AppType = "excel"
$unselectedCustomTool.RiskLevel = "safe"
$unselectedCustomTool.AccessMode = "write"
$customRegistry.RegisterTool($unselectedCustomTool)
$script:unselectedCustomExecutions = 0
$customRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:unselectedCustomExecutions += 1
    return [ShareRibbon.Agent.ToolResult]::Succeed("UnselectedCustomWrite", "unexpected")
}
$customResult = $customRegistry.ExecuteToolAsync(
    $syntheticContext,
    "UnselectedCustomWrite",
    [Newtonsoft.Json.Linq.JObject]::new()).GetAwaiter().GetResult()
if ($customResult.Success -or
    $customResult.ErrorCode -ne [ShareRibbon.ExceptionClassifier]::CodeToolNotAllowed -or
    $script:unselectedCustomExecutions -ne 0) {
    throw "A public/custom registration was treated as an implicitly trusted host tool"
}

# Descriptor provenance and owner are registry facts.  A Skill cannot authorize another
# Skill's script by copying its id into allowed-tools, and string-only checks cannot bypass
# that descriptor-level boundary.
$registrationMethod = [ShareRibbon.Agent.ToolDescriptor].GetMethod(
    "AssignRegistrationTrust",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $registrationMethod) {
    throw "Tool descriptors do not expose a registry-owned provenance/owner seam"
}
$hostCreateSheet = $registry.GetTool("CreateSheet")
if ($null -eq $hostCreateSheet -or
    $hostCreateSheet.RegistrationProvenance -ne [ShareRibbon.Agent.ToolRegistrationProvenance]::HostManifest -or
    $unselectedMemoryPromotion.RegistrationProvenance -ne [ShareRibbon.Agent.ToolRegistrationProvenance]::AgentInternal -or
    $customRegistry.GetTool("UnselectedCustomWrite").RegistrationProvenance -ne [ShareRibbon.Agent.ToolRegistrationProvenance]::Custom) {
    throw "Registry paths did not assign trustworthy host/internal/custom provenance"
}
$customCollision = [ShareRibbon.Agent.ToolDescriptor]::new()
$customCollision.Id = "UnselectedCustomWrite"
$customCollision.Name = "custom collision"
$customCollision.AppType = "excel"
$customCollision.AccessMode = "read"
$customCollision.RiskLevel = "risky"
$customRegistry.RegisterTool($customCollision)
$customAfterCollision = $customRegistry.GetTool("UnselectedCustomWrite")
if (-not [object]::ReferenceEquals($customAfterCollision, $unselectedCustomTool) -or
    $customAfterCollision.AccessMode -ne "write" -or
    $customAfterCollision.RiskLevel -ne "safe") {
    throw "Two unowned custom registrations were merged across an ambiguous owner boundary"
}
$hostAccessMode = $hostCreateSheet.AccessMode
$hostRiskLevel = $hostCreateSheet.RiskLevel
$hostOwnerId = $hostCreateSheet.RegistrationOwnerId
$collision = [ShareRibbon.Agent.ToolDescriptor]::new()
$collision.Id = "CreateSheet"
$collision.Name = "collision"
$collision.AppType = "excel"
$collision.AccessMode = "read"
$collision.RiskLevel = "risky"
$registry.RegisterTool($collision)
$createSheetAfterCollision = $registry.GetTool("CreateSheet")
if (-not [object]::ReferenceEquals($createSheetAfterCollision, $hostCreateSheet) -or
    $createSheetAfterCollision.RegistrationProvenance -ne [ShareRibbon.Agent.ToolRegistrationProvenance]::HostManifest -or
    $createSheetAfterCollision.RegistrationOwnerId -ne $hostOwnerId -or
    $createSheetAfterCollision.AccessMode -ne $hostAccessMode -or
    $createSheetAfterCollision.RiskLevel -ne $hostRiskLevel) {
    throw "A custom registration collision changed a trusted host manifest descriptor"
}

$ownedScript = [ShareRibbon.Agent.ToolDescriptor]::new()
$ownedScript.Id = $selectedPrivateTool
$ownedScript.Name = "owned script"
$ownedScript.AppType = "common"
$ownedScript.AccessMode = "compute"
$null = $registrationMethod.Invoke(
    $ownedScript,
    @([ShareRibbon.Agent.ToolRegistrationProvenance]::SkillScript, $syntheticAgentSkill.Name))
if (-not $syntheticContext.IsToolAllowed($ownedScript)) {
    throw "The selected Skill could not use its own explicitly granted private script"
}
$ownerlessScript = [ShareRibbon.Agent.ToolDescriptor]::new()
$ownerlessScript.Id = "skill_script.ownerless.local.py"
$ownerlessScript.Name = "ownerless script"
$ownerlessScript.AppType = "common"
$ownerlessScript.AccessMode = "compute"
$syntheticAgentSkill.RequiredTools.Add($ownerlessScript.Id)
$null = $registrationMethod.Invoke(
    $ownerlessScript,
    @([ShareRibbon.Agent.ToolRegistrationProvenance]::SkillScript, ""))
$ownerContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession(
    $syntheticSession,
    $syntheticAgentSkill)
if ($ownerContext.IsToolAllowed($ownerlessScript)) {
    throw "An ownerless Skill script did not fail closed"
}
$foreignScript = [ShareRibbon.Agent.ToolDescriptor]::new()
$foreignScript.Id = "skill_script.Other advisor.local.py"
$foreignScript.Name = "foreign script"
$foreignScript.AppType = "common"
$foreignScript.AccessMode = "compute"
$null = $registrationMethod.Invoke(
    $foreignScript,
    @([ShareRibbon.Agent.ToolRegistrationProvenance]::SkillScript, "Other advisor"))
$syntheticAgentSkill.RequiredTools.Add($foreignScript.Id)
$ownerContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession(
    $syntheticSession,
    $syntheticAgentSkill)
if ($ownerContext.IsToolAllowed($foreignScript) -or
    $ownerContext.IsToolAllowed($foreignScript.Id)) {
    throw "A Skill authorized a foreign private script by id membership alone"
}

$selectedMcp = [ShareRibbon.Agent.ToolDescriptor]::new()
$selectedMcp.Id = $selectedMcpTool
$selectedMcp.Name = "selected MCP"
$selectedMcp.AppType = "common"
$selectedMcp.AccessMode = "read"
$null = $registrationMethod.Invoke(
    $selectedMcp,
    @([ShareRibbon.Agent.ToolRegistrationProvenance]::Mcp, "test-mcp"))
$unselectedMcp = [ShareRibbon.Agent.ToolDescriptor]::new()
$unselectedMcp.Id = "mcp.unselected-external-tool"
$unselectedMcp.Name = "unselected MCP"
$unselectedMcp.AppType = "common"
$unselectedMcp.AccessMode = "read"
$null = $registrationMethod.Invoke(
    $unselectedMcp,
    @([ShareRibbon.Agent.ToolRegistrationProvenance]::Mcp, "test-mcp"))
if (-not $ownerContext.IsToolAllowed($selectedMcp) -or
    $ownerContext.IsToolAllowed($unselectedMcp)) {
    throw "MCP visibility is not controlled by an explicit run grant"
}

# Planning receives only the tools needed by the semantic contract, while execution keeps
# the complete Skill authorization checked below. This reduces prompt size without hiding
# supporting capabilities from execution or repair.
$planningSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$planningSpec.RequiredTools.Add("CreateSheet")
$planningView = @([ShareRibbon.Agent.AgentPlanningScope]::SelectTools(
    $registry.GetAvailableTools("Excel"),
    $planningSpec))
if ($planningView.Count -ne 1 -or $planningView[0].Id -ne "CreateSheet") {
    throw "Planner tool view is not scoped by the semantic task contract"
}

# An engine name inside an object identifier (for example a worksheet name) is data, not
# an instruction to invoke that engine. Invocation requires a verb/operation relationship.
$pythonClassifier = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
    "IsExplicitPythonComputeRequest",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $pythonClassifier -or
    -not [bool]$pythonClassifier.Invoke($null, @("Run Python to calculate regional averages")) -or
    [bool]$pythonClassifier.Invoke($null, @("Create a worksheet named Python Average 0821"))) {
    throw "Python computation classification still treats identifier text as an invocation"
}

# A plan is explanatory guidance, not an executable workflow contract. Data dependency
# order is chosen from runtime observations, so even an implausible predicted order must
# not become a pre-execution gate.
$orderedSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$orderedSpec.RequiredCapabilities.Add("PythonCompute")
$wrongOrderPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
foreach ($id in @("PythonCompute", "ReadRange", "WriteData")) {
    $wrongOrderPlan.Steps.Add([ShareRibbon.Agent.PlanStep]@{
        Code = "{`"command`":`"$id`",`"params`":{}}"
        Language = "json"
    })
}
if (-not [string]::IsNullOrWhiteSpace(
    [ShareRibbon.Agent.AgentExecutionContract]::ValidatePlan($orderedSpec, $wrongOrderPlan))) {
    throw "Soft plan was incorrectly treated as an executable dependency contract"
}

# Outcome verification is a typed host-evidence contract.  These helpers build proof
# ledgers directly so the regression cases cannot accidentally pass through prompt text,
# model wording, or a catch inside ToolRegistry.
function New-TestOutcomeRequirement {
    param(
        [string]$Id,
        [string]$TargetRef,
        [string]$EffectType,
        [string]$PropertyName = "",
        [Newtonsoft.Json.Linq.JToken]$ExpectedValue = $null,
        [string]$DerivedFromCapability = "",
        [string]$Operator = "equals"
    )

    $requirement = [ShareRibbon.Agent.OutcomeRequirement]::new()
    $requirement.Id = $Id
    $requirement.AppType = "Excel"
    $requirement.TargetRef = $TargetRef
    $requirement.EffectType = $EffectType
    $requirement.PropertyName = $PropertyName
    $requirement.Operator = $Operator
    $requirement.ExpectedValue = $ExpectedValue
    $requirement.DerivedFromCapability = $DerivedFromCapability
    $requirement.Required = $true
    return $requirement
}

function New-TestOutcomeSession {
    param(
        [string]$Goal,
        [object[]]$Requirements
    )

    $session = [ShareRibbon.Agent.AgentSession]::new($Goal, "Excel", "")
    $session.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
    $session.Spec.MutationPolicy = "allow"
    $session.Spec.Goal = $Goal
    $contract = [ShareRibbon.Agent.OutcomeContract]::new()
    foreach ($requirement in $Requirements) {
        $contract.Requirements.Add($requirement)
    }
    $session.Spec.OutcomeContract = $contract
    $null = $sealOutcomeContractMethod.Invoke($null, @($session.Spec, $contract))
    return $session
}

function Seal-TestOutcomeContract {
    param([ShareRibbon.Agent.AgentSession]$Session)

    $null = $sealOutcomeContractMethod.Invoke(
        $null,
        @($Session.Spec, $Session.Spec.OutcomeContract))
}

function New-TestOutcomeEvidence {
    param(
        [string]$EvidenceId,
        [string]$IterationEvidenceId,
        [string]$TargetRef,
        [string]$EffectType,
        [string]$PropertyName = "",
        [Newtonsoft.Json.Linq.JToken]$Expected = $null,
        [string]$SourceToolId = "TestTool",
        [string[]]$DependsOn = @(),
        [bool]$Satisfied = $true,
        [bool]$InvalidatesPrior = $false,
        [long]$WorldRevision = 0
    )

    $record = [ShareRibbon.Agent.OutcomeEvidenceRecord]::new()
    $record.EvidenceId = $EvidenceId
    $record.IterationEvidenceId = $IterationEvidenceId
    $record.TargetRef = $TargetRef
    $record.EffectType = $EffectType
    $record.PropertyName = $PropertyName
    $record.Expected = $Expected
    if ($null -eq $Expected) {
        $record.Actual = $null
    } else {
        $record.Actual = $Expected.DeepClone()
    }
    $record.Satisfied = $Satisfied
    $record.InvalidatesPrior = $InvalidatesPrior
    $record.SourceToolId = $SourceToolId
    $record.WorldRevision = $WorldRevision
    foreach ($dependency in $DependsOn) {
        $record.DerivedFromEvidenceIds.Add($dependency)
    }
    return $record
}

function Add-TestOutcomeIteration {
    param(
        [ShareRibbon.Agent.AgentSession]$Session,
        [string]$IterationEvidenceId,
        [string]$ToolId,
        [object[]]$EvidenceRecords = @(),
        [string[]]$DependsOn = @(),
        [bool]$IterationSucceeded = $true
    )

    $iteration = [ShareRibbon.Agent.ReActIteration]::new()
    $iteration.EvidenceId = $IterationEvidenceId
    $iteration.Action = [ShareRibbon.Agent.ToolCall]::new()
    $iteration.Action.ToolId = $ToolId
    $iteration.Explanation = [ShareRibbon.Agent.ExecutionExplanation]::new()
    $iteration.Explanation.Success = $IterationSucceeded
    foreach ($dependency in $DependsOn) {
        $iteration.DependsOnEvidenceIds.Add($dependency)
    }
    foreach ($record in $EvidenceRecords) {
        $iteration.ContractEvidence.Add($record)
    }
    $Session.Iterations.Add($iteration)
}

function Test-OutcomeContract {
    param(
        [ShareRibbon.Agent.AgentSession]$Session,
        [string[]]$EvidenceClaims
    )

    $claims = [System.Collections.Generic.List[string]]::new()
    foreach ($claim in $EvidenceClaims) {
        $claims.Add($claim)
    }
    return [ShareRibbon.Agent.AgentGoalVerifier]::Validate($Session, $null, $claims, 0)
}

$green = [Newtonsoft.Json.Linq.JValue]::CreateString("#90EE90")

# Excel completion fails closed without a sealed verification projection. A successful local
# mutation can never become the implicit definition of the user's whole goal.
$noContractSession = [ShareRibbon.Agent.AgentSession]::new("format sales", "Excel", "")
$noContractSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$noContractSession.Spec.Goal = "format sales"
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $noContractSession @()))) {
    throw "An Excel task without a sealed verification projection was accepted"
}

# A contract range without a worksheet is not stable evidence. It must not be satisfied
# by an identically addressed range on whichever sheet happened to execute.
$unqualifiedRequirement = New-TestOutcomeRequirement `
    -Id "unqualified-range" `
    -TargetRef "A1:D25" `
    -EffectType "read_coverage" `
    -Operator "covers"
$wrongSheetSession = New-TestOutcomeSession "read a specific table" @($unqualifiedRequirement)
$wrongSheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:WrongSheet!A1:D25" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange"
Add-TestOutcomeIteration $wrongSheetSession "obs-1" "ReadRange" @($wrongSheetEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $wrongSheetSession @("obs-1/e1")))) {
    throw "An unqualified contract range was satisfied by evidence from an arbitrary worksheet"
}

$formatRequirement = New-TestOutcomeRequirement `
    -Id "format-sales" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -ExpectedValue $green

# A mutation of only part of a contracted range must not prove the whole requested range.
$partialRangeSession = New-TestOutcomeSession "format complete sales range" @($formatRequirement)
$partialEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D5" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange"
Add-TestOutcomeIteration $partialRangeSession "obs-1" "FormatRange" @($partialEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $partialRangeSession @("obs-1/e1")))) {
    throw "Partial range evidence was accepted for a whole-range outcome contract"
}

# Correct target plus the wrong property is still the wrong end state.
$wrongPropertySession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$wrongPropertyEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fontColor" `
    -Expected $green `
    -SourceToolId "FormatRange"
Add-TestOutcomeIteration $wrongPropertySession "obs-1" "FormatRange" @($wrongPropertyEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $wrongPropertySession @("obs-1/e1")))) {
    throw "Evidence for the wrong property was accepted by the outcome contract"
}

# A model may cite only evidence IDs present in the successful typed ledger.
$fakeEvidenceSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$fullFormatEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange"
Add-TestOutcomeIteration $fakeEvidenceSession "obs-1" "FormatRange" @($fullFormatEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $fakeEvidenceSession @("obs-404/e1")))) {
    throw "A fabricated evidence ID was accepted by the outcome contract"
}
$completeEvidenceError = Test-OutcomeContract $fakeEvidenceSession @("obs-1/e1")
if (-not [string]::IsNullOrWhiteSpace($completeEvidenceError)) {
    throw "A complete matching host-evidence record was rejected: $completeEvidenceError actual=$($fullFormatEvidence.Actual) property=$($fullFormatEvidence.PropertyName) expected=$($formatRequirement.ExpectedValue)"
}

# Tool request parameters are not host evidence. Even if the request asked for the exact
# contracted value and the adapter marked the operation satisfied, a conflicting observed
# actual value must fail the completion gate.
$actualMismatchSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$actualMismatchEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange"
$actualMismatchEvidence.Actual = [Newtonsoft.Json.Linq.JValue]::CreateString("#FF0000")
Add-TestOutcomeIteration $actualMismatchSession "obs-1" "FormatRange" @($actualMismatchEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $actualMismatchSession @("obs-1/e1")))) {
    throw "Requested parameters were accepted even though host-observed actual state disagreed"
}

# Historical evidence is not the final state. A later overlapping write to the same property
# supersedes an earlier match, even when the model cites only the stale evidence ID.
$staleStateSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$oldGreenEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange" `
    -WorldRevision 1
$newRedEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-2/e1" `
    -IterationEvidenceId "obs-2" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected ([Newtonsoft.Json.Linq.JValue]::CreateString("#FF0000")) `
    -SourceToolId "FormatRange" `
    -WorldRevision 2
Add-TestOutcomeIteration $staleStateSession "obs-1" "FormatRange" @($oldGreenEvidence)
Add-TestOutcomeIteration $staleStateSession "obs-2" "FormatRange" @($newRedEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $staleStateSession @("obs-1/e1")))) {
    throw "Stale matching evidence was accepted after a later overlapping state change"
}

# `changed` is audit metadata, not postcondition proof. A host rejection must stay rejected,
# while an idempotent no-change operation with an explicit satisfied state is valid.
$unsatisfiedSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$unsatisfiedEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange" `
    -Satisfied $false
Add-TestOutcomeIteration $unsatisfiedSession "obs-1" "FormatRange" @($unsatisfiedEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $unsatisfiedSession @("obs-1/e1")))) {
    throw "An unsatisfied host observation was accepted as goal completion"
}

$idempotentSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$idempotentEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange"
Add-TestOutcomeIteration $idempotentSession "obs-1" "FormatRange" @($idempotentEvidence)
if (-not [string]::IsNullOrWhiteSpace((Test-OutcomeContract $idempotentSession @("obs-1/e1")))) {
    throw "An idempotent explicitly satisfied host state was rejected"
}

$factoryFormatTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$factoryFormatTool.Id = "FactoryFormat"
$factoryFormatTool.AppType = "Excel"
$factoryFormatTool.AccessMode = "write"
$factoryFormatTool.OutcomeEffects.Add("property_state")
$factoryFormatCall = [ShareRibbon.Agent.ToolCall]::new()
$factoryFormatCall.ToolId = "FactoryFormat"
$factoryFormatCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse('{"range":"SalesData!D2:D25","fillColor":"#90EE90"}')
$changedButRejectedResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "FactoryFormat",
    "host rejected expected state",
    $null,
    [Newtonsoft.Json.Linq.JObject]::Parse('{"changed":true,"satisfied":false,"targetRefs":["Excel:SalesData!D2:D25"]}'))
$changedButRejectedEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $factoryFormatTool,
    $factoryFormatCall,
    $changedButRejectedResult,
    "obs-factory-1",
    [string[]]@()))
if ($changedButRejectedEvidence.Count -ne 1 -or $changedButRejectedEvidence[0].Satisfied) {
    throw "OutcomeEvidenceFactory still treats changed=true as satisfied=true"
}
if (-not $changedButRejectedEvidence[0].InvalidatesPrior) {
    throw "A changed-but-unverified write did not revoke older state evidence"
}
$idempotentHostResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "FactoryFormat",
    "already in expected state",
    $null,
    [Newtonsoft.Json.Linq.JObject]::Parse('{"changed":false,"satisfied":true,"targetRefs":["Excel:SalesData!D2:D25"]}'))
$idempotentHostEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $factoryFormatTool,
    $factoryFormatCall,
    $idempotentHostResult,
    "obs-factory-2",
    [string[]]@()))
if ($idempotentHostEvidence.Count -ne 1 -or -not $idempotentHostEvidence[0].Satisfied) {
    throw "OutcomeEvidenceFactory rejected explicit idempotent host satisfaction"
}

# A structured host verification may safely link user-facing tool parameters to normalized
# COM properties. This keeps high-level tools usable without falling back to raw request data.
$verifiedAliasCall = [ShareRibbon.Agent.ToolCall]::new()
$verifiedAliasCall.ToolId = "FactoryFormat"
$verifiedAliasCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"range":"SalesData!D2:D25","backgroundColor":"#90EE90","fontColor":"#90EE90"}')
$verifiedAliasObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "writeExpected":true,
  "changed":true,
  "satisfied":true,
  "verification":[{
    "targetRef":"Excel:workbooks/active/worksheets/SalesData/ranges/D2:D25/interior",
    "property":"Color",
    "status":"passed",
    "required":true,
    "expected":9498256,
    "actual":9498256,
    "requestProperty":"backgroundColor",
    "requestExpected":"#90EE90"
  }]
}
'@)
$verifiedAliasResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "FactoryFormat",
    "normalized color verified",
    $null,
    $verifiedAliasObservation)
$verifiedAliasEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $factoryFormatTool,
    $verifiedAliasCall,
    $verifiedAliasResult,
    "obs-alias-1",
    [string[]]@(),
    1))
$aliasRequirement = New-TestOutcomeRequirement `
    -Id "verified-background" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "backgroundColor" `
    -ExpectedValue $green
$aliasSession = New-TestOutcomeSession "format background" @($aliasRequirement)
Add-TestOutcomeIteration $aliasSession "obs-alias-1" "FactoryFormat" @($verifiedAliasEvidence)
if (-not $verifiedAliasEvidence[0].RequestVerified -or
    -not [string]::IsNullOrWhiteSpace((Test-OutcomeContract $aliasSession @("obs-alias-1/e1")))) {
    throw "A host-verified high-level parameter alias was not accepted as grounded evidence"
}
$unrelatedAliasRequirement = New-TestOutcomeRequirement `
    -Id "unverified-font" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fontColor" `
    -ExpectedValue $green
$unrelatedAliasSession = New-TestOutcomeSession "format font" @($unrelatedAliasRequirement)
Add-TestOutcomeIteration $unrelatedAliasSession "obs-alias-1" "FactoryFormat" @($verifiedAliasEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $unrelatedAliasSession @("obs-alias-1/e1")))) {
    throw "Verification of one request field implicitly blessed an unrelated tool parameter"
}

# A host hash can prove that WriteData committed the user-level matrix, but contracts should
# compare the matrix rather than depend on an observer-internal hash representation.
$writeFactoryTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$writeFactoryTool.Id = "WriteData"
$writeFactoryTool.AppType = "Excel"
$writeFactoryTool.AccessMode = "write"
$writeMatrix = [Newtonsoft.Json.Linq.JArray]::Parse('[["Region","Average"],["East",1500.0],["North",2000.0]]')
$writeFactoryCall = [ShareRibbon.Agent.ToolCall]::new()
$writeFactoryCall.ToolId = "WriteData"
$writeFactoryCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse('{"targetRange":"PythonAverage!A1:B3"}')
$writeFactoryCall.Parameters["data"] = $writeMatrix.DeepClone()
$writeProjectionObservation = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"kind":"office_operation_batch","writeExpected":true,"changed":true,"satisfied":true}')
$writeProjectionVerification = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"targetRef":"Excel:workbooks/active/worksheets/PythonAverage/ranges/A1:B3","effectType":"data_state","property":"ValueHash","status":"passed","required":true,"expected":"verified-hash","actual":"verified-hash","requestProperty":"data"}')
$writeProjectionVerification["requestExpected"] = $writeMatrix.DeepClone()
$writeProjectionObservation["verification"] = [Newtonsoft.Json.Linq.JArray]::new($writeProjectionVerification)
$writeProjectionResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "WriteData", "matrix hash verified", $null, $writeProjectionObservation)
$writeProjectionEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $writeFactoryTool,
    $writeFactoryCall,
    $writeProjectionResult,
    "obs-write-projection",
    [string[]]@(),
    2,
    "Book.xlsx"))
$writeProjectionRequirement = New-TestOutcomeRequirement `
    -Id "semantic-write" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/PythonAverage/ranges/A1:B3" `
    -EffectType "data_state" `
    -PropertyName "data" `
    -ExpectedValue $writeMatrix
$writeProjectionSession = New-TestOutcomeSession "write exact matrix" @($writeProjectionRequirement)
Add-TestOutcomeIteration $writeProjectionSession "obs-write-projection" "WriteData" @($writeProjectionEvidence)
if (-not $writeProjectionEvidence[0].RequestVerified -or
    -not [string]::IsNullOrWhiteSpace((Test-OutcomeContract $writeProjectionSession @("obs-write-projection/e1")))) {
    throw "A host-verified WriteData matrix could not satisfy its semantic data contract"
}

# The generic Office object bridge may prove only the property state explicitly verified by
# the host. One successful property mutation must not fan out into object existence/absence,
# data, order, filter, and artifact evidence.
$objectBridgeTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$objectBridgeTool.Id = "OfficeObjectOperation"
$objectBridgeTool.AppType = "Excel"
$objectBridgeTool.AccessMode = "write"
$objectBridgeCall = [ShareRibbon.Agent.ToolCall]::new()
$objectBridgeCall.ToolId = "OfficeObjectOperation"
$objectBridgeCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse('{"batch":{"schemaVersion":"1.0"}}')
$objectBridgeObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "writeExpected":true,
  "changed":true,
  "satisfied":true,
  "targetRefs":["Excel:workbooks/active/worksheets/Sheet1"],
  "verification":[{
    "targetRef":"Excel:workbooks/active/worksheets/Sheet1",
    "effectType":"property_state",
    "property":"Name",
    "status":"passed",
    "required":true,
    "expected":"Sheet1",
    "actual":"Sheet1"
  }]
}
'@)
$objectBridgeResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "OfficeObjectOperation",
    "verified one property",
    $null,
    $objectBridgeObservation)
$objectBridgeEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $objectBridgeTool,
    $objectBridgeCall,
    $objectBridgeResult,
    "obs-object-1",
    [string[]]@()))
if ($objectBridgeEvidence.Count -ne 1 -or
    $objectBridgeEvidence[0].EffectType -ne "property_state" -or
    $objectBridgeEvidence[0].PropertyName -ne "Name" -or
    $objectBridgeEvidence[0].Actual["Name"].ToString() -ne "Sheet1") {
    throw "OfficeObjectOperation fabricated effects beyond its verified property state"
}

# A required verification array is authoritative per target. Global satisfied/after fields
# must not turn a failed target verification into positive completion evidence.
$contradictoryObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "changed":true,
  "satisfied":true,
  "after":{"Name":"Wanted"},
  "verification":[{
    "targetRef":"Excel:workbooks/active/worksheets/Sheet1",
    "property":"Name",
    "status":"failed",
    "required":true,
    "expected":"Wanted",
    "actual":"Wrong"
  }]
}
'@)
$contradictoryResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "OfficeObjectOperation",
    "global flag conflicts with required verification",
    $null,
    $contradictoryObservation)
$contradictoryEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $objectBridgeTool,
    $objectBridgeCall,
    $contradictoryResult,
    "obs-object-2",
    [string[]]@(),
    2))
if ($contradictoryEvidence.Count -ne 1 -or
    $contradictoryEvidence[0].Satisfied -or
    -not $contradictoryEvidence[0].InvalidatesPrior -or
    $contradictoryEvidence[0].Actual["Name"].ToString() -ne "Wrong") {
    throw "A failed required verification was overridden by global satisfied/after state"
}

# Verification tuples are atomic. A mixed batch that deletes Chart1 and reads Chart2.Name
# must never manufacture object_absent evidence for Chart2 from batch-level target/effect sets.
$mixedObjectObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "writeExpected":true,
  "changed":true,
  "satisfied":true,
  "verification":[
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/charts/Chart1","effectType":"object_absent","property":"Exists","status":"passed","required":true,"expected":false,"actual":false},
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/charts/Chart2","effectType":"read_coverage","property":"Name","status":"passed","required":true,"expected":"Chart2","actual":"Chart2"}
  ]
}
'@)
$mixedObjectResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "OfficeObjectOperation", "mixed batch verified", $null, $mixedObjectObservation)
$mixedObjectEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $objectBridgeTool,
    $objectBridgeCall,
    $mixedObjectResult,
    "obs-object-mixed",
    [string[]]@(),
    3,
    "Book.xlsx"))
$chart2AbsentRequirement = New-TestOutcomeRequirement `
    -Id "chart-2-absent" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData/charts/Chart2" `
    -EffectType "object_absent" `
    -Operator "exists"
$mixedObjectSession = New-TestOutcomeSession "delete Chart2" @($chart2AbsentRequirement)
Add-TestOutcomeIteration $mixedObjectSession "obs-object-mixed" "OfficeObjectOperation" @($mixedObjectEvidence)
$mixedObjectClaims = @($mixedObjectEvidence | ForEach-Object { $_.EvidenceId })
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $mixedObjectSession $mixedObjectClaims))) {
    throw "A mixed-target verification array manufactured Chart2 deletion evidence"
}

# A passed artifact anchor cannot override a failed required verification in the same host
# observation, even if an inconsistent producer reports Result.Success=true.
$chartFactoryTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$chartFactoryTool.Id = "CreateChart"
$chartFactoryTool.AppType = "Excel"
$chartFactoryTool.AccessMode = "write"
$chartFactoryCall = [ShareRibbon.Agent.ToolCall]::new()
$chartFactoryCall.ToolId = "CreateChart"
$chartFactoryCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse('{"position":"SalesData!D1","type":"line"}')
$failedChartObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "writeExpected":true,
  "changed":true,
  "satisfied":false,
  "verification":[
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/chartObjects/Chart1/chart","effectType":"artifact","property":"ChartType","status":"failed","required":true,"expected":4,"actual":51}
  ],
  "artifactAnchors":[
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/ranges/D1","artifactRef":"Excel:workbooks/active/worksheets/SalesData/chartObjects/Chart1","status":"passed"}
  ]
}
'@)
$failedChartResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "CreateChart", "inconsistent chart result", $null, $failedChartObservation)
$failedChartEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $chartFactoryTool,
    $chartFactoryCall,
    $failedChartResult,
    "obs-chart-required-failed",
    [string[]]@(),
    4,
    "Book.xlsx"))
if ($failedChartEvidence | Where-Object { $_.Satisfied -and $_.TargetRef -match '/ranges/D1$' }) {
    throw "A failed required chart verification was bypassed by a positive artifact anchor"
}

# A successful chart produces two distinct proof families: host properties of the generated
# chart and the stable placement anchor. Catalog fallback must not relabel every atomic
# property assertion as an artifact merely because CreateChart creates an artifact.
$passedChartObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "writeExpected":true,
  "changed":true,
  "satisfied":true,
  "verification":[
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/chartObjects/Chart2/chart","effectType":"property_state","property":"ChartType","status":"passed","required":true,"expected":4,"actual":4,"requestProperty":"type","requestExpected":"line"},
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/chartObjects/Chart2/chart/charttitle","effectType":"property_state","property":"Text","status":"passed","required":true,"expected":"Daily Sales","actual":"Daily Sales","requestProperty":"title","requestExpected":"Daily Sales"}
  ],
  "artifactAnchors":[
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData/ranges/M1","artifactRef":"Excel:workbooks/active/worksheets/SalesData/chartObjects/Chart2","status":"passed"}
  ]
}
'@)
$passedChartResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "CreateChart", "chart verified", $null, $passedChartObservation)
$passedChartEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $chartFactoryTool,
    $chartFactoryCall,
    $passedChartResult,
    "obs-chart-passed",
    [string[]]@(),
    5,
    "Book.xlsx"))
if (@($passedChartEvidence | Where-Object EffectType -eq "property_state").Count -ne 2 -or
    @($passedChartEvidence | Where-Object EffectType -eq "artifact").Count -ne 1 -or
    -not @($passedChartEvidence | Where-Object PropertyName -eq "ChartType")[0].RequestVerified) {
    throw "CreateChart host properties and artifact anchor were not preserved as separate evidence"
}

# ConditionalFormat uses a compact single verification object rather than a batch array. It
# must still expose a concrete property slot so deterministic completion can bind it.
$conditionalTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$conditionalTool.Id = "ConditionalFormat"
$conditionalTool.AppType = "Excel"
$conditionalTool.AccessMode = "write"
$conditionalCall = [ShareRibbon.Agent.ToolCall]::new()
$conditionalCall.ToolId = "ConditionalFormat"
$conditionalCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"range":"SalesData!D2:D25","rule":"highlight","condition":">2500","color":"#90EE90"}')
$conditionalObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "effectType":"property_state",
  "changed":true,
  "satisfied":true,
  "targetRefs":["Excel:workbooks/active/worksheets/SalesData/ranges/D2:D25"],
  "verification":{"satisfied":true,"property":"FormatConditions","expected":{"rule":"highlight"},"requestProperty":"FormatConditions","requestExpected":{"rule":"highlight","condition":">2500","color":"#90EE90"}}
}
'@)
$conditionalResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "ConditionalFormat", "conditional format verified", $null, $conditionalObservation)
$conditionalEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $conditionalTool,
    $conditionalCall,
    $conditionalResult,
    "obs-conditional",
    [string[]]@(),
    6,
    "Book.xlsx"))
if ($conditionalEvidence.Count -ne 1 -or
    $conditionalEvidence[0].EffectType -ne "property_state" -or
    $conditionalEvidence[0].PropertyName -ne "FormatConditions") {
    throw "ConditionalFormat compact verification did not expose its observable property"
}

# Failed declarative batches often have a targetRef but no resultRef. Binding active-workbook
# aliases in that observation must not throw and hide the actual COM/schema error.
$missingResultRefObservation = [Newtonsoft.Json.Linq.JObject]::Parse(@'
{
  "kind":"office_operation_batch",
  "writeExpected":true,
  "changed":false,
  "satisfied":false,
  "operations":[
    {"targetRef":"Excel:workbooks/active/worksheets/SalesData","status":"failed"}
  ]
}
'@)
$missingResultRefFailure = [ShareRibbon.Agent.ToolResult]::Failed(
    "OfficeObjectOperation",
    "host operation failed",
    $null,
    "COM_ERROR",
    "host operation failed",
    "HRESULT:0x800AC472",
    $true,
    $missingResultRefObservation)
$missingResultRefTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$missingResultRefTool.Id = "OfficeObjectOperation"
$missingResultRefTool.AppType = "Excel"
$missingResultRefTool.AccessMode = "write"
$missingResultRefCall = [ShareRibbon.Agent.ToolCall]::new()
$missingResultRefCall.ToolId = "OfficeObjectOperation"
$missingResultRefCall.Parameters = [Newtonsoft.Json.Linq.JObject]::new()
try {
    $null = [ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
        $missingResultRefTool,
        $missingResultRefCall,
        $missingResultRefFailure,
        "obs-missing-result-ref",
        [string[]]@(),
        5,
        "Book.xlsx")
} catch {
    throw "Missing optional resultRef crashed evidence binding instead of preserving the real tool failure: $($_.Exception.Message)"
}

# Explicit invalidation/artifact records are world changes even when an adapter has no
# generic changed/after field. Otherwise their revision can tie with stale evidence.
$invalidationOnlyTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$invalidationOnlyTool.Id = "InvalidateOnly"
$invalidationOnlyTool.AccessMode = "write"
$invalidationOnlyResult = [ShareRibbon.Agent.ToolResult]::Succeed(
    "InvalidateOnly",
    "structural invalidation",
    $null,
    [Newtonsoft.Json.Linq.JObject]::Parse('{"changed":false,"invalidationRefs":["Excel:SalesData"]}'))
if (-not [ShareRibbon.Agent.OutcomeEvidenceFactory]::ObservationAdvancesWorld(
    $invalidationOnlyTool,
    $invalidationOnlyResult)) {
    throw "Explicit invalidation refs did not advance the evidence world revision"
}

# Object identity uses canonical path segments. Sheet1 must never be confused with Sheet10.
$sheet1Requirement = New-TestOutcomeRequirement `
    -Id "sheet-created" `
    -TargetRef "Excel:Sheet1" `
    -EffectType "object_exists"
$sheetBoundarySession = New-TestOutcomeSession "create Sheet1" @($sheet1Requirement)
$sheet10Evidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:Sheet10" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
Add-TestOutcomeIteration $sheetBoundarySession "obs-1" "CreateSheet" @($sheet10Evidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $sheetBoundarySession @("obs-1/e1")))) {
    throw "Sheet1 outcome contract incorrectly matched Sheet10 evidence"
}

$cellSuffixEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-2/e1" `
    -IterationEvidenceId "obs-2" `
    -TargetRef "Excel:OtherSheet!EET1" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
$objectSuffixSession = New-TestOutcomeSession "create Sheet1" @($sheet1Requirement)
Add-TestOutcomeIteration $objectSuffixSession "obs-2" "CreateSheet" @($cellSuffixEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $objectSuffixSession @("obs-2/e1")))) {
    throw "Worksheet object Sheet1 was misparsed as the cell suffix EET1"
}

# Child Office objects retain exact identity. A chart on the right worksheet is not proof
# that another chart on that worksheet exists.
$chart2Requirement = New-TestOutcomeRequirement `
    -Id "chart-2-exists" `
    -TargetRef "Excel:workbooks/active/worksheets/SalesData/charts/Chart2" `
    -EffectType "artifact"
$chartIdentitySession = New-TestOutcomeSession "create Chart2" @($chart2Requirement)
$chart1Evidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-chart/e1" `
    -IterationEvidenceId "obs-chart" `
    -TargetRef "Excel:workbooks/active/worksheets/SalesData/charts/Chart1" `
    -EffectType "artifact" `
    -SourceToolId "CreateChart"
Add-TestOutcomeIteration $chartIdentitySession "obs-chart" "CreateChart" @($chart1Evidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $chartIdentitySession @("obs-chart/e1")))) {
    throw "Distinct child objects on the same worksheet were collapsed to worksheet identity"
}

# Standard Excel:workbooks/... refs must retain workbook identity. The Excel: prefix is part
# of the URI scheme, not a reason to fall back to whichever workbook is currently active.
$bookBRequirement = New-TestOutcomeRequirement `
    -Id "book-b-sheet" `
    -TargetRef "Excel:workbooks/BookB.xlsx/worksheets/SalesData" `
    -EffectType "object_exists" `
    -Operator "exists"
$crossWorkbookSession = New-TestOutcomeSession "create sheet in BookB" @($bookBRequirement)
$bookAEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-book-a/e1" `
    -IterationEvidenceId "obs-book-a" `
    -TargetRef "Excel:workbooks/BookA.xlsx/worksheets/SalesData" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
Add-TestOutcomeIteration $crossWorkbookSession "obs-book-a" "CreateSheet" @($bookAEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $crossWorkbookSession @("obs-book-a/e1")))) {
    throw "Object evidence crossed explicit workbook boundaries"
}

# Binding applies to canonical active-workbook URIs as well as short Sheet!Range refs.
$boundAliasEvidence = @([ShareRibbon.Agent.OutcomeEvidenceFactory]::Create(
    $factoryFormatTool,
    $verifiedAliasCall,
    $verifiedAliasResult,
    "obs-bound-workbook",
    [string[]]@(),
    5,
    "Book A.xlsx"))
if ($boundAliasEvidence.Count -eq 0 -or
    $boundAliasEvidence[0].TargetRef -notmatch 'Excel:workbooks/Book%20A\.xlsx/') {
    throw "Canonical active-workbook evidence was not frozen to the observed workbook"
}

# Ordinary user data, worksheet names, and titles use exact string semantics. Formatting
# enums/colors remain deliberately case-insensitive.
$exactNameRequirement = New-TestOutcomeRequirement `
    -Id "exact-name" `
    -TargetRef "Excel:SalesData" `
    -EffectType "property_state" `
    -PropertyName "Name" `
    -ExpectedValue ([Newtonsoft.Json.Linq.JValue]::CreateString("A B"))
$exactNameSession = New-TestOutcomeSession "set exact name" @($exactNameRequirement)
$collapsedNameEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-name/e1" `
    -IterationEvidenceId "obs-name" `
    -TargetRef "Excel:SalesData" `
    -EffectType "property_state" `
    -PropertyName "Name" `
    -Expected ([Newtonsoft.Json.Linq.JValue]::CreateString("AB")) `
    -SourceToolId "RenameSheet"
Add-TestOutcomeIteration $exactNameSession "obs-name" "RenameSheet" @($collapsedNameEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $exactNameSession @("obs-name/e1")))) {
    throw "Whitespace-distinct ordinary strings were treated as equal"
}

$caseNameRequirement = New-TestOutcomeRequirement `
    -Id "case-name" `
    -TargetRef "Excel:SalesData" `
    -EffectType "property_state" `
    -PropertyName "Name" `
    -ExpectedValue ([Newtonsoft.Json.Linq.JValue]::CreateString("ABC"))
$caseNameSession = New-TestOutcomeSession "set case-sensitive name" @($caseNameRequirement)
$lowerNameEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-name-case/e1" `
    -IterationEvidenceId "obs-name-case" `
    -TargetRef "Excel:SalesData" `
    -EffectType "property_state" `
    -PropertyName "Name" `
    -Expected ([Newtonsoft.Json.Linq.JValue]::CreateString("abc")) `
    -SourceToolId "RenameSheet"
Add-TestOutcomeIteration $caseNameSession "obs-name-case" "RenameSheet" @($lowerNameEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $caseNameSession @("obs-name-case/e1")))) {
    throw "Case-distinct ordinary strings were treated as equal"
}

$semanticColorRequirement = New-TestOutcomeRequirement `
    -Id "semantic-color" `
    -TargetRef "Excel:SalesData!D2" `
    -EffectType "property_state" `
    -PropertyName "Color" `
    -ExpectedValue ([Newtonsoft.Json.Linq.JValue]::CreateString("#90EE90"))
$semanticColorSession = New-TestOutcomeSession "set semantic color" @($semanticColorRequirement)
$lowerColorEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-color/e1" `
    -IterationEvidenceId "obs-color" `
    -TargetRef "Excel:SalesData!D2" `
    -EffectType "property_state" `
    -PropertyName "Color" `
    -Expected ([Newtonsoft.Json.Linq.JValue]::CreateString("#90ee90")) `
    -SourceToolId "FormatRange"
Add-TestOutcomeIteration $semanticColorSession "obs-color" "FormatRange" @($lowerColorEvidence)
if (-not [string]::IsNullOrWhiteSpace((Test-OutcomeContract $semanticColorSession @("obs-color/e1")))) {
    throw "Known semantic formatting values lost controlled normalization"
}

# A stable chart-position anchor is no longer current after the related generated object is
# mutated. Require a fresh host anchor observation instead of accepting stale creation proof.
$anchoredChartRequirement = New-TestOutcomeRequirement `
    -Id "chart-at-d1" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData/ranges/D1" `
    -EffectType "artifact" `
    -Operator "exists"
$movedChartSession = New-TestOutcomeSession "create chart at D1" @($anchoredChartRequirement)
$anchoredChartEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-anchor/e1" `
    -IterationEvidenceId "obs-anchor" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData/ranges/D1" `
    -EffectType "artifact" `
    -SourceToolId "CreateChart" `
    -WorldRevision 1
$anchoredChartEvidence.RelatedTargetRefs.Add("Excel:workbooks/Book.xlsx/worksheets/SalesData/chartObjects/Chart1")
$movedChartEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-move/e1" `
    -IterationEvidenceId "obs-move" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData/chartObjects/Chart1" `
    -EffectType "property_state" `
    -PropertyName "Left" `
    -Expected ([Newtonsoft.Json.Linq.JToken]::FromObject([double]200.0)) `
    -SourceToolId "OfficeObjectOperation" `
    -InvalidatesPrior $true `
    -WorldRevision 2
Add-TestOutcomeIteration $movedChartSession "obs-anchor" "CreateChart" @($anchoredChartEvidence)
Add-TestOutcomeIteration $movedChartSession "obs-move" "OfficeObjectOperation" @($movedChartEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $movedChartSession @("obs-anchor/e1")))) {
    throw "A related chart mutation left stale position-anchor evidence valid"
}

# criterionIds are structural coverage, not permission to collapse independent outcomes.
# One object-existence assertion cannot simultaneously prove both creation and data writing.
$criterionRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($id in @("CreateSheet", "WriteData", "FormatRange")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $id
    $descriptor.Name = $id
    $descriptor.AppType = "excel"
    $descriptor.AccessMode = "write"
    $criterionRegistry.RegisterTool($descriptor)
}
$criterionLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $criterionRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    [ShareRibbon.Agent.PromptManager]::new((Join-Path $repoRoot "ShareRibbon\Prompts")))
$criterionSession = [ShareRibbon.Agent.AgentSession]::new("create and write", "Excel", "")
$criterionSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$criterionSession.Spec.MutationPolicy = "allow"
$criterionSession.Spec.Goal = "create and write"
$criterionSession.Spec.SuccessCriteria.Add("Target sheet exists")
$criterionSession.Spec.SuccessCriteria.Add("Requested values are written")
$weakRequirement = New-TestOutcomeRequirement `
    -Id "only-create" `
    -TargetRef "Excel:PythonAverage" `
    -EffectType "object_exists"
$weakRequirement.CriterionIds.Add("criterion-1")
$weakRequirement.CriterionIds.Add("criterion-2")
$weakPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$weakPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$weakPlan.OutcomeContract.Requirements.Add($weakRequirement)
$freezeMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "SealVerificationProjection",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)

# A worksheet lifecycle target may be emitted in workbook!worksheet shorthand.  It is an
# object identity, not a range, and must be normalized before the frozen plan is validated.
$workbookSheetSession = [ShareRibbon.Agent.AgentSession]::new("create summary worksheet", "Excel", "")
$workbookSheetSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$workbookSheetSession.Spec.MutationPolicy = "allow"
$workbookSheetSession.Spec.Goal = "create summary worksheet"
$workbookSheetSession.Spec.SuccessCriteria.Add("Summary worksheet exists")
$unicodeWorkbookName = (-join ([char[]]@(0x5DE5, 0x4F5C, 0x7C3F))) + "1.xlsx"
$unicodeWorksheetName = (-join ([char[]]@(0x533A, 0x57DF, 0x6C47, 0x603B))) + "0821"
$workbookSheetRequirement = New-TestOutcomeRequirement `
    -Id "summary-sheet-exists" `
    -TargetRef "$unicodeWorkbookName!$unicodeWorksheetName" `
    -EffectType "object_exists"
$workbookSheetRequirement.CriterionIds.Add("criterion-1")
$workbookSheetPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$workbookSheetPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$workbookSheetPlan.OutcomeContract.Requirements.Add($workbookSheetRequirement)
$workbookSheetError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($workbookSheetSession, $workbookSheetPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($workbookSheetError) -or
    -not $workbookSheetPlan.OutcomeContract.Frozen) {
    throw "A workbook-qualified worksheet object target was rejected as an invalid range: $workbookSheetError"
}
$expectedWorkbookSheetRef = "Excel:workbooks/$([Uri]::EscapeDataString($unicodeWorkbookName))/worksheets/$([Uri]::EscapeDataString($unicodeWorksheetName))"
if ($workbookSheetRequirement.TargetRef -ne $expectedWorkbookSheetRef) {
    throw "Workbook-qualified worksheet shorthand was not normalized to a stable worksheet object ref: $($workbookSheetRequirement.TargetRef)"
}
$workbookSheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-create-summary/e1" `
    -IterationEvidenceId "obs-create-summary" `
    -TargetRef "Excel:workbooks/$unicodeWorkbookName/worksheets/$unicodeWorksheetName" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
Add-TestOutcomeIteration `
    $workbookSheetSession `
    "obs-create-summary" `
    "CreateSheet" `
    @($workbookSheetEvidence)
if (-not [string]::IsNullOrWhiteSpace(
    (Test-OutcomeContract $workbookSheetSession @("obs-create-summary/e1")))) {
    throw "Canonical CreateSheet evidence did not satisfy a normalized workbook-qualified worksheet target"
}

$freezeError = [string]$freezeMethod.Invoke($criterionLoop, @($criterionSession, $weakPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($freezeError) -or $weakPlan.OutcomeContract.Frozen) {
    throw "Multiple independent success criteria were collapsed into one weak outcome requirement"
}

# Once GoalContract exists it is the only source of completion obligations. A stale legacy
# SuccessCriteria list cannot hide a required Goal criterion or rename it to criterion-N.
$twoCriterionGoal = New-TestGoalContract `
    -RawRequest "create target worksheet; write result data" `
    -ClauseTexts @("create target worksheet", "write result data") `
    -CriterionIds @("goal-sheet-exists", "goal-result-written") `
    -CriterionKinds @("state", "state")
$goalOmissionSession = New-TestGoalSession $twoCriterionGoal
$goalOmissionSession.Spec.SuccessCriteria.Add("stale legacy criterion")
$goalOmissionRequirement = New-TestOutcomeRequirement `
    -Id "goal-only-sheet" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists"
$goalOmissionRequirement.CriterionIds.Add("goal-sheet-exists")
$goalOmissionPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$goalOmissionPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$goalOmissionPlan.OutcomeContract.Requirements.Add($goalOmissionRequirement)
$goalOmissionError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($goalOmissionSession, $goalOmissionPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($goalOmissionError) -or
    $goalOmissionPlan.OutcomeContract.Frozen -or
    -not $goalOmissionError.Contains("goal-result-written")) {
    throw "A stale legacy criterion hid an omitted required Goal criterion"
}

function New-TwoCriterionGoalPlan {
    $sheet = New-TestOutcomeRequirement `
        -Id "goal-sheet" `
        -TargetRef "Excel:GoalOutput" `
        -EffectType "object_exists"
    $sheet.CriterionIds.Add("goal-sheet-exists")
    $expectedRows = [Newtonsoft.Json.Linq.JArray]::Parse(
        '[["Region","Average Sales"],["East",1986.67]]')
    $written = New-TestOutcomeRequirement `
        -Id "goal-write" `
        -TargetRef "Excel:GoalOutput!A1:B2" `
        -EffectType "data_state" `
        -ExpectedValue $expectedRows
    $written.CriterionIds.Add("goal-result-written")
    $plan = [ShareRibbon.Agent.ExecutionPlan]::new()
    $plan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
    $plan.OutcomeContract.Requirements.Add($sheet)
    $plan.OutcomeContract.Requirements.Add($written)
    return $plan
}

$goalBoundSession = New-TestGoalSession $twoCriterionGoal
$goalBoundSession.Spec.SuccessCriteria.Add("legacy must be ignored")
$goalBoundPlan = New-TwoCriterionGoalPlan
$goalBoundError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($goalBoundSession, $goalBoundPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($goalBoundError) -or
    -not $goalBoundPlan.OutcomeContract.Frozen -or
    $goalBoundPlan.OutcomeContract.BindingMode -ne "goal-v1" -or
    $goalBoundPlan.OutcomeContract.BoundGoalContractHash -ne $twoCriterionGoal.ContractHash) {
    throw "A complete Goal criterion mapping was not frozen against the authoritative GoalContract: $goalBoundError"
}
$goalSheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-goal-sheet/e1" `
    -IterationEvidenceId "obs-goal-sheet" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
$goalWriteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-goal-write/e1" `
    -IterationEvidenceId "obs-goal-write" `
    -TargetRef "Excel:GoalOutput!A1:B2" `
    -EffectType "data_state" `
    -Expected $goalBoundPlan.OutcomeContract.Requirements[1].ExpectedValue `
    -SourceToolId "WriteData"
Add-TestOutcomeIteration $goalBoundSession "obs-goal-sheet" "CreateSheet" @($goalSheetEvidence)
Add-TestOutcomeIteration $goalBoundSession "obs-goal-write" "WriteData" @($goalWriteEvidence)
$goalClaims = @("obs-goal-sheet/e1", "obs-goal-write/e1")
if (-not [string]::IsNullOrWhiteSpace((Test-OutcomeContract $goalBoundSession $goalClaims))) {
    throw "A fully evidenced Goal-bound outcome contract was rejected"
}

# Completion projection is a view over typed host evidence, not a second model-authored
# description of the Office world. A worksheet-root guess for a generated data range must be
# grounded to the exact cited observation before the projection is sealed.
$groundProjectionMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "GroundCompletionProjectionInEvidence",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $groundProjectionMethod) {
    throw "LoopEngine completion projection has no evidence-grounding seam"
}
$groundSession = New-TestGoalSession $twoCriterionGoal
$groundSheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-ground-sheet/e1" `
    -IterationEvidenceId "obs-ground-sheet" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
$groundDataActual = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"rowCount":6,"columnCount":3,"valueHash":"verified-host-hash"}')
$groundDataEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-ground-data/e1" `
    -IterationEvidenceId "obs-ground-data" `
    -TargetRef "Excel:GoalOutput!A1:C6" `
    -EffectType "data_state" `
    -Expected $groundDataActual `
    -SourceToolId "WriteData"
Add-TestOutcomeIteration $groundSession "obs-ground-sheet" "CreateSheet" @($groundSheetEvidence)
Add-TestOutcomeIteration $groundSession "obs-ground-data" "WriteData" @($groundDataEvidence)
$groundContract = [ShareRibbon.Agent.OutcomeContract]::new()
$groundSheetRequirement = New-TestOutcomeRequirement `
    -Id "ground-sheet" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists" `
    -Operator "exists"
$groundSheetRequirement.CriterionIds.Add("goal-sheet-exists")
$groundDataRequirement = New-TestOutcomeRequirement `
    -Id "ground-data" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "data_state" `
    -ExpectedValue ([Newtonsoft.Json.Linq.JObject]::Parse('{"headers":["Region","sum"]}')) `
    -Operator "contains"
$groundDataRequirement.CriterionIds.Add("goal-result-written")
$groundContract.Requirements.Add($groundSheetRequirement)
$groundContract.Requirements.Add($groundDataRequirement)
$groundClaims = [System.Collections.Generic.List[string]]::new()
$groundClaims.Add("obs-ground-sheet/e1")
$groundClaims.Add("obs-ground-data/e1")
$groundError = [string]$groundProjectionMethod.Invoke(
    $criterionLoop,
    @($groundSession, $groundContract, $groundClaims))
if (-not [string]::IsNullOrWhiteSpace($groundError) -or
    -not [string]::IsNullOrWhiteSpace($groundSheetRequirement.PropertyName) -or
    $null -ne $groundSheetRequirement.ExpectedValue -or
    $groundDataRequirement.TargetRef -ne "Excel:GoalOutput!A1:C6" -or
    $groundDataRequirement.Operator -ne "equals" -or
    $groundDataRequirement.ExpectedValue.ToString([Newtonsoft.Json.Formatting]::None) -ne
        $groundDataActual.ToString([Newtonsoft.Json.Formatting]::None)) {
    throw "Completion projection was not grounded in exact cited host evidence: $groundError"
}

# A normal final model turn supplies a message after observing the last tool result; it should
# not have to repeat criterionEvidence for every independent Goal clause. Reuse the provisional
# Goal mapping, but ground each requirement to its own exact host assertion instead of copying
# one shared evidence list onto every criterion.
$groundProvisionalMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "BuildGroundedProvisionalCompletionProjection",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $groundProvisionalMethod) {
    throw "LoopEngine completion has no grounded provisional Goal projection path"
}
$provisionalSession = New-TestGoalSession $twoCriterionGoal
$provisionalSession.Plan = New-TwoCriterionGoalPlan
$provisionalSession.Plan.OutcomeContract.Requirements[0].DerivedFromCapability = "CreateSheet"
$provisionalSession.Plan.OutcomeContract.Requirements[1].DerivedFromCapability = "WriteData"
$provisionalSession.Plan.OutcomeContract.Requirements[0].TargetRef = "Excel:GoalOutput"
$provisionalSession.Plan.OutcomeContract.Requirements[1].TargetRef = "Excel:GoalOutput!A1"
$provisionalDataActual = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"rowCount":6,"columnCount":3,"valueHash":"provisional-host-hash"}')
$provisionalSheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-provisional-sheet/e1" `
    -IterationEvidenceId "obs-provisional-sheet" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
$provisionalDataEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-provisional-data/e1" `
    -IterationEvidenceId "obs-provisional-data" `
    -TargetRef "Excel:GoalOutput!A1:C6" `
    -EffectType "data_state" `
    -Expected $provisionalDataActual `
    -SourceToolId "WriteData"
Add-TestOutcomeIteration $provisionalSession "obs-provisional-sheet" "CreateSheet" @($provisionalSheetEvidence)
Add-TestOutcomeIteration $provisionalSession "obs-provisional-data" "WriteData" @($provisionalDataEvidence)
$provisionalClaims = [System.Collections.Generic.List[string]]::new()
$provisionalContract = $null
$provisionalArgs = [object[]]@($provisionalSession, $provisionalClaims, $provisionalContract)
$provisionalError = [string]$groundProvisionalMethod.Invoke($criterionLoop, $provisionalArgs)
$provisionalContract = [ShareRibbon.Agent.OutcomeContract]$provisionalArgs[2]
if (-not [string]::IsNullOrWhiteSpace($provisionalError) -or
    $null -eq $provisionalContract -or
    $provisionalContract.Requirements.Count -ne 2 -or
    $provisionalContract.Requirements[0].TargetRef -ne "Excel:GoalOutput" -or
    $provisionalContract.Requirements[1].TargetRef -ne "Excel:GoalOutput!A1:C6" -or
    $provisionalClaims.Count -ne 2) {
    throw "A multi-criterion final answer could not reuse and ground the provisional Goal mapping: $provisionalError"
}
$groundPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$groundPlan.OutcomeContract = $groundContract
$groundSealError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($groundSession, $groundPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($groundSealError) -or
    -not [string]::IsNullOrWhiteSpace((Test-OutcomeContract $groundSession $groundClaims))) {
    throw "Evidence-grounded completion projection did not pass deterministic verification: $groundSealError"
}

# Tool observations feed the next adaptive decision; they must never become control decisions
# themselves. Only an explicit decision=complete may enter deterministic Goal verification.
$formatModelObservationMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "FormatModelObservation",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $formatModelObservationMethod) {
    throw "Adaptive loop has no model-only diagnostic observation seam"
}
$diagnosticFailure = [ShareRibbon.Agent.ToolResult]::Failed(
    "PythonCompute",
    "Python calculation failed",
    ([Newtonsoft.Json.Linq.JObject]::Parse('{"phase":"python","exitCode":17}')),
    "PYTHON_RUNTIME_ERROR",
    "Python calculation failed",
    (("x" * 3500) + " FULL_DIAGNOSTIC_TAIL"),
    $true,
    ([Newtonsoft.Json.Linq.JObject]::Parse('{"kind":"python","stderr":"KeyError: values","attempt":2}')))
$modelObservation = [string]$formatModelObservationMethod.Invoke(
    $null,
    @($diagnosticFailure, "public failure"))
if (-not $modelObservation.Contains("KeyError: values") -or
    -not $modelObservation.Contains("public failure") -or
    -not $modelObservation.Contains("PYTHON_RUNTIME_ERROR") -or
    -not $modelObservation.Contains("FULL_DIAGNOSTIC_TAIL") -or
    -not $modelObservation.Contains('"exitCode":17') -or
    -not $modelObservation.Contains('"attempt":2')) {
    throw "The next adaptive model decision still cannot see the complete structured tool failure"
}

# A model turn cannot both request another mutation and claim completion with evidence. The
# contradictory turn must be rejected before dispatch so an already-proven side effect is not
# repeated merely because the model emitted the wrong control discriminator.
$parseReactDecisionMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "ParseReactDecision",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$conflictingCompletionMethod = [ShareRibbon.Agent.LoopEngine].GetMethod(
    "HasConflictingCompletionPayload",
    [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic)
if ($null -eq $parseReactDecisionMethod -or $null -eq $conflictingCompletionMethod) {
    throw "Adaptive loop has no contradictory act/completion control guard"
}
$conflictingDecision = $parseReactDecisionMethod.Invoke(
    $criterionLoop,
    @('{"decision":"act","thought":"already done","action":{"tool":"CreateSheet","params":{"name":"Duplicate"}},"message":"done"}'))
if (-not [bool]$conflictingCompletionMethod.Invoke($null, @($conflictingDecision))) {
    throw "An act decision carrying a final completion message would still dispatch another mutation"
}

# Planning and ReAct share the same semantic authority boundary. Once GoalContract exists,
# stale mutable TaskSpec projections must be absent from both prompts.
$legacyPromptMarkers = @(
    "LEGACY_SUCCESS_MARKER",
    "LEGACY_CONSTRAINT_MARKER",
    "LEGACY_OUTPUT_MARKER",
    "LEGACY_CAPABILITY_MARKER",
    "LEGACY_TOOL_MARKER")
$goalBoundSession.Spec.SuccessCriteria.Add($legacyPromptMarkers[0])
$goalBoundSession.Spec.Constraints.Add($legacyPromptMarkers[1])
$goalBoundSession.Spec.ExpectedOutputs.Add($legacyPromptMarkers[2])
$goalBoundSession.Spec.RequiredCapabilities.Add($legacyPromptMarkers[3])
$goalBoundSession.Spec.RequiredTools.Add($legacyPromptMarkers[4])
$goalPromptSkill = [ShareRibbon.Agent.AgentSkill]::new()
$goalPromptSkill.Name = "Goal prompt authority"
$goalPromptSkill.Description = "non-authoritative strategy hints"
$goalPromptSkill.RequiredTools.Add("CreateSheet")
$goalPromptManager = [ShareRibbon.Agent.PromptManager]::new(
    (Join-Path $repoRoot "ShareRibbon\Prompts"))
$goalPlanningPrompt = $goalPromptManager.BuildPlanningPrompt(
    $goalBoundSession,
    "",
    $goalPromptSkill,
    [ShareRibbon.Agent.AgentMemory]::new())
$goalBoundSession.Plan = $goalBoundPlan
$goalReactPrompt = $goalPromptManager.BuildReactPrompt(
    $goalBoundSession,
    [ShareRibbon.Agent.PlanStep]@{ StepNumber = 1; Description = "continue goal" },
    [ShareRibbon.Agent.AgentMemory]::new(),
    "")
foreach ($legacyPromptMarker in $legacyPromptMarkers) {
    if ($goalPlanningPrompt.Contains($legacyPromptMarker) -or
        $goalReactPrompt.Contains($legacyPromptMarker)) {
        throw "A legacy TaskSpec projection leaked into a Goal-authoritative prompt: $legacyPromptMarker"
    }
}
if (-not $goalPlanningPrompt.Contains("goal-sheet-exists") -or
    -not $goalPlanningPrompt.Contains("goal-result-written") -or
    -not $goalReactPrompt.Contains("goal-sheet-exists") -or
    -not $goalReactPrompt.Contains("goal-result-written")) {
    throw "Planning or ReAct prompt omitted an authoritative Goal criterion"
}

# Initial freeze is not the only guard. Removing a Goal mapping after freeze must be caught
# again at the final completion seam even when all host evidence still matches.
$goalBoundPlan.OutcomeContract.Requirements[1].CriterionIds.Clear()
$postFreezeMappingError = Test-OutcomeContract $goalBoundSession $goalClaims
if ([string]::IsNullOrWhiteSpace($postFreezeMappingError)) {
    throw "Final verification accepted a post-freeze Goal mapping mutation"
}

$unknownGoalSession = New-TestGoalSession $twoCriterionGoal
$unknownGoalRequirement = New-TestOutcomeRequirement `
    -Id "legacy-id-in-goal-mode" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists"
$unknownGoalRequirement.CriterionIds.Add("criterion-1")
$unknownGoalPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$unknownGoalPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$unknownGoalPlan.OutcomeContract.Requirements.Add($unknownGoalRequirement)
$unknownGoalError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($unknownGoalSession, $unknownGoalPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($unknownGoalError) -or $unknownGoalPlan.OutcomeContract.Frozen) {
    throw "Goal mode accepted a fabricated legacy criterion id"
}

$extraGoalSession = New-TestGoalSession $twoCriterionGoal
$extraGoalPlan = New-TwoCriterionGoalPlan
$extraMutation = New-TestOutcomeRequirement `
    -Id "unbound-extra-write" `
    -TargetRef "Excel:Unrequested!A1" `
    -EffectType "data_state" `
    -ExpectedValue ([Newtonsoft.Json.Linq.JValue]::CreateString("unrequested"))
$extraGoalPlan.OutcomeContract.Requirements.Add($extraMutation)
$extraGoalError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($extraGoalSession, $extraGoalPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($extraGoalError) -or $extraGoalPlan.OutcomeContract.Frozen) {
    throw "An unbound mutating requirement expanded the frozen GoalContract"
}

$goalHashSession = New-TestGoalSession $twoCriterionGoal
$goalHashPlan = New-TwoCriterionGoalPlan
$goalHashFreezeError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($goalHashSession, $goalHashPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($goalHashFreezeError)) {
    throw "Goal hash fixture could not be frozen: $goalHashFreezeError"
}
$otherGoal = New-TestGoalContract `
    -RawRequest "perform another goal" `
    -ClauseTexts @("perform another goal") `
    -CriterionIds @("goal-other") `
    -CriterionKinds @("state")
$boundGoalHashField = [ShareRibbon.Agent.OutcomeContract].GetField(
    "_boundGoalContractHash",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$boundGoalHashField.SetValue($goalHashPlan.OutcomeContract, $otherGoal.ContractHash)
$goalHashError = Test-OutcomeContract $goalHashSession @()
if ([string]::IsNullOrWhiteSpace($goalHashError) -or
    -not $goalHashError.Contains("GoalContract")) {
    throw "Final verification did not reject a mismatched GoalContract hash binding"
}

$outcomeHashSession = New-TestGoalSession $twoCriterionGoal
$outcomeHashPlan = New-TwoCriterionGoalPlan
$outcomeHashFreezeError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($outcomeHashSession, $outcomeHashPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($outcomeHashFreezeError)) {
    throw "Outcome integrity fixture could not be frozen: $outcomeHashFreezeError"
}
$outcomeHashPlan.OutcomeContract.Requirements[1].ExpectedValue =
    [Newtonsoft.Json.Linq.JArray]::Parse('[["tampered"]]')
$outcomeHashError = Test-OutcomeContract $outcomeHashSession @()
if ([string]::IsNullOrWhiteSpace($outcomeHashError) -or
    -not $outcomeHashError.Contains("OUTCOME_CONTRACT_MUTATED")) {
    throw "Final verification did not reject a modified frozen OutcomeContract"
}

# Persisted sessions without GoalContract retain the old criterion-N mapping during the
# migration, but this compatibility path is never consulted once a GoalContract exists.
$legacySuccessSession = [ShareRibbon.Agent.AgentSession]::new("legacy goal", "Excel", "")
$legacySuccessSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$legacySuccessSession.Spec.MutationPolicy = "allow"
$legacySuccessSession.Spec.SuccessCriteria.Add("first legacy state")
$legacySuccessSession.Spec.SuccessCriteria.Add("second legacy state")
$legacyRequirement1 = New-TestOutcomeRequirement `
    -Id "legacy-first" `
    -TargetRef "Excel:LegacyOne" `
    -EffectType "object_exists"
$legacyRequirement1.CriterionIds.Add("criterion-1")
$legacyRequirement2 = New-TestOutcomeRequirement `
    -Id "legacy-second" `
    -TargetRef "Excel:LegacyTwo" `
    -EffectType "object_exists"
$legacyRequirement2.CriterionIds.Add("criterion-2")
$legacyPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$legacyPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$legacyPlan.OutcomeContract.Requirements.Add($legacyRequirement1)
$legacyPlan.OutcomeContract.Requirements.Add($legacyRequirement2)
$legacyFreezeError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($legacySuccessSession, $legacyPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($legacyFreezeError) -or
    -not $legacyPlan.OutcomeContract.Frozen -or
    $legacyPlan.OutcomeContract.BindingMode -ne "legacy-v1") {
    throw "Legacy criterion-N compatibility was broken: $legacyFreezeError"
}

# A capability criterion is verified by successful execution and compute lineage, not by
# pretending that the capability itself is an Office host-state requirement.
$pythonGoal = New-TestGoalContract `
    -RawRequest "use Python to calculate regional averages" `
    -ClauseTexts @("use Python to calculate regional averages") `
    -CriterionIds @("goal-python-result") `
    -CriterionKinds @("state")
if (-not $pythonGoal.RequiredCapabilities.Contains("PythonCompute")) {
    throw "Python capability evidence did not enter the authoritative GoalContract"
}
$pythonGoalLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $registry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    [ShareRibbon.Agent.PromptManager]::new((Join-Path $repoRoot "ShareRibbon\Prompts")))
$pythonGoalSession = New-TestGoalSession $pythonGoal
$pythonReadRequirement = New-TestOutcomeRequirement `
    -Id "python-source-read" `
    -TargetRef "Excel:SalesData!A1:B5" `
    -EffectType "read_coverage" `
    -Operator "covers"
$pythonResultRequirement = New-TestOutcomeRequirement `
    -Id "python-result" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -DerivedFromCapability "PythonCompute"
$pythonResultRequirement.CriterionIds.Add("goal-python-result")
$pythonGoalPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$pythonGoalPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$pythonGoalPlan.OutcomeContract.Requirements.Add($pythonReadRequirement)
$pythonGoalPlan.OutcomeContract.Requirements.Add($pythonResultRequirement)
$pythonGoalFreezeError = [string]$freezeMethod.Invoke(
    $pythonGoalLoop,
    @($pythonGoalSession, $pythonGoalPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($pythonGoalFreezeError) -or
    -not $pythonGoalPlan.OutcomeContract.Frozen) {
    throw "Capability Goal could not freeze without mapping the capability criterion as host state: $pythonGoalFreezeError"
}
$pythonGoalReadEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-python-read/e1" `
    -IterationEvidenceId "obs-python-read" `
    -TargetRef "Excel:SalesData!A1:B5" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange"
$pythonGoalWriteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-python-write/e1" `
    -IterationEvidenceId "obs-python-write" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -SourceToolId "WriteData" `
    -DependsOn @("obs-python-compute")
Add-TestOutcomeIteration $pythonGoalSession "obs-python-read" "ReadRange" @($pythonGoalReadEvidence)
Add-TestOutcomeIteration $pythonGoalSession "obs-python-write" "WriteData" @($pythonGoalWriteEvidence) @("obs-python-compute")
$pythonGoalClaims = @("obs-python-read/e1", "obs-python-write/e1")
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $pythonGoalSession $pythonGoalClaims))) {
    throw "Host state alone bypassed the required Python capability"
}
Add-TestOutcomeIteration $pythonGoalSession "obs-python-compute" "PythonCompute" @() @("obs-python-read")
if (-not [string]::IsNullOrWhiteSpace(
    (Test-OutcomeContract $pythonGoalSession $pythonGoalClaims))) {
    throw "A complete Goal-bound ReadRange -> PythonCompute -> WriteData proof was rejected"
}

$duplicateCriterionSession = [ShareRibbon.Agent.AgentSession]::new("create and write", "Excel", "")
$duplicateCriterionSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$duplicateCriterionSession.Spec.MutationPolicy = "allow"
$duplicateCriterionSession.Spec.Goal = "create and write"
$duplicateCriterionSession.Spec.SuccessCriteria.Add("Target sheet exists")
$duplicateCriterionSession.Spec.SuccessCriteria.Add("Requested values are written")
$duplicateRequirement1 = New-TestOutcomeRequirement `
    -Id "duplicate-1" `
    -TargetRef "Excel:PythonAverage" `
    -EffectType "object_exists"
$duplicateRequirement1.CriterionIds.Add("criterion-1")
$duplicateRequirement2 = New-TestOutcomeRequirement `
    -Id "duplicate-2" `
    -TargetRef "Excel:workbooks/active/worksheets/PythonAverage" `
    -EffectType "object_exists"
$duplicateRequirement2.CriterionIds.Add("criterion-2")
$duplicatePlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$duplicatePlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$duplicatePlan.OutcomeContract.Requirements.Add($duplicateRequirement1)
$duplicatePlan.OutcomeContract.Requirements.Add($duplicateRequirement2)
$duplicateFreezeError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($duplicateCriterionSession, $duplicatePlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($duplicateFreezeError) -or $duplicatePlan.OutcomeContract.Frozen) {
    throw "Duplicated host assertions were used to fake coverage of independent criteria"
}

# State effects cannot weaken a concrete postcondition into mere existence, and contains
# cannot use an empty value that every observation would satisfy.
$weakExistsSession = [ShareRibbon.Agent.AgentSession]::new("write concrete data", "Excel", "")
$weakExistsSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$weakExistsSession.Spec.MutationPolicy = "allow"
$weakExistsSession.Spec.Goal = "write concrete data"
$weakExistsPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$weakExistsPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$weakExistsPlan.OutcomeContract.Requirements.Add((New-TestOutcomeRequirement `
    -Id "weak-data-exists" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -Operator "exists"))
$weakExistsError = [string]$freezeMethod.Invoke($criterionLoop, @($weakExistsSession, $weakExistsPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($weakExistsError) -or $weakExistsPlan.OutcomeContract.Frozen) {
    throw "A concrete data_state contract was weakened to operator=exists"
}

$vacuousContainsSession = [ShareRibbon.Agent.AgentSession]::new("write concrete data", "Excel", "")
$vacuousContainsSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$vacuousContainsSession.Spec.MutationPolicy = "allow"
$vacuousContainsSession.Spec.Goal = "write concrete data"
$vacuousContainsPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$vacuousContainsPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$vacuousContainsPlan.OutcomeContract.Requirements.Add((New-TestOutcomeRequirement `
    -Id "empty-contains" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -ExpectedValue ([Newtonsoft.Json.Linq.JObject]::new()) `
    -Operator "contains"))
$vacuousContainsError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($vacuousContainsSession, $vacuousContainsPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($vacuousContainsError) -or $vacuousContainsPlan.OutcomeContract.Frozen) {
    throw "A vacuous contains={} contract was accepted"
}

# derivedFromCapability is trusted lineage metadata only after resolving to a registered
# compute descriptor. A mutating tool name cannot waive the expected data postcondition.
$fakeDerivedSession = [ShareRibbon.Agent.AgentSession]::new("write computed data", "Excel", "")
$fakeDerivedSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$fakeDerivedSession.Spec.MutationPolicy = "allow"
$fakeDerivedSession.Spec.Goal = "write computed data"
$fakeDerivedPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$fakeDerivedPlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$fakeDerivedPlan.OutcomeContract.Requirements.Add((New-TestOutcomeRequirement `
    -Id "fake-compute-source" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -DerivedFromCapability "CreateSheet"))
$fakeDerivedError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($fakeDerivedSession, $fakeDerivedPlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($fakeDerivedError) -or $fakeDerivedPlan.OutcomeContract.Frozen) {
    throw "A non-compute tool was accepted as derivedFromCapability"
}

# Different predicate spellings cannot let one host assertion slot satisfy two independent
# success criteria. The planner must split the target/property into distinct observable facts.
$operatorReuseSession = [ShareRibbon.Agent.AgentSession]::new("two independent color outcomes", "Excel", "")
$operatorReuseSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$operatorReuseSession.Spec.MutationPolicy = "allow"
$operatorReuseSession.Spec.Goal = "two independent color outcomes"
$operatorReuseSession.Spec.SuccessCriteria.Add("first color criterion")
$operatorReuseSession.Spec.SuccessCriteria.Add("second color criterion")
$operatorReuseExpected = [Newtonsoft.Json.Linq.JObject]::Parse('{"value":"red"}')
$operatorEqualsRequirement = New-TestOutcomeRequirement `
    -Id "operator-equals" `
    -TargetRef "Excel:SalesData!A1" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -ExpectedValue $operatorReuseExpected `
    -Operator "equals"
$operatorEqualsRequirement.CriterionIds.Add("criterion-1")
$operatorContainsRequirement = New-TestOutcomeRequirement `
    -Id "operator-contains" `
    -TargetRef "Excel:SalesData!A1" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -ExpectedValue $operatorReuseExpected.DeepClone() `
    -Operator "contains"
$operatorContainsRequirement.CriterionIds.Add("criterion-2")
$operatorReusePlan = [ShareRibbon.Agent.ExecutionPlan]::new()
$operatorReusePlan.OutcomeContract = [ShareRibbon.Agent.OutcomeContract]::new()
$operatorReusePlan.OutcomeContract.Requirements.Add($operatorEqualsRequirement)
$operatorReusePlan.OutcomeContract.Requirements.Add($operatorContainsRequirement)
$operatorReuseError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($operatorReuseSession, $operatorReusePlan, "Excel"))
if ([string]::IsNullOrWhiteSpace($operatorReuseError) -or $operatorReusePlan.OutcomeContract.Frozen) {
    throw "equals/contains reused one host assertion slot for two independent criteria"
}

# A failed/partial later write is not completion proof, but its observed mutation revokes
# stale matching state. Failed iterations must remain in the invalidation timeline.
$failedMutationSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$preFailureGreen = New-TestOutcomeEvidence `
    -EvidenceId "obs-before/e1" `
    -IterationEvidenceId "obs-before" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange" `
    -WorldRevision 1
$failedRedObservation = New-TestOutcomeEvidence `
    -EvidenceId "obs-failed/e1" `
    -IterationEvidenceId "obs-failed" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected ([Newtonsoft.Json.Linq.JValue]::CreateString("#FF0000")) `
    -SourceToolId "FormatRange" `
    -Satisfied $false `
    -InvalidatesPrior $true `
    -WorldRevision 2
Add-TestOutcomeIteration $failedMutationSession "obs-before" "FormatRange" @($preFailureGreen)
Add-TestOutcomeIteration $failedMutationSession "obs-failed" "FormatRange" @($failedRedObservation) @() $false
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $failedMutationSession @("obs-before/e1")))) {
    throw "Stale evidence survived a later failed write that changed the observed state"
}

$unknownPropertySession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$unknownPropertyGreen = New-TestOutcomeEvidence `
    -EvidenceId "obs-known/e1" `
    -IterationEvidenceId "obs-known" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange" `
    -WorldRevision 1
$unknownPropertyTombstone = New-TestOutcomeEvidence `
    -EvidenceId "obs-unknown/e1" `
    -IterationEvidenceId "obs-unknown" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -SourceToolId "FormatRange" `
    -Satisfied $false `
    -InvalidatesPrior $true `
    -WorldRevision 2
Add-TestOutcomeIteration $unknownPropertySession "obs-known" "FormatRange" @($unknownPropertyGreen)
Add-TestOutcomeIteration $unknownPropertySession "obs-unknown" "FormatRange" @($unknownPropertyTombstone) @() $false
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $unknownPropertySession @("obs-known/e1")))) {
    throw "An unknown-property mutation tombstone did not conservatively revoke old property proof"
}

$artifactRequirement = New-TestOutcomeRequirement `
    -Id "chart-exists" `
    -TargetRef "Excel:workbooks/active/worksheets/SalesData/charts/Chart1" `
    -EffectType "artifact"
$unknownArtifactSession = New-TestOutcomeSession "create Chart1" @($artifactRequirement)
$oldChartEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-chart-old/e1" `
    -IterationEvidenceId "obs-chart-old" `
    -TargetRef "Excel:workbooks/active/worksheets/SalesData/charts/Chart1" `
    -EffectType "artifact" `
    -SourceToolId "CreateChart" `
    -WorldRevision 1
$unknownArtifactTombstone = New-TestOutcomeEvidence `
    -EvidenceId "obs-chart-failed/e1" `
    -IterationEvidenceId "obs-chart-failed" `
    -TargetRef "*" `
    -EffectType "artifact" `
    -SourceToolId "CreateChart" `
    -Satisfied $false `
    -InvalidatesPrior $true `
    -WorldRevision 2
Add-TestOutcomeIteration $unknownArtifactSession "obs-chart-old" "CreateChart" @($oldChartEvidence)
Add-TestOutcomeIteration $unknownArtifactSession "obs-chart-failed" "CreateChart" @($unknownArtifactTombstone) @() $false
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $unknownArtifactSession @("obs-chart-old/e1")))) {
    throw "A wildcard mutation tombstone did not revoke stale artifact proof"
}

# Worksheet lifecycle changes invalidate all child/range evidence from the previous object.
$deletedSheetSession = New-TestOutcomeSession "format fill color" @($formatRequirement)
$beforeDeleteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-format/e1" `
    -IterationEvidenceId "obs-format" `
    -TargetRef "Excel:SalesData!D2:D25" `
    -EffectType "property_state" `
    -PropertyName "fillColor" `
    -Expected $green `
    -SourceToolId "FormatRange" `
    -WorldRevision 1
$deleteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-delete/e1" `
    -IterationEvidenceId "obs-delete" `
    -TargetRef "Excel:SalesData" `
    -EffectType "object_absent" `
    -SourceToolId "DeleteSheet" `
    -WorldRevision 2
Add-TestOutcomeIteration $deletedSheetSession "obs-format" "FormatRange" @($beforeDeleteEvidence)
Add-TestOutcomeIteration $deletedSheetSession "obs-delete" "DeleteSheet" @($deleteEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $deletedSheetSession @("obs-format/e1")))) {
    throw "Range evidence survived deletion of its containing worksheet"
}

# Row/column structural changes invalidate every older range assertion on that worksheet.
# Canonical Excel:workbooks/... object refs must overlap ranges only within the same workbook.
$structuralReadRequirement = New-TestOutcomeRequirement `
    -Id "structural-read" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData/ranges/D2:D25" `
    -EffectType "read_coverage" `
    -Operator "covers"
$structuralReadSession = New-TestOutcomeSession "read stable source" @($structuralReadRequirement)
$beforeStructureEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-structure-read/e1" `
    -IterationEvidenceId "obs-structure-read" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData/ranges/D2:D25" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange" `
    -WorldRevision 1
$worksheetStructureTombstone = New-TestOutcomeEvidence `
    -EvidenceId "obs-structure-change/e1" `
    -IterationEvidenceId "obs-structure-change" `
    -TargetRef "Excel:workbooks/Book.xlsx/worksheets/SalesData" `
    -EffectType "unclassified_mutation" `
    -SourceToolId "DeleteRowCol" `
    -Satisfied $false `
    -InvalidatesPrior $true `
    -WorldRevision 2
Add-TestOutcomeIteration $structuralReadSession "obs-structure-read" "ReadRange" @($beforeStructureEvidence)
Add-TestOutcomeIteration $structuralReadSession "obs-structure-change" "DeleteRowCol" @($worksheetStructureTombstone) @() $false
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $structuralReadSession @("obs-structure-read/e1")))) {
    throw "A worksheet structural mutation did not invalidate old range-read evidence"
}

# Multiple adjacent observations may jointly cover a requested range, but every contributing
# evidence ID must be cited by the model's completion decision.
$readUnionRequirement = New-TestOutcomeRequirement `
    -Id "full-read" `
    -TargetRef "Excel:SalesData!A1:D25" `
    -EffectType "read_coverage" `
    -Operator "covers"
$readUnionSession = New-TestOutcomeSession "read full table" @($readUnionRequirement)
$readTop = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!A1:D10" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange"
$readBottom = New-TestOutcomeEvidence `
    -EvidenceId "obs-2/e1" `
    -IterationEvidenceId "obs-2" `
    -TargetRef "Excel:SalesData!A11:D25" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange"
Add-TestOutcomeIteration $readUnionSession "obs-1" "ReadRange" @($readTop)
Add-TestOutcomeIteration $readUnionSession "obs-2" "ReadRange" @($readBottom)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $readUnionSession @("obs-1/e1")))) {
    throw "Partial citation was accepted for multi-evidence range coverage"
}
if (-not [string]::IsNullOrWhiteSpace(
    (Test-OutcomeContract $readUnionSession @("obs-1/e1", "obs-2/e1")))) {
    throw "Adjacent read evidence was not accepted as complete range coverage"
}

# Composite goals are conjunctive.  Creating the destination alone cannot prove that the
# requested result was also written into it.
$sheetRequirement = New-TestOutcomeRequirement `
    -Id "destination-exists" `
    -TargetRef "Excel:PythonAverage" `
    -EffectType "object_exists"
$writeRequirement = New-TestOutcomeRequirement `
    -Id "result-written" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state"
$partialGoalSession = New-TestOutcomeSession `
    "create destination and write result" `
    @($sheetRequirement, $writeRequirement)
$sheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:PythonAverage" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
Add-TestOutcomeIteration $partialGoalSession "obs-1" "CreateSheet" @($sheetEvidence)
if ([string]::IsNullOrWhiteSpace((Test-OutcomeContract $partialGoalSession @("obs-1/e1")))) {
    throw "A partially completed composite goal was accepted"
}

# Provenance is a complete ReadRange -> PythonCompute -> WriteData chain, not merely the
# independent existence of one Python action somewhere in the successful history.
$readRequirement = New-TestOutcomeRequirement `
    -Id "source-read" `
    -TargetRef "Excel:SalesData!A1:B5" `
    -EffectType "read_coverage"
$pythonWriteRequirement = New-TestOutcomeRequirement `
    -Id "python-result-written" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -DerivedFromCapability "PythonCompute"

$pythonLineageSession = New-TestOutcomeSession `
    "read, compute with Python, and write" `
    @($readRequirement, $pythonWriteRequirement)
$pythonLineageSession.Spec.OutcomeContract.ValidatedComputeCapabilities.Add("PythonCompute")
Seal-TestOutcomeContract $pythonLineageSession
$readEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-1/e1" `
    -IterationEvidenceId "obs-1" `
    -TargetRef "Excel:SalesData!A1:B5" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange"
Add-TestOutcomeIteration $pythonLineageSession "obs-1" "ReadRange" @($readEvidence)
Add-TestOutcomeIteration $pythonLineageSession "obs-2" "PythonCompute" @() @("obs-1")
$pythonWriteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-3/e1" `
    -IterationEvidenceId "obs-3" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -SourceToolId "WriteData" `
    -DependsOn @("obs-2")
Add-TestOutcomeIteration $pythonLineageSession "obs-3" "WriteData" @($pythonWriteEvidence) @("obs-2")
if (-not [string]::IsNullOrWhiteSpace(
    (Test-OutcomeContract $pythonLineageSession @("obs-1/e1", "obs-3/e1")))) {
    throw "A complete ReadRange -> PythonCompute -> WriteData evidence chain was rejected"
}

$disconnectedPythonSession = New-TestOutcomeSession `
    "read, compute with Python, and write" `
    @($readRequirement, $pythonWriteRequirement)
$disconnectedPythonSession.Spec.OutcomeContract.ValidatedComputeCapabilities.Add("PythonCompute")
Seal-TestOutcomeContract $disconnectedPythonSession
Add-TestOutcomeIteration $disconnectedPythonSession "obs-1" "PythonCompute"
$lateReadEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-2/e1" `
    -IterationEvidenceId "obs-2" `
    -TargetRef "Excel:SalesData!A1:B5" `
    -EffectType "read_coverage" `
    -SourceToolId "ReadRange"
Add-TestOutcomeIteration $disconnectedPythonSession "obs-2" "ReadRange" @($lateReadEvidence)
$disconnectedWriteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-3/e1" `
    -IterationEvidenceId "obs-3" `
    -TargetRef "Excel:PythonAverage!A1:B3" `
    -EffectType "data_state" `
    -SourceToolId "WriteData" `
    -DependsOn @("obs-1")
Add-TestOutcomeIteration $disconnectedPythonSession "obs-3" "WriteData" @($disconnectedWriteEvidence) @("obs-1")
if ([string]::IsNullOrWhiteSpace(
    (Test-OutcomeContract $disconnectedPythonSession @("obs-2/e1", "obs-3/e1")))) {
    throw "PythonCompute -> independent ReadRange -> WriteData was accepted as complete source lineage"
}

# A registered host producer that directly emits the final effect does not need an
# artificial compute -> write chain. Its own successful host evidence is the lineage.
$analysisRequirement = New-TestOutcomeRequirement `
    -Id "regional-summary" `
    -TargetRef "Excel:RegionalSummary!A1:C7" `
    -EffectType "data_state" `
    -DerivedFromCapability "DataAnalysis"
$analysisSession = New-TestOutcomeSession "write a regional summary" @($analysisRequirement)
$analysisSession.Spec.OutcomeContract.ValidatedProducerCapabilities.Add("DataAnalysis")
Seal-TestOutcomeContract $analysisSession
$analysisEvidence = New-TestOutcomeEvidence `
    -EvidenceId "obs-analysis/e1" `
    -IterationEvidenceId "obs-analysis" `
    -TargetRef "Excel:RegionalSummary!A1:C7" `
    -EffectType "data_state" `
    -SourceToolId "DataAnalysis"
Add-TestOutcomeIteration $analysisSession "obs-analysis" "DataAnalysis" @($analysisEvidence)
if (-not [string]::IsNullOrWhiteSpace(
    (Test-OutcomeContract $analysisSession @("obs-analysis/e1")))) {
    throw "A direct DataAnalysis producer was rejected despite matching successful host evidence"
}

$session = [ShareRibbon.Agent.AgentSession]::new("compute", "Excel", "")
$session.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$session.Spec.MutationPolicy = "allow"
$session.Spec.RequiredTools.Add("ReadRange")
$session.Spec.RequiredTools.Add("PythonCompute")
$session.Spec.RequiredTools.Add("WriteData")
$toolContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession($session, $skill)
$hiddenByTaskSpec = @($registry.GetAvailableTools("Excel") | Where-Object {
    -not $toolContext.IsToolAllowed($_)
})
if ($hiddenByTaskSpec.Count -gt 0) {
    throw "TaskSpec still hides Skill-authorized Excel tools: $($hiddenByTaskSpec.Id -join ', ')"
}

# Read-only mutation policy remains a hard boundary and may narrow the Skill to the
# explicitly requested read tools.
$session.Spec.MutationPolicy = "read_only"
$session.Spec.RequiredTools.Clear()
$session.Spec.RequiredTools.Add("ReadRange")
$readOnlyContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession($session, $skill)
if (-not $readOnlyContext.IsToolAllowed($registry.GetTool("ReadRange")) -or
    $readOnlyContext.IsToolAllowed($registry.GetTool("WriteData")) -or
    $readOnlyContext.IsToolAllowed($registry.GetTool("CreateSheet"))) {
    throw "Read-only mutation policy does not remain a hard tool boundary"
}
$session.Spec.RequiredTools.Add("WriteData")
$adversarialReadOnlyContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession($session, $skill)
if ($adversarialReadOnlyContext.IsToolAllowed($registry.GetTool("WriteData")) -or
    @($registry.GetVisibleTools("Excel", $adversarialReadOnlyContext) | Where-Object Id -eq "WriteData").Count -gt 0) {
    throw "A stale/malformed read-only task hint authorized a write-capable descriptor"
}

# Internal planning respects the configured provider mode and does not add a second,
# hidden token or wall-clock budget on top of the provider/user configuration.
$policyType = [ShareRibbon.AgentInternalRequestPolicy]
if ($policyType::ResolveReasoningMode("enabled") -ne "enabled" -or
    $policyType::ResolveReasoningMode("disabled") -ne "disabled" -or
    $null -ne $policyType::MaxTokens -or
    $null -ne $policyType::RequestTimeout) {
    throw "Internal agent request policy still overrides reasoning, tokens, or timeout"
}
$agentRequestOptions = [ShareRibbon.AiRequestOptions]::new()
$agentRequestOptions.ApiUrl = "https://example.invalid/v1/chat/completions"
$agentRequestOptions.ModelName = "reasoning-model"
$agentRequestOptions.Platform = "openai-compatible"
$agentRequestOptions.ReasoningMode = $policyType::ResolveReasoningMode("enabled")
$agentRequestOptions.Messages = [Newtonsoft.Json.Linq.JArray]::new()
$agentRequestOptions.Messages.Add([Newtonsoft.Json.Linq.JObject]::Parse('{"role":"user","content":"plan"}'))
$agentRequestOptions.MaxTokens = $policyType::MaxTokens
$agentRequest = [ShareRibbon.AiGateway]::BuildProviderRequest($agentRequestOptions)
if ($null -ne $agentRequest.Property("max_tokens")) {
    throw "Internal agent request still serializes a max_tokens limit"
}
$transportCancellation = [System.Threading.CancellationTokenSource]::new()
$linkedTransportCancellation = $policyType::CreateCancellationSource($transportCancellation.Token)
$transportCancellation.Cancel()
if (-not $linkedTransportCancellation.IsCancellationRequested) {
    throw "Agent request cancellation is not linked to the provider transport"
}
$linkedTransportCancellation.Dispose()
$transportCancellation.Dispose()

# A generated plan is only a high-level skeleton. Every action must be selected after the
# latest observation, so planned command JSON must never be executed directly.
$testTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$testTool.Id = "TestAction"
$testTool.Name = "Test action"
$testTool.AppType = "excel"
$testTool.RiskLevel = "safe"
$testTool.AccessMode = "write"
$testTool.OutcomeEffects.Add("artifact")
$registry.RegisterTool($testTool)
$script:adaptiveExecutions = 0
$registry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $commandObject = [Newtonsoft.Json.Linq.JObject]::Parse($code)
    if ([bool]$commandObject["params"]["planned"]) {
        throw "Loop executed stale PlanStep.Code instead of the current ReAct action"
    }
    $script:adaptiveExecutions += 1
    $message = if ($script:adaptiveExecutions -eq 1) { "first-observation" } else { "second-observation" }
    $target = if ($script:adaptiveExecutions -eq 1) { "Excel:test/first" } else { "Excel:test/second" }
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"changed","changed":true,"satisfied":true,"targetRefs":[]}')
    $observation["summary"] = [Newtonsoft.Json.Linq.JValue]::CreateString($message)
    $observation["targetRefs"].Add($target)
    return [ShareRibbon.Agent.ToolResult]::Succeed("TestAction", $message, $null, $observation)
}

$promptManager = [ShareRibbon.Agent.PromptManager]::new((Join-Path $repoRoot "ShareRibbon\Prompts"))

# Keep narrative history bounded for latency, but never truncate evidence IDs required to
# complete a long compound task.
$longLedgerRequirement = New-TestOutcomeRequirement `
    -Id "old-evidence" `
    -TargetRef "Excel:LongTask!A1" `
    -EffectType "data_state"
$longLedgerSession = New-TestOutcomeSession "long task" @($longLedgerRequirement)
for ($ledgerIndex = 1; $ledgerIndex -le 8; $ledgerIndex += 1) {
    $ledgerEvidence = New-TestOutcomeEvidence `
        -EvidenceId "obs-$ledgerIndex/e1" `
        -IterationEvidenceId "obs-$ledgerIndex" `
        -TargetRef "Excel:LongTask!A$ledgerIndex" `
        -EffectType "data_state" `
        -SourceToolId "WriteData"
    Add-TestOutcomeIteration $longLedgerSession "obs-$ledgerIndex" "WriteData" @($ledgerEvidence)
}
$ledgerPrompt = $promptManager.BuildReactPrompt(
    $longLedgerSession,
    [ShareRibbon.Agent.PlanStep]@{ StepNumber = 1; Description = "continue" },
    [ShareRibbon.Agent.AgentMemory]::new(),
    "")
if (-not $ledgerPrompt.Contains("obs-1/e1") -or -not $ledgerPrompt.Contains("obs-8/e1")) {
    throw "ReAct prompt truncates evidence needed by long compound tasks"
}

# A tool command is encoded as JSON, but embedded source code remains source code. The
# final system prompt must never prohibit JSON-escaped newlines while PythonCompute
# requires compound statements to be multiline.
$emptyToolList = [System.Collections.Generic.List[ShareRibbon.Agent.ToolDescriptor]]::new()
$excelSystemPrompt = $promptManager.BuildSystemPrompt("Excel", $emptyToolList)
$prohibitsEscapedNewlines = [regex]::IsMatch($excelSystemPrompt, '\\n[^\\r]{0,4}\\r')
$allowsEscapedNewlines = $excelSystemPrompt.Contains('PythonCompute.code') -and
    $excelSystemPrompt.Contains('for/if/try/with') -and
    $excelSystemPrompt.Contains('\n')
if ($prohibitsEscapedNewlines -or -not $allowsEscapedNewlines) {
    throw "Excel prompt still gives contradictory newline rules for embedded Python source: prohibits=$prohibitsEscapedNewlines allows=$allowsEscapedNewlines"
}

# A user-required capability policy is self-contained. The planner should not receive an
# entire broad Skill handbook when semantic classification already identified the method.
$contractSession = [ShareRibbon.Agent.AgentSession]::new("calculate with Python", "Excel", "")
$contractSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$contractSession.Spec.MutationPolicy = "allow"
$contractSession.Spec.Goal = "calculate with Python"
$contractSession.Spec.RequiredTools.Add("PythonCompute")
$contractSession.Spec.RequiredCapabilities.Add("PythonCompute")
$contractSkill = [ShareRibbon.Agent.AgentSkill]::new()
$contractSkill.Name = "Broad Excel skill"
$contractSkill.Description = "general spreadsheet operations"
$contractSkill.RequiredTools.Add("PythonCompute")
$contractSkill.PromptTemplate = "SKILL-HANDBOOK-MARKER-" + ("x" * 7000)
$contractPrompt = $promptManager.BuildPlanningPrompt($contractSession, "", $contractSkill)
if ($contractPrompt.Contains("SKILL-HANDBOOK-MARKER") -or
    -not $contractPrompt.Contains("PythonCompute")) {
    throw "Planner prompt does not use the compact required-capability policy"
}

$memory = [ShareRibbon.Agent.AgentMemory]::new()
$loop = [ShareRibbon.Agent.LoopEngine]::new($registry, $memory, $promptManager)
$script:modelCalls = 0
if (-not ("AgentRuntimeContextProbe" -as [type])) {
    $newtonsoftAssemblyPath = Join-Path (Split-Path $assemblyPath) "Newtonsoft.Json.dll"
    Add-Type -ReferencedAssemblies $assemblyPath,$newtonsoftAssemblyPath -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ShareRibbon;
using ShareRibbon.Agent;
using ShareRibbon.Agent.Context;

public static class AgentRuntimeContextProbe
{
    public static int Captures { get; private set; }
    public static readonly Func<ContextPack> CaptureDelegate = Capture;
    public static int PythonModelCalls { get; private set; }
    public static int PythonHostExecutions { get; private set; }
    public static bool PythonWritePayloadValid { get; private set; }
    public static readonly Func<string, string, List<HistoryMessage>, Task<string>> PythonModelDelegate = PythonModel;
    public static readonly Func<string, string, bool, ToolResult> PythonHostDelegate = PythonHost;

    public static void Reset()
    {
        Captures = 0;
        PythonModelCalls = 0;
        PythonHostExecutions = 0;
        PythonWritePayloadValid = false;
    }

    public static ContextPack Capture()
    {
        Captures++;
        var pack = new ContextPack();
        pack.AppType = "Excel";
        pack.Document.Preview = "context-version-" + Captures;
        return pack;
    }

    public static Task<string> PythonModel(string prompt, string system, List<HistoryMessage> history)
    {
        PythonModelCalls++;
        switch (PythonModelCalls)
        {
            case 1:
                return Task.FromResult("{\"understanding\":\"python workflow\",\"steps\":[{\"step\":1,\"description\":\"read full data\",\"toolHint\":\"ReadRange\"},{\"step\":2,\"description\":\"compute averages\",\"toolHint\":\"PythonCompute\"},{\"step\":3,\"description\":\"create destination\",\"toolHint\":\"CreateSheet\"},{\"step\":4,\"description\":\"write result\",\"toolHint\":\"WriteData\"}],\"summary\":\"written\",\"capabilityGap\":\"\",\"outcomeContract\":{\"schemaVersion\":\"1.0\",\"requirements\":[{\"id\":\"source-read\",\"targetRef\":\"Excel:SalesData!A1:B5\",\"effectType\":\"read_coverage\",\"operator\":\"covers\",\"required\":true},{\"id\":\"destination-created\",\"targetRef\":\"Excel:PythonAverage\",\"effectType\":\"object_exists\",\"operator\":\"exists\",\"criterionIds\":[\"goal-python-sheet\"],\"required\":true},{\"id\":\"python-result-written\",\"targetRef\":\"Excel:PythonAverage!A1:B3\",\"effectType\":\"data_state\",\"derivedFromCapability\":\"PythonCompute\",\"criterionIds\":[\"goal-python-result\"],\"required\":true}]}}");
            case 2:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"read first\",\"action\":{\"tool\":\"ReadRange\",\"params\":{\"range\":\"SalesData!A1:B5\"}}}");
            case 3:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"compute observed rows\",\"action\":{\"tool\":\"PythonCompute\",\"params\":{\"code\":\"groups = {}\\nfor row in input_data['values'][1:]:\\n    key = row[0]\\n    groups.setdefault(key, []).append(float(row[1]))\\nresult = [['Region', 'Average']] + [[key, sum(values) / len(values)] for key, values in sorted(groups.items())]\",\"input\":{\"preview\":true}}}}");
            case 4:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"create destination\",\"action\":{\"tool\":\"CreateSheet\",\"params\":{\"name\":\"PythonAverage\"}}}");
            case 5:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"write computed rows\",\"action\":{\"tool\":\"WriteData\",\"params\":{\"targetRange\":\"PythonAverage!A1\",\"data\":[[\"wrong\",-1]]}}}");
            default:
                return Task.FromResult("{\"decision\":\"complete\",\"thought\":\"all four verified observations exist\",\"message\":\"done\",\"evidence\":[\"obs-3/e1\",\"obs-4/e1\"]}");
        }
    }

    public static ToolResult PythonHost(string code, string language, bool preview)
    {
        PythonHostExecutions++;
        var commandObject = JObject.Parse(code);
        var command = commandObject.Value<string>("command");
        var parameters = (JObject)commandObject["params"];
        if (command == "ReadRange")
        {
            var data = JObject.Parse("{\"workbook\":\"book.xlsx\",\"sheet\":\"SalesData\",\"address\":\"A1:B5\",\"rowCount\":5,\"columnCount\":2,\"values\":[[\"Region\",\"Sales\"],[\"East\",1000],[\"East\",2000],[\"North\",1500],[\"North\",2500]]}");
            return ToolResult.Succeed("ReadRange", "full-read-observed", data);
        }
        if (command == "CreateSheet")
        {
            var observation = JObject.Parse("{\"kind\":\"write\",\"summary\":\"sheet-created\",\"changed\":true,\"satisfied\":true,\"targetRefs\":[\"Excel:PythonAverage\"]}");
            return ToolResult.Succeed("CreateSheet", "sheet-created", null, observation);
        }
        if (command == "WriteData")
        {
            var written = parameters["data"].ToString(Newtonsoft.Json.Formatting.None);
            PythonWritePayloadValid = written.Contains("1500.0") && written.Contains("2000.0");
            var observation = JObject.Parse("{\"kind\":\"write\",\"summary\":\"python-output-written\",\"changed\":true,\"satisfied\":true,\"targetRefs\":[\"Excel:PythonAverage!A1:B3\"]}");
            return ToolResult.Succeed("WriteData", "python-output-written", null, observation);
        }
        throw new InvalidOperationException("Unexpected host tool: " + command);
    }
}
'@
}
$loop.CaptureContextPack = [AgentRuntimeContextProbe]::CaptureDelegate
if ($null -eq $loop.CaptureContextPack) {
    throw "LoopEngine did not retain the ContextPack capture delegate"
}
$directContext = $loop.CaptureContextPack.Invoke()
if ($directContext.Document.Preview -ne "context-version-1") {
    throw "ContextPack capture delegate is not callable"
}
[AgentRuntimeContextProbe]::Reset()
$loop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:modelCalls += 1
    if ($script:modelCalls -eq 1) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"understanding":"test","steps":[{"step":1,"description":"first","toolHint":"TestAction"},{"step":2,"description":"second","toolHint":"TestAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"test-first","targetRef":"Excel:test/first","effectType":"artifact","operator":"exists","derivedFromCapability":"TestAction","criterionIds":["criterion-raw-user-request"],"required":true},{"id":"test-second","targetRef":"Excel:test/second","effectType":"artifact","operator":"exists","derivedFromCapability":"TestAction","criterionIds":["criterion-raw-user-request"],"required":true}]}}')
    }
    if (-not $prompt.Contains("[adaptive-react]")) {
        throw "Adaptive ReAct prompt marker is missing"
    }
    $expectedContext = "context-version-$($script:modelCalls - 1)"
    if (-not $prompt.Contains($expectedContext)) {
        throw "The ReAct decision did not receive refreshed ContextPack: expected=$expectedContext"
    }
    if ($script:modelCalls -eq 3 -and -not $prompt.Contains("first-observation")) {
        throw "The next ReAct decision did not receive the previous tool observation"
    }
    if ($script:modelCalls -eq 3 -and
        (-not $prompt.Contains("context-version-1") -or
         -not $prompt.Contains("Prior task ContextPacks"))) {
        throw "Switching the active Office object discarded the task's previously observed source context"
    }
    if ($script:modelCalls -eq 4) {
        if (-not $prompt.Contains("second-observation")) {
            throw "The completion decision did not receive the final tool observation"
        }
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"decision":"complete","thought":"the requested state is now fully observed","message":"goal completed after explicit final decision","evidence":["obs-1/e1","obs-2/e1"]}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"act","thought":"choose from current facts","action":{"tool":"TestAction","params":{}}}')
}
$planSession = [ShareRibbon.Agent.AgentSession]::new("run", "Excel", "")
$planSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$planSession.Spec.MutationPolicy = "allow"
$planSession.Spec.Goal = "run"
$planSession.Spec.RequiredTools.Add("TestAction")
$planSkill = [ShareRibbon.Agent.AgentSkill]::new()
$planSkill.RequiredTools.Add("TestAction")
$runTask = [System.Threading.Tasks.Task[ShareRibbon.Agent.AgentResult]]$loop.RunAsync(
    $planSession,
    "system",
    $planSkill)
$runResult = $runTask.GetAwaiter().GetResult()
if (-not $runResult.Success -or $script:modelCalls -ne 4 -or
    $script:adaptiveExecutions -ne 2 -or [AgentRuntimeContextProbe]::Captures -ne 3) {
    throw "Adaptive plan execution did not close the observation loop: calls=$script:modelCalls executions=$script:adaptiveExecutions contexts=$([AgentRuntimeContextProbe]::Captures) result=$($runResult.Message)"
}
if ($runResult.FinalOutput -ne "goal completed after explicit final decision") {
    throw "The explicit completion decision was not delivered as AgentResult.FinalOutput: $($runResult.FinalOutput)"
}
if ($null -eq $planSession.Plan -or $planSession.Plan.Steps.Count -ne 0) {
    throw "Initial projection retained precomputed executable steps instead of leaving next actions to adaptive ReAct"
}

# Exhausting the action/decision budget without an explicit completion control state is not
# permission to run the deterministic Goal verifier. Preserve the final observation and report
# the missing decision truthfully instead of fabricating a completion-projection failure.
$limitLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $registry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:limitModelCalls = 0
$script:adaptiveExecutions = 0
$limitLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:limitModelCalls += 1
    if ($script:limitModelCalls -eq 1) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"understanding":"one action then budget ends","summary":"pending explicit completion","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"limit-state","targetRef":"Excel:test/first","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"act","thought":"perform the only allowed action","action":{"tool":"TestAction","params":{}}}')
}
$limitSession = [ShareRibbon.Agent.AgentSession]::new("one action then budget ends", "Excel", "")
$limitSession.MaxIterations = 1
$limitSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$limitSession.Spec.MutationPolicy = "allow"
$limitSession.Spec.Goal = "one action then budget ends"
$limitSkill = [ShareRibbon.Agent.AgentSkill]::new()
$limitSkill.RequiredTools.Add("TestAction")
$limitResult = $limitLoop.RunAsync($limitSession, "system", $limitSkill).GetAwaiter().GetResult()
if ($limitResult.Success -or
    $script:limitModelCalls -ne 2 -or
    $script:adaptiveExecutions -ne 1 -or
    $limitSession.CurrentIteration -ne 1 -or
    -not $limitResult.Message.Contains("decision=complete")) {
    throw "Decision-budget exhaustion still attempted implicit completion: calls=$script:limitModelCalls actions=$script:adaptiveExecutions result=$($limitResult.Message)"
}

# Cancellation is part of the runtime contract, not a UI-only state change.  A model
# request that never returns must release the adaptive loop promptly when the request
# token is cancelled, without waiting for the provider timeout.
$cancelLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $registry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$cancelModelWait = [System.Threading.Tasks.TaskCompletionSource[string]]::new(
    [System.Threading.Tasks.TaskCreationOptions]::RunContinuationsAsynchronously)
$cancelLoop.SendAIRequest = {
    param($prompt, $system, $history)
    return $cancelModelWait.Task
}
$cancelSession = [ShareRibbon.Agent.AgentSession]::new("cancel a waiting model", "Excel", "")
$cancelSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$cancelSession.Spec.Goal = "cancel a waiting model"
$cancelSource = [System.Threading.CancellationTokenSource]::new()
$cancelWatch = [System.Diagnostics.Stopwatch]::StartNew()
$cancelTask = $cancelLoop.RunAsync($cancelSession, "system", $null, $cancelSource.Token)
[System.Threading.Thread]::Sleep(50)
$cancelSource.Cancel()
$cancelResult = $cancelTask.GetAwaiter().GetResult()
$cancelWatch.Stop()
if ($cancelResult.Success -or
    $cancelResult.ErrorCode -ne [ShareRibbon.ExceptionClassifier]::CodeCancelled -or
    $cancelWatch.ElapsedMilliseconds -gt 1500) {
    throw "Adaptive model wait did not cancel promptly: success=$($cancelResult.Success) code=$($cancelResult.ErrorCode) elapsed=$($cancelWatch.ElapsedMilliseconds)ms"
}
$cancelSource.Dispose()

# The model owns semantic completion. The harness owns the low-level verification projection
# and can bind the already-recorded satisfied host observation when a provider returns a terse
# complete decision without repeating evidence ids or transcribing an OutcomeContract.
$autoCloseRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$autoCloseTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$autoCloseTool.Id = "AutoCloseAction"
$autoCloseTool.Name = "Create verified artifact"
$autoCloseTool.AppType = "excel"
$autoCloseTool.RiskLevel = "safe"
$autoCloseTool.AccessMode = "write"
$autoCloseTool.OutcomeEffects.Add("artifact")
$autoCloseRegistry.RegisterTool($autoCloseTool)
$script:autoCloseModelCalls = 0
$script:autoCloseHostExecutions = 0
$autoCloseRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:autoCloseHostExecutions += 1
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"artifact-created","changed":true,"satisfied":true,"targetRefs":["Excel:auto-close-artifact"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed("AutoCloseAction", "artifact-created", $null, $observation)
}
$autoCloseLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $autoCloseRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$autoCloseLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:autoCloseModelCalls += 1
    if (-not $prompt.Contains("[adaptive-react]")) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"understanding":"create an artifact","summary":"artifact created","capabilityGap":""}')
    }
    if ($script:autoCloseHostExecutions -eq 0) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"decision":"act","thought":"create the artifact","action":{"tool":"AutoCloseAction","params":{}}}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"complete","thought":"Office already contains the requested artifact","message":"done"}')
}
$autoCloseSession = [ShareRibbon.Agent.AgentSession]::new("create an artifact", "Excel", "")
$autoCloseSession.MaxIterations = 3
$autoCloseSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$autoCloseSession.Spec.MutationPolicy = "allow"
$autoCloseSession.Spec.Goal = "create an artifact"
$autoCloseSkill = [ShareRibbon.Agent.AgentSkill]::new()
$autoCloseSkill.RequiredTools.Add("AutoCloseAction")
$autoCloseResult = $autoCloseLoop.RunAsync(
    $autoCloseSession,
    "system",
    $autoCloseSkill).GetAwaiter().GetResult()
if (-not $autoCloseResult.Success -or
    $script:autoCloseHostExecutions -ne 1 -or
    $script:autoCloseModelCalls -ne 3 -or
    $autoCloseSession.CurrentIteration -ne 1) {
    throw "Terse complete did not close from observed host state: modelCalls=$script:autoCloseModelCalls hostExecutions=$script:autoCloseHostExecutions iterations=$($autoCloseSession.CurrentIteration) result=$($autoCloseResult.Message)"
}

# Atomic host verification stores a named property as an object so the property identity is
# not lost. Completion projection must compare the selected property value, not the wrapper
# object, or a passed ValueHash assertion can never satisfy its own generated requirement.
$wrappedValueRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$wrappedValueTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$wrappedValueTool.Id = "WrappedValueWrite"
$wrappedValueTool.Name = "Write a verified value"
$wrappedValueTool.AppType = "excel"
$wrappedValueTool.RiskLevel = "safe"
$wrappedValueTool.AccessMode = "write"
$wrappedValueTool.OutcomeEffects.Add("data_state")
$wrappedValueRegistry.RegisterTool($wrappedValueTool)
$script:wrappedValueModelCalls = 0
$script:wrappedValueHostExecutions = 0
$wrappedValueRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:wrappedValueHostExecutions += 1
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"office_operation_batch","writeExpected":true,"changed":true,"satisfied":true,"verification":[{"id":"value-hash","required":true,"status":"passed","targetRef":"Excel:workbooks/active/worksheets/Sales/ranges/T1","effectType":"data_state","property":"ValueHash","expected":"hash-1","actual":"hash-1"}]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed(
        "WrappedValueWrite",
        "value hash verified",
        $null,
        $observation)
}
$wrappedValueLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $wrappedValueRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$wrappedValueLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:wrappedValueModelCalls += 1
    if (-not $prompt.Contains("[adaptive-react]")) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"understanding":"write a verified value","summary":"value written","capabilityGap":""}')
    }
    if ($script:wrappedValueHostExecutions -eq 0) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"decision":"act","thought":"write once","action":{"tool":"WrappedValueWrite","params":{"targetRange":"Sales!T1","data":[["margin"]]}}}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"complete","thought":"the host verified the value hash","message":"done"}')
}
$wrappedValueSession = [ShareRibbon.Agent.AgentSession]::new("write a verified value", "Excel", "")
$wrappedValueSession.MaxIterations = 3
$wrappedValueSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$wrappedValueSession.Spec.MutationPolicy = "allow"
$wrappedValueSession.Spec.Goal = "write a verified value"
$wrappedValueSkill = [ShareRibbon.Agent.AgentSkill]::new()
$wrappedValueSkill.RequiredTools.Add("WrappedValueWrite")
$wrappedValueResult = $wrappedValueLoop.RunAsync(
    $wrappedValueSession,
    "system",
    $wrappedValueSkill).GetAwaiter().GetResult()
if (-not $wrappedValueResult.Success -or
    $script:wrappedValueHostExecutions -ne 1 -or
    $script:wrappedValueModelCalls -ne 3) {
    throw "Wrapped atomic verification did not complete: calls=$script:wrappedValueModelCalls hostExecutions=$script:wrappedValueHostExecutions result=$($wrappedValueResult.Message)"
}

# A provider may ignore a successful observation and emit the exact same write action again.
# ReAct must still return the first observation to the model, but it must not dispatch an
# identical mutation twice while the Office world revision is unchanged.
$repeatGuardRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$repeatGuardTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$repeatGuardTool.Id = "RepeatGuardAction"
$repeatGuardTool.Name = "Create one verified artifact"
$repeatGuardTool.AppType = "excel"
$repeatGuardTool.RiskLevel = "safe"
$repeatGuardTool.AccessMode = "write"
$repeatGuardTool.OutcomeEffects.Add("artifact")
$repeatGuardRegistry.RegisterTool($repeatGuardTool)
$script:repeatGuardModelCalls = 0
$script:repeatGuardHostExecutions = 0
$repeatGuardRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:repeatGuardHostExecutions += 1
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"one-artifact-created","changed":true,"satisfied":true,"targetRefs":["Excel:repeat-guard-artifact"],"verification":[{"id":"artifact-request","required":true,"status":"passed","targetRef":"Excel:repeat-guard-artifact","effectType":"artifact","requestProperty":"Name","requestExpected":"OnlyOnce","actual":"OnlyOnce"}]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed(
        "RepeatGuardAction",
        "one-artifact-created",
        $null,
        $observation)
}
$repeatGuardLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $repeatGuardRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$repeatGuardLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:repeatGuardModelCalls += 1
    switch ($script:repeatGuardModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"create one artifact","summary":"created","capabilityGap":""}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"create it","action":{"tool":"RepeatGuardAction","params":{"name":"OnlyOnce"}}}') }
        3 {
            if (-not $prompt.Contains("one-artifact-created")) {
                throw "Successful host observation was not returned before the repeated action"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"repeat despite success","action":{"tool":"RepeatGuardAction","params":{"name":"OnlyOnce"}}}')
        }
        default {
            if (-not $prompt.Contains("duplicate_success_guard")) {
                throw "The model did not receive the duplicate-success guard observation"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"the first verified mutation already satisfied the request","message":"created once","evidence":["obs-1/e1"]}')
        }
    }
}
$repeatGuardSession = [ShareRibbon.Agent.AgentSession]::new("create one artifact", "Excel", "")
$repeatGuardSession.MaxIterations = 4
$repeatGuardSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$repeatGuardSession.Spec.MutationPolicy = "allow"
$repeatGuardSession.Spec.Goal = "create one artifact"
$repeatGuardSkill = [ShareRibbon.Agent.AgentSkill]::new()
$repeatGuardSkill.RequiredTools.Add("RepeatGuardAction")
$repeatGuardResult = $repeatGuardLoop.RunAsync(
    $repeatGuardSession,
    "system",
    $repeatGuardSkill).GetAwaiter().GetResult()
if (-not $repeatGuardResult.Success -or
    $script:repeatGuardHostExecutions -ne 1 -or
    $script:repeatGuardModelCalls -ne 4 -or
    $repeatGuardSession.CurrentIteration -ne 2) {
    throw "Identical successful mutation was dispatched more than once: modelCalls=$script:repeatGuardModelCalls hostExecutions=$script:repeatGuardHostExecutions iterations=$($repeatGuardSession.CurrentIteration) result=$($repeatGuardResult.Message)"
}

# After the final tool observation returns to the model, the provisional Goal mapping is
# grounded in exact host evidence. A direct producer binding may remain when the registered
# tool really emits that effect; unsupported lineage is still removed before sealing.
$lineageRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$lineageCreateTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$lineageCreateTool.Id = "CreateSheet"
$lineageCreateTool.Name = "Create worksheet"
$lineageCreateTool.AppType = "excel"
$lineageCreateTool.RiskLevel = "safe"
$lineageCreateTool.AccessMode = "write"
$lineageCreateTool.OutcomeEffects.Add("object_exists")
$lineageRegistry.RegisterTool($lineageCreateTool)
$script:lineageModelCalls = 0
$script:lineageHostExecutions = 0
$script:lineagePlanEvents = 0
$lineageRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:lineageHostExecutions += 1
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"worksheet-created","changed":true,"satisfied":true,"targetRefs":["Excel:workbooks/active/worksheets/NormalizedSheet"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed("CreateSheet", "worksheet-created", $null, $observation)
}
$lineageLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $lineageRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$lineageLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:lineageModelCalls += 1
    switch ($script:lineageModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"create a worksheet","steps":[{"step":1,"description":"create destination","toolHint":"CreateSheet"}],"summary":"created","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"destination-created","targetRef":"Excel:workbooks/active/worksheets/NormalizedSheet","effectType":"object_exists","operator":"exists","derivedFromCapability":"CreateSheet","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"create the requested worksheet","action":{"tool":"CreateSheet","params":{"name":"NormalizedSheet"}}}') }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"the worksheet exists","message":"done","evidence":["obs-1/e1"]}') }
    }
}
$lineageLoop.OnPlanGenerated = {
    param($plan)
    $script:lineagePlanEvents += 1
}
$lineageSession = [ShareRibbon.Agent.AgentSession]::new("create a destination worksheet", "Excel", "")
$lineageSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$lineageSession.Spec.MutationPolicy = "allow"
$lineageSession.Spec.Goal = "create a destination worksheet"
$lineageSkill = [ShareRibbon.Agent.AgentSkill]::new()
$lineageSkill.RequiredTools.Add("CreateSheet")
$lineageResult = $lineageLoop.RunAsync($lineageSession, "system", $lineageSkill).GetAwaiter().GetResult()
$lineageFrozenRequirement = $lineageSession.Spec.OutcomeContract.Requirements[0]
if (-not $lineageResult.Success -or
    $script:lineageModelCalls -ne 3 -or
    $script:lineageHostExecutions -ne 1 -or
    $script:lineagePlanEvents -ne 1 -or
    $lineageSession.CurrentIteration -ne 1 -or
    -not $lineageSession.Spec.OutcomeContract.Frozen -or
    $lineageFrozenRequirement.DerivedFromCapability -ne "CreateSheet") {
    throw "Explicit completion projection failed: calls=$script:lineageModelCalls executions=$script:lineageHostExecutions planEvents=$script:lineagePlanEvents result=$($lineageResult.Message)"
}

# A genuinely invalid provisional projection must not block the first adaptive action.
# The completion decision replaces it after real host evidence exists.
$repairRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$repairTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$repairTool.Id = "RepairAction"
$repairTool.Name = "Repair action"
$repairTool.AppType = "excel"
$repairTool.RiskLevel = "safe"
$repairTool.AccessMode = "write"
$repairTool.OutcomeEffects.Add("artifact")
$repairRegistry.RegisterTool($repairTool)
$script:initialRepairModelCalls = 0
$script:initialRepairExecutions = 0
$script:initialRepairPlanEvents = 0
$repairRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:initialRepairExecutions += 1
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"artifact-created","changed":true,"satisfied":true,"targetRefs":["Excel:repair-artifact"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed("RepairAction", "artifact-created", $null, $observation)
}
$repairLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $repairRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$repairLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:initialRepairModelCalls += 1
    switch ($script:initialRepairModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"produce an artifact","steps":[{"step":1,"description":"produce artifact","toolHint":"RepairAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"artifact-ready","targetRef":"Excel:repair-artifact","effectType":"artifact","operator":"unsupported","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"produce the artifact","action":{"tool":"RepairAction","params":{}}}') }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"the artifact is verified","message":"done","evidence":["obs-1/e1"],"outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"artifact-ready","targetRef":"Excel:repair-artifact","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
    }
}
$repairLoop.OnPlanGenerated = {
    param($plan)
    $script:initialRepairPlanEvents += 1
}
$repairSession = [ShareRibbon.Agent.AgentSession]::new("produce a verified artifact", "Excel", "")
$repairSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$repairSession.Spec.MutationPolicy = "allow"
$repairSession.Spec.Goal = "produce a verified artifact"
$repairSkill = [ShareRibbon.Agent.AgentSkill]::new()
$repairSkill.RequiredTools.Add("RepairAction")
$repairResult = $repairLoop.RunAsync($repairSession, "system", $repairSkill).GetAwaiter().GetResult()
if (-not $repairResult.Success -or
    $script:initialRepairModelCalls -ne 3 -or
    $script:initialRepairExecutions -ne 1 -or
    $script:initialRepairPlanEvents -ne 1 -or
    $repairSession.CurrentIteration -ne 1 -or
    -not $repairSession.Spec.OutcomeContract.Frozen -or
    $repairSession.Spec.OutcomeContract.Requirements[0].Operator -ne "exists") {
    throw "Initial planning repair loop failed: calls=$script:initialRepairModelCalls executions=$script:initialRepairExecutions planEvents=$script:initialRepairPlanEvents result=$($repairResult.Message)"
}

# A frozen resume contract is authoritative persisted state.  It must be integrity-checked
# before any model call, and a valid persisted plan must not be replaced by a fresh proposal.
$resumeSession = New-TestGoalSession $twoCriterionGoal
$resumePlan = New-TwoCriterionGoalPlan
$resumePlan.Summary = "persisted-plan-marker"
$resumeStep = [ShareRibbon.Agent.PlanStep]::new()
$resumeStep.StepNumber = 1
$resumeStep.Description = "persisted completed hint"
$resumeStep.Status = [ShareRibbon.Agent.StepStatus]::Completed
$resumePlan.Steps.Add($resumeStep)
$resumeFreezeError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($resumeSession, $resumePlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($resumeFreezeError)) {
    throw "Valid resume fixture could not be frozen: $resumeFreezeError"
}
$resumeSession.Plan = $resumePlan
$resumeContract = $resumeSession.Spec.OutcomeContract
$resumeContractHash = $resumeContract.FrozenOutcomeContractHash
$resumeSheetEvidence = New-TestOutcomeEvidence `
    -EvidenceId "resume-sheet/e1" `
    -IterationEvidenceId "resume-sheet" `
    -TargetRef "Excel:GoalOutput" `
    -EffectType "object_exists" `
    -SourceToolId "CreateSheet"
$resumeWriteEvidence = New-TestOutcomeEvidence `
    -EvidenceId "resume-write/e1" `
    -IterationEvidenceId "resume-write" `
    -TargetRef "Excel:GoalOutput!A1:B2" `
    -EffectType "data_state" `
    -Expected $resumeContract.Requirements[1].ExpectedValue `
    -SourceToolId "WriteData"
Add-TestOutcomeIteration $resumeSession "resume-sheet" "CreateSheet" @($resumeSheetEvidence)
Add-TestOutcomeIteration $resumeSession "resume-write" "WriteData" @($resumeWriteEvidence)
$resumeLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $criterionRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:resumePlanningCalls = 0
$script:resumeReactCalls = 0
$script:resumeHostExecutions = 0
$script:resumePlanEvents = 0
$script:resumePlanEventValue = $null
$criterionRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:resumeHostExecutions += 1
    return [ShareRibbon.Agent.ToolResult]::Succeed("unexpected", "unexpected")
}
$resumeLoop.SendAIRequest = {
    param($prompt, $system, $history)
    if (-not $prompt.Contains("[adaptive-react]")) {
        $script:resumePlanningCalls += 1
        throw "A valid frozen resume unexpectedly called the planning model"
    }
    $script:resumeReactCalls += 1
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"complete","thought":"persisted evidence already proves the goal","message":"resumed","evidence":["resume-sheet/e1","resume-write/e1"]}')
}
$resumeLoop.OnPlanGenerated = {
    param($plan)
    $script:resumePlanEvents += 1
    $script:resumePlanEventValue = $plan
}
$resumeResult = $resumeLoop.RunAsync($resumeSession, "system", $null).GetAwaiter().GetResult()
if (-not $resumeResult.Success -or
    $script:resumePlanningCalls -ne 0 -or
    $script:resumeReactCalls -ne 1 -or
    $script:resumeHostExecutions -ne 0 -or
    $script:resumePlanEvents -ne 1 -or
    -not [object]::ReferenceEquals($resumeSession.Plan, $resumePlan) -or
    -not [object]::ReferenceEquals($resumeSession.Spec.OutcomeContract, $resumeContract) -or
    -not [object]::ReferenceEquals($script:resumePlanEventValue, $resumePlan) -or
    $resumeSession.Plan.Summary -ne "persisted-plan-marker" -or
    $resumeSession.Spec.OutcomeContract.FrozenOutcomeContractHash -ne $resumeContractHash) {
    throw "A valid frozen resume was replanned or mutated before adaptive execution: planning=$script:resumePlanningCalls react=$script:resumeReactCalls result=$($resumeResult.Message)"
}

# A tampered frozen contract must fail before planning, ReAct, plan publication, or host IO.
$tamperedSession = New-TestGoalSession $twoCriterionGoal
$tamperedPlan = New-TwoCriterionGoalPlan
$tamperedPlan.Summary = "tampered-plan-marker"
$tamperedPlan.Steps.Add([ShareRibbon.Agent.PlanStep]@{
    StepNumber = 1
    Description = "persisted hint"
    Status = [ShareRibbon.Agent.StepStatus]::Pending
})
$tamperedFreezeError = [string]$freezeMethod.Invoke(
    $criterionLoop,
    @($tamperedSession, $tamperedPlan, "Excel"))
if (-not [string]::IsNullOrWhiteSpace($tamperedFreezeError)) {
    throw "Tampered resume fixture could not first be frozen: $tamperedFreezeError"
}
$tamperedSession.Plan = $tamperedPlan
$tamperedContract = $tamperedSession.Spec.OutcomeContract
$tamperedContract.Requirements[0].TargetRef = "Excel:TamperedTarget"
$tamperedLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $criterionRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:tamperedModelCalls = 0
$script:tamperedHostExecutions = 0
$script:tamperedPlanEvents = 0
$criterionRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:tamperedHostExecutions += 1
    return [ShareRibbon.Agent.ToolResult]::Succeed("unexpected", "unexpected")
}
$tamperedLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:tamperedModelCalls += 1
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"understanding":"replacement","steps":[],"summary":"replacement","outcomeContract":{"requirements":[]}}')
}
$tamperedLoop.OnPlanGenerated = {
    param($plan)
    $script:tamperedPlanEvents += 1
}
$tamperedResult = $tamperedLoop.RunAsync($tamperedSession, "system", $null).GetAwaiter().GetResult()
if ($tamperedResult.Success -or
    -not $tamperedResult.Message.Contains("[OUTCOME_CONTRACT_MUTATED]") -or
    $script:tamperedModelCalls -ne 0 -or
    $script:tamperedHostExecutions -ne 0 -or
    $script:tamperedPlanEvents -ne 0 -or
    $tamperedSession.Iterations.Count -ne 0 -or
    -not [object]::ReferenceEquals($tamperedSession.Plan, $tamperedPlan) -or
    -not [object]::ReferenceEquals($tamperedSession.Spec.OutcomeContract, $tamperedContract)) {
    throw "A tampered frozen resume was not rejected before all model/host activity: model=$script:tamperedModelCalls host=$script:tamperedHostExecutions result=$($tamperedResult.Message)"
}

# The model cannot end a run merely by saying "complete". Deterministic acceptance must
# reject premature completion and feed the missing contract back as the next observation.
$acceptRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$acceptTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$acceptTool.Id = "RequiredAction"
$acceptTool.Name = "Required action"
$acceptTool.AppType = "excel"
$acceptTool.RiskLevel = "safe"
$acceptTool.AccessMode = "write"
$acceptTool.OutcomeEffects.Add("artifact")
$acceptRegistry.RegisterTool($acceptTool)
$supportTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$supportTool.Id = "SupportAction"
$supportTool.Name = "Support action"
$supportTool.AppType = "excel"
$supportTool.RiskLevel = "safe"
$supportTool.AccessMode = "write"
$supportTool.OutcomeEffects.Add("artifact")
$acceptRegistry.RegisterTool($supportTool)
$script:acceptExecutions = 0
$acceptRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:acceptExecutions += 1
    $commandObject = [Newtonsoft.Json.Linq.JObject]::Parse($code)
    $executedTool = $commandObject["command"].ToString()
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"required-action-observed","changed":true,"satisfied":true,"targetRefs":["Excel:test"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed($executedTool, "$executedTool-observed", $null, $observation)
}
$acceptLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $acceptRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:acceptModelCalls = 0
$acceptLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:acceptModelCalls += 1
    switch ($script:acceptModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"acceptance","steps":[{"step":1,"description":"perform required action","toolHint":"RequiredAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"accepted-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"too early","message":"done","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"accepted-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        3 {
            if (-not $prompt.Contains("Completion verification projection was rejected")) {
                throw "Premature completion rejection was not fed back into ReAct"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"supporting action first","action":{"tool":"SupportAction","params":{}}}')
        }
        4 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"support action is insufficient","message":"done","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"accepted-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        5 {
            if (-not $prompt.Contains("Completion verification projection was rejected")) {
                throw "An unrelated successful tool incorrectly completed the plan milestone"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"close contract gap","action":{"tool":"RequiredAction","params":{}}}')
        }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"accepted evidence exists","message":"done","evidence":["obs-1/e1","obs-2/e1"],"outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"accepted-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
    }
}
$acceptSession = [ShareRibbon.Agent.AgentSession]::new("accept", "Excel", "")
$acceptSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$acceptSession.Spec.MutationPolicy = "allow"
$acceptSession.Spec.Goal = "accept"
$acceptSession.Spec.RequiredCapabilities.Add("RequiredAction")
$acceptSkill = [ShareRibbon.Agent.AgentSkill]::new()
$acceptSkill.RequiredTools.Add("RequiredAction")
$acceptSkill.RequiredTools.Add("SupportAction")
$acceptResult = $acceptLoop.RunAsync($acceptSession, "system", $acceptSkill).GetAwaiter().GetResult()
if (-not $acceptResult.Success -or $script:acceptExecutions -ne 2 -or $script:acceptModelCalls -ne 6) {
    throw "Deterministic completion gate failed: calls=$script:acceptModelCalls executions=$script:acceptExecutions result=$($acceptResult.Message)"
}

# Strategy reset is a normal decision after a successful observation. It must return to the
# same adaptive loop without generating a replacement future-step skeleton.
$replanRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$replanTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$replanTool.Id = "ReplanAction"
$replanTool.Name = "Replan action"
$replanTool.AppType = "excel"
$replanTool.RiskLevel = "safe"
$replanTool.AccessMode = "write"
$replanTool.OutcomeEffects.Add("artifact")
$replanRegistry.RegisterTool($replanTool)
$script:replanExecutions = 0
$replanRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:replanExecutions += 1
    $summary = "replan-observation-$script:replanExecutions"
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"changed","changed":true,"satisfied":true,"targetRefs":["Excel:test"]}')
    $observation["summary"] = [Newtonsoft.Json.Linq.JValue]::CreateString($summary)
    return [ShareRibbon.Agent.ToolResult]::Succeed("ReplanAction", $summary, $null, $observation)
}
$replanMemory = [ShareRibbon.Agent.AgentMemory]::new()
$initialReplanContext = [ShareRibbon.Agent.Context.ContextPack]::new()
$initialReplanContext.AppType = "Excel"
$initialReplanContext.Document.Preview = "replan-context-0"
$replanMemory.SetWorking("lastContextPack", $initialReplanContext)
$replanLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $replanRegistry,
    $replanMemory,
    $promptManager)
$script:replanModelCalls = 0
$script:replanCaptures = 0
$script:replanDecisionSawObservation = $false
$script:replacementPlanningPrompt = $null
$script:replacementSawObservation = $false
$script:replacementSawCurrentWorld = $false
$replanLoop.CaptureContextPack = [System.Func[ShareRibbon.Agent.Context.ContextPack]] {
    $script:replanCaptures += 1
    $pack = [ShareRibbon.Agent.Context.ContextPack]::new()
    $pack.AppType = "Excel"
    $pack.Document.Preview = "replan-context-$script:replanCaptures"
    return $pack
}
$replanLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:replanModelCalls += 1
    switch ($script:replanModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"initial","steps":[{"step":1,"description":"initial action","toolHint":"ReplanAction"}],"summary":"initial","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"replan-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"first action","action":{"tool":"ReplanAction","params":{}}}') }
        3 {
            $script:replanDecisionSawObservation = $prompt.Contains("replan-observation-1")
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"replan","thought":"new facts require a revised skeleton","message":"revise"}')
        }
        4 {
            $script:replacementPlanningPrompt = $prompt
            $firstObservation = $prompt.IndexOf("replan-observation-1")
            $worldValue = $prompt.IndexOf("replan-context-")
            $script:replacementSawObservation = $firstObservation -ge 0
            $script:replacementSawCurrentWorld = $worldValue -ge 0
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"act after strategy reset","action":{"tool":"ReplanAction","params":{}}}')
        }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"all observations accepted","message":"done","evidence":["obs-1/e1","obs-2/e1"],"outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"replan-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
    }
}
$replanSession = [ShareRibbon.Agent.AgentSession]::new("replan", "Excel", "")
$replanSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$replanSession.Spec.MutationPolicy = "allow"
$replanSession.Spec.Goal = "replan"
$replanSkill = [ShareRibbon.Agent.AgentSkill]::new()
$replanSkill.RequiredTools.Add("ReplanAction")
$replanResult = $replanLoop.RunAsync($replanSession, "system", $replanSkill).GetAwaiter().GetResult()
if (-not $replanResult.Success -or $script:replanExecutions -ne 2 -or
    $script:replanModelCalls -ne 5 -or $script:replanCaptures -lt 3) {
    throw "Success-triggered replan did not execute adaptively: calls=$script:replanModelCalls executions=$script:replanExecutions result=$($replanResult.Message)"
}
if ($null -eq $script:replacementPlanningPrompt) {
    throw "Post-reset adaptive prompt was not captured"
}
if (-not $script:replanDecisionSawObservation -or
    -not $script:replacementSawObservation -or
    -not $script:replacementSawCurrentWorld) {
    throw "Strategy reset lost its triggering observation or current world snapshot: decision=$script:replanDecisionSawObservation observation=$script:replacementSawObservation world=$script:replacementSawCurrentWorld"
}

# A retryable tool failure is still an ordinary observation for the one adaptive loop.
# There is no hidden repair request; the next standard ReAct decision may select a
# different implementation, and the soft plan hint cannot block outcome completion.
$failureRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($toolId in @("FailingAction", "AlternativeAction")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $toolId
    $descriptor.Name = $toolId
    $descriptor.AppType = "excel"
    $descriptor.RiskLevel = "safe"
    $descriptor.AccessMode = "write"
    $descriptor.OutcomeEffects.Add("artifact")
    $failureRegistry.RegisterTool($descriptor)
}
$script:failureExecutions = 0
$failureRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:failureExecutions += 1
    $command = [Newtonsoft.Json.Linq.JObject]::Parse($code)
    if ($command["command"].ToString() -eq "FailingAction") {
        $failedObservation = [Newtonsoft.Json.Linq.JObject]::Parse(
            '{"kind":"write","summary":"first implementation unavailable","changed":false,"satisfied":false}')
        $retryableFailure = [ShareRibbon.Agent.ToolResult]::Failed(
            "FailingAction",
            "first implementation unavailable",
            $null,
            "NO_RETRY_CALL",
            "first implementation unavailable",
            "deterministic failure",
            $true,
            $failedObservation,
            $null)
        $retryableFailure.Retryable = $true
        return $retryableFailure
    }

    $successObservation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"goal satisfied by alternative","changed":true,"satisfied":true,"targetRefs":["Excel:test"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed(
        "AlternativeAction",
        "goal satisfied by alternative",
        $null,
        $successObservation,
        "",
        $null)
}
$failureLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $failureRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:failureModelCalls = 0
$script:failureWasObserved = $false
$failureLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:failureModelCalls += 1
    switch ($script:failureModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"adaptive failure","steps":[{"step":1,"description":"satisfy goal","toolHint":"FailingAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"alternative-state","targetRef":"Excel:test","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"try primary implementation","action":{"tool":"FailingAction","params":{}}}') }
        3 {
            $script:failureWasObserved = $prompt.Contains("NO_RETRY_CALL") -and $prompt.Contains("[adaptive-react]")
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"use a different implementation","action":{"tool":"AlternativeAction","params":{}}}')
        }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"the world state now satisfies the goal","message":"done","evidence":["obs-2/e1"]}') }
    }
}
$failureSession = [ShareRibbon.Agent.AgentSession]::new("adaptive failure", "Excel", "")
$failureSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$failureSession.Spec.MutationPolicy = "allow"
$failureSession.Spec.Goal = "satisfy the requested world state"
$failureSkill = [ShareRibbon.Agent.AgentSkill]::new()
$failureSkill.RequiredTools.Add("FailingAction")
$failureSkill.RequiredTools.Add("AlternativeAction")
$failureResult = $failureLoop.RunAsync($failureSession, "system", $failureSkill).GetAwaiter().GetResult()
if (-not $failureResult.Success -or
    $script:failureExecutions -ne 2 -or
    $script:failureModelCalls -ne 4 -or
    -not $script:failureWasObserved) {
    throw "Tool failure did not return to the single adaptive loop: calls=$script:failureModelCalls executions=$script:failureExecutions result=$($failureResult.Message)"
}

# Transient failures may repeat an identical call only for a non-mutating tool contract.
# Retryability is inferred from both the stable error code and the descriptor access mode.
$safeRetryRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$safeRetryDescriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
$safeRetryDescriptor.Id = "SafeRetryRead"
$safeRetryDescriptor.Name = "Safe retry read"
$safeRetryDescriptor.AppType = "excel"
$safeRetryDescriptor.RiskLevel = "safe"
$safeRetryDescriptor.AccessMode = "read"
$safeRetryRegistry.RegisterTool($safeRetryDescriptor)
$script:safeRetryExecutions = 0
$safeRetryRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:safeRetryExecutions += 1
    if ($script:safeRetryExecutions -eq 1) {
        $failedObservation = [Newtonsoft.Json.Linq.JObject]::Parse(
            '{"kind":"read","summary":"transient read failure","changed":false,"satisfied":false}')
        return [ShareRibbon.Agent.ToolResult]::Failed(
            "SafeRetryRead",
            "transient read failure",
            $null,
            "NETWORK_ERROR",
            "transient read failure",
            "network unavailable",
            $true,
            $failedObservation)
    }
    $successObservation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"read","summary":"safe read complete","changed":false,"satisfied":true,"targetRefs":["Excel:SafeRead!A1"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed(
        "SafeRetryRead",
        "safe read complete",
        @("value"),
        $successObservation)
}
$safeRetryLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $safeRetryRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:safeRetryModelCalls = 0
$safeRetryLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:safeRetryModelCalls += 1
    switch ($script:safeRetryModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"safe retry","steps":[{"step":1,"description":"read","toolHint":"SafeRetryRead"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"read-complete","targetRef":"Excel:SafeRead!A1","effectType":"read_coverage","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"read","action":{"tool":"SafeRetryRead","params":{"options":{"b":2,"a":1}}}}') }
        3 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"retry transient read","action":{"tool":"SafeRetryRead","params":{"options":{"a":1,"b":2}}}}') }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"read observation satisfies goal","message":"The exact read result is value.","evidence":["obs-2/e1"]}') }
    }
}
$safeRetrySession = [ShareRibbon.Agent.AgentSession]::new("safe retry", "Excel", "")
$safeRetrySession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$safeRetrySession.Spec.Goal = "read with bounded transient retry"
$safeRetrySession.Spec.MutationPolicy = "read_only"
$safeRetrySkill = [ShareRibbon.Agent.AgentSkill]::new()
$safeRetrySkill.RequiredTools.Add("SafeRetryRead")
$safeRetryResult = $safeRetryLoop.RunAsync($safeRetrySession, "system", $safeRetrySkill).GetAwaiter().GetResult()
if (-not $safeRetryResult.Success -or
    $script:safeRetryModelCalls -ne 4 -or
    $script:safeRetryExecutions -ne 2 -or
    $safeRetrySession.Iterations.Count -lt 2 -or
    -not $safeRetrySession.Iterations[0].Observation.Contains("retryable=True") -or
    $safeRetryResult.FinalOutput -ne "The exact read result is value.") {
    throw "Safe transient retry policy failed: executions=$script:safeRetryExecutions result=$($safeRetryResult.Message)"
}

# Exercise the complete read -> compute -> create -> write workflow through the adaptive
# loop. Runtime data, not preview literals generated during planning, must cross each seam.
$pythonLoopRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($toolId in @("ReadRange", "PythonCompute", "CreateSheet", "WriteData")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $toolId
    $descriptor.Name = $toolId
    $descriptor.AppType = "excel"
    $descriptor.RiskLevel = "safe"
    $descriptor.AccessMode = if ($toolId -eq "ReadRange") { "read" } elseif ($toolId -eq "PythonCompute") { "compute" } else { "write" }
    $pythonLoopRegistry.RegisterTool($descriptor)
}
[AgentRuntimeContextProbe]::Reset()
$pythonLoopRegistry.ExecuteCodeWithToolResult = [AgentRuntimeContextProbe]::PythonHostDelegate
$pythonLoopMemory = [ShareRibbon.Agent.AgentMemory]::new()
$pythonLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $pythonLoopRegistry,
    $pythonLoopMemory,
    $promptManager)
$pythonLoop.SendAIRequest = [AgentRuntimeContextProbe]::PythonModelDelegate
$pythonLoopGoal = New-TestGoalContract `
    -RawRequest "create a destination worksheet; use Python to calculate regional averages and write them there" `
    -ClauseTexts @("create a destination worksheet", "use Python to calculate regional averages and write them there") `
    -CriterionIds @("goal-python-sheet", "goal-python-result") `
    -CriterionKinds @("state", "state")
if (-not $pythonLoopGoal.RequiredCapabilities.Contains("PythonCompute")) {
    throw "Python workflow GoalContract did not preserve the explicit compute capability"
}
$pythonLoopSession = New-TestGoalSession $pythonLoopGoal
$pythonLoopSkill = [ShareRibbon.Agent.AgentSkill]::new()
foreach ($toolId in @("ReadRange", "PythonCompute", "CreateSheet", "WriteData")) {
    $pythonLoopSession.Spec.RequiredTools.Add($toolId)
    $pythonLoopSkill.RequiredTools.Add($toolId)
}
$pythonLoopResult = $pythonLoop.RunAsync($pythonLoopSession, "system", $pythonLoopSkill).GetAwaiter().GetResult()
if (-not $pythonLoopResult.Success -or $pythonLoopSession.CurrentIteration -ne 4 -or
    [AgentRuntimeContextProbe]::PythonHostExecutions -ne 3 -or
    [AgentRuntimeContextProbe]::PythonModelCalls -ne 6 -or
    -not [AgentRuntimeContextProbe]::PythonWritePayloadValid -or
    $pythonLoopSession.Spec.OutcomeContract.ValidatedComputeCapabilities.Count -ne 1 -or
    $pythonLoopSession.Spec.OutcomeContract.ValidatedComputeCapabilities[0] -ne "PythonCompute") {
    $pythonTrace = (($pythonLoopSession.Iterations | ForEach-Object { "$($_.Action.ToolId):$($_.Explanation.Success):$($_.Observation)" }) -join " | ")
    throw "Adaptive Python workflow failed: calls=$([AgentRuntimeContextProbe]::PythonModelCalls) actions=$($pythonLoopSession.CurrentIteration) hostExecutions=$([AgentRuntimeContextProbe]::PythonHostExecutions) result=$($pythonLoopResult.Message) observation=$($pythonLoopMemory.GetWorkingString('lastObservation')) trace=$pythonTrace"
}

# A soft plan may suggest one implementation while the adaptive loop chooses another.
# Plan status/toolHint must not gate the action or the outcome-based completion decision.
$coverageRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($id in @("PlannedAction", "AlternativeAction")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $id
    $descriptor.Name = $id
    $descriptor.AppType = "excel"
    $descriptor.RiskLevel = "safe"
    $descriptor.AccessMode = "write"
    $descriptor.OutcomeEffects.Add("artifact")
    $coverageRegistry.RegisterTool($descriptor)
}
$script:plannedExecutions = 0
$script:alternativeExecutions = 0
$coverageRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $command = ([Newtonsoft.Json.Linq.JObject]::Parse($code))["command"].ToString()
    if ($command -eq "PlannedAction") {
        $script:plannedExecutions += 1
    } else {
        $script:alternativeExecutions += 1
    }
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"goal satisfied by runtime alternative","changed":true,"satisfied":true,"targetRefs":["Excel:alternative"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed($command, "runtime alternative observed", $null, $observation)
}
$coverageLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $coverageRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:coverageModelCalls = 0
$coverageLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:coverageModelCalls += 1
    switch ($script:coverageModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"soft plan","steps":[{"step":1,"description":"satisfy goal","toolHint":"PlannedAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"alternative-state","targetRef":"Excel:alternative","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"current facts favor the alternative","action":{"tool":"AlternativeAction","params":{}}}') }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"host observation proves the goal","message":"done","evidence":["obs-1/e1"]}') }
    }
}
$coverageSession = [ShareRibbon.Agent.AgentSession]::new("soft plan", "Excel", "")
$coverageSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$coverageSession.Spec.MutationPolicy = "allow"
$coverageSession.Spec.Goal = "satisfy requested world state"
$coverageSkill = [ShareRibbon.Agent.AgentSkill]::new()
$coverageSkill.RequiredTools.Add("PlannedAction")
$coverageSkill.RequiredTools.Add("AlternativeAction")
$coverageResult = $coverageLoop.RunAsync($coverageSession, "system", $coverageSkill).GetAwaiter().GetResult()
if (-not $coverageResult.Success -or
    $script:plannedExecutions -ne 0 -or
    $script:alternativeExecutions -ne 1 -or
    $script:coverageModelCalls -ne 3) {
    throw "Soft plan still controlled execution/completion: calls=$script:coverageModelCalls planned=$script:plannedExecutions alternative=$script:alternativeExecutions result=$($coverageResult.Message)"
}

# A mutating call must not be blindly repeated even if a faulty producer marks it retryable.
# Canonical signatures must also treat reordered object keys as the same parameter value.
$repairRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($id in @("NonRetryableAction", "FallbackAction")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $id
    $descriptor.Name = $id
    $descriptor.AppType = "excel"
    $descriptor.RiskLevel = "safe"
    $descriptor.AccessMode = "write"
    $descriptor.OutcomeEffects.Add("artifact")
    $repairRegistry.RegisterTool($descriptor)
}
$script:nonRetryableExecutions = 0
$script:fallbackExecutions = 0
$repairRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $commandObject = [Newtonsoft.Json.Linq.JObject]::Parse($code)
    $command = $commandObject["command"].ToString()
    if ($command -eq "FallbackAction") {
        $script:fallbackExecutions += 1
        $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
            '{"kind":"write","summary":"fallback satisfied goal","changed":true,"satisfied":true,"targetRefs":["Excel:fallback"]}')
        return [ShareRibbon.Agent.ToolResult]::Succeed("FallbackAction", "fallback satisfied goal", $null, $observation)
    }
    $script:nonRetryableExecutions += 1
    $failure = [ShareRibbon.Agent.ToolResult]::Failed(
        "NonRetryableAction",
        "do not repeat the same call",
        $null,
        "TEST_NON_RETRYABLE",
        "do not repeat the same call",
        "deterministic failure",
        $false)
    $failure.Retryable = $true
    return $failure
}
$repairLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $repairRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:repairModelCalls = 0
$repairLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:repairModelCalls += 1
    switch ($script:repairModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"adaptive fallback","steps":[{"step":1,"description":"satisfy goal","toolHint":"NonRetryableAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"fallback-complete","targetRef":"Excel:fallback","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"try first implementation","action":{"tool":"NonRetryableAction","params":{"value":1,"options":{"b":2,"a":1}}}}') }
        3 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"attempted duplicate should be guarded","action":{"tool":"NonRetryableAction","params":{"options":{"a":1,"b":2},"value":1}}}') }
        4 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"choose fallback","action":{"tool":"FallbackAction","params":{}}}') }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"fallback observation proves goal","message":"done","evidence":["obs-3/e1"]}') }
    }
}
$repairSession = [ShareRibbon.Agent.AgentSession]::new("adaptive fallback", "Excel", "")
$repairSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$repairSession.Spec.MutationPolicy = "allow"
$repairSession.Spec.Goal = "satisfy goal without duplicate side effects"
$repairSession.Spec.RequiredTools.Add("NonRetryableAction")
$repairSession.Spec.RequiredTools.Add("FallbackAction")
$repairSkill = [ShareRibbon.Agent.AgentSkill]::new()
$repairSkill.RequiredTools.Add("NonRetryableAction")
$repairSkill.RequiredTools.Add("FallbackAction")
$repairResult = $repairLoop.RunAsync($repairSession, "system", $repairSkill).GetAwaiter().GetResult()
if (-not $repairResult.Success -or
    $script:nonRetryableExecutions -ne 1 -or
    $script:fallbackExecutions -ne 1 -or
    $script:repairModelCalls -ne 5) {
    throw "Non-retryable duplicate guard/adaptive fallback failed: calls=$script:repairModelCalls primary=$script:nonRetryableExecutions fallback=$script:fallbackExecutions result=$($repairResult.Message)"
}

# Only explicit task/session fatal results terminate the loop. They are separate from the
# legacy Recoverable flag and from Retryable (which concerns only an identical call).
$fatalRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$fatalDescriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
$fatalDescriptor.Id = "FatalAction"
$fatalDescriptor.Name = "Fatal action"
$fatalDescriptor.AppType = "excel"
$fatalDescriptor.RiskLevel = "safe"
$fatalDescriptor.AccessMode = "write"
$fatalDescriptor.OutcomeEffects.Add("artifact")
$fatalRegistry.RegisterTool($fatalDescriptor)
$script:fatalExecutions = 0
$fatalRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:fatalExecutions += 1
    $fatal = [ShareRibbon.Agent.ToolResult]::new()
    $fatal.Success = $false
    $fatal.ToolId = "FatalAction"
    $fatal.Message = "runtime cannot continue"
    $fatal.UserMessage = "runtime cannot continue"
    $fatal.ErrorCode = "TEST_TASK_FATAL"
    $fatal.TaskFatal = $true
    $fatal.SessionFatal = $true
    $fatal.Retryable = $false
    return $fatal
}
$fatalLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $fatalRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:fatalModelCalls = 0
$fatalLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:fatalModelCalls += 1
    if ($script:fatalModelCalls -eq 1) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"understanding":"fatal","steps":[{"step":1,"description":"run","toolHint":"FatalAction"}],"summary":"done","capabilityGap":"","outcomeContract":{"schemaVersion":"1.0","requirements":[{"id":"fatal-action","targetRef":"Excel:fatal","effectType":"artifact","operator":"exists","criterionIds":["criterion-raw-user-request"],"required":true}]}}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"act","thought":"run","action":{"tool":"FatalAction","params":{}}}')
}
$fatalSession = [ShareRibbon.Agent.AgentSession]::new("fatal", "Excel", "")
$fatalSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$fatalSession.Spec.MutationPolicy = "allow"
$fatalSession.Spec.Goal = "fatal"
$fatalSkill = [ShareRibbon.Agent.AgentSkill]::new()
$fatalSkill.RequiredTools.Add("FatalAction")
$fatalResult = $fatalLoop.RunAsync($fatalSession, "system", $fatalSkill).GetAwaiter().GetResult()
if ($fatalResult.Success -or
    $script:fatalExecutions -ne 1 -or
    $script:fatalModelCalls -ne 2 -or
    -not $fatalResult.TaskFatal -or
    -not $fatalResult.SessionFatal -or
    $fatalResult.ErrorCode -ne "TEST_TASK_FATAL" -or
    -not $fatalResult.Message.Contains("taskFatal=True")) {
    throw "Explicit task-fatal result did not terminate exactly once: calls=$script:fatalModelCalls executions=$script:fatalExecutions result=$($fatalResult.Message)"
}

$sessionFatal = [ShareRibbon.Agent.ToolResult]::Failed(
    "SessionFatal",
    "session fatal",
    $null,
    "TEST_SESSION_FATAL",
    "session fatal",
    "session fatal",
    $true,
    $null,
    $null,
    $false,
    $true,
    $true)
if (-not $sessionFatal.SessionFatal -or -not $sessionFatal.TaskFatal -or $sessionFatal.Retryable) {
    throw "Session-fatal result does not imply task-fatal/non-retryable semantics"
}

# AgentKernel must fail closed instead of waiting forever when no approval consumer exists.
$unhandledApprovalKernel = [ShareRibbon.Agent.AgentKernel]::new()
$unhandledApprovalKernel.Initialize()
$loopField = [ShareRibbon.Agent.AgentKernel].GetField(
    "_loopEngine",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$unhandledApprovalLoop = [ShareRibbon.Agent.LoopEngine]$loopField.GetValue($unhandledApprovalKernel)
$unhandledApprovalTask = $unhandledApprovalLoop.OnRequestApproval.Invoke("approval required")
$unhandledApprovalWinner = [System.Threading.Tasks.Task]::WhenAny(
    $unhandledApprovalTask,
    [System.Threading.Tasks.Task]::Delay(2000)).GetAwaiter().GetResult()
if ($unhandledApprovalWinner -ne $unhandledApprovalTask -or -not $unhandledApprovalTask.IsFaulted -or
    -not $unhandledApprovalTask.Exception.ToString().Contains("No approval handler is registered")) {
    throw "Approval without a subscriber did not fail closed"
}

# The Harness approval handshake is reusable. Two destructive actions in one Agent run must
# surface two separate AwaitingApproval results; rejecting an approval must preserve task-fatal
# semantics through AgentResult and HarnessRunResult.
if (-not ("ApprovalRuntimeProbe" -as [type])) {
    $newtonsoftAssemblyPath = Join-Path (Split-Path $assemblyPath) "Newtonsoft.Json.dll"
    Add-Type -ReferencedAssemblies $assemblyPath,$newtonsoftAssemblyPath -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ShareRibbon;
using ShareRibbon.Agent;

public static class ApprovalRuntimeProbe
{
    public static string Scenario { get; private set; }
    public static int ModelCalls { get; private set; }
    public static int HostExecutions { get; private set; }
    public static readonly Func<string, string, List<HistoryMessage>, Task<string>> ModelDelegate = Model;
    public static readonly Func<string, string, bool, ToolResult> HostDelegate = Host;

    public static void Reset(string scenario)
    {
        Scenario = scenario;
        ModelCalls = 0;
        HostExecutions = 0;
    }

    private static Task<string> Model(string prompt, string system, List<HistoryMessage> history)
    {
        ModelCalls++;
        if (Scenario == "two")
        {
            switch (ModelCalls)
            {
                case 1:
                    return Task.FromResult("{\"understanding\":\"two approvals\",\"steps\":[{\"step\":1,\"description\":\"first\",\"toolHint\":\"ApprovalWrite\"},{\"step\":2,\"description\":\"second\",\"toolHint\":\"ApprovalWrite\"}],\"summary\":\"done\",\"capabilityGap\":\"\",\"outcomeContract\":{\"schemaVersion\":\"1.0\",\"requirements\":[{\"id\":\"approved-write-1\",\"targetRef\":\"Excel:approval/1\",\"effectType\":\"artifact\",\"operator\":\"exists\",\"criterionIds\":[\"criterion-raw-user-request\"],\"required\":true},{\"id\":\"approved-write-2\",\"targetRef\":\"Excel:approval/2\",\"effectType\":\"artifact\",\"operator\":\"exists\",\"criterionIds\":[\"criterion-raw-user-request\"],\"required\":true}]}}");
                case 2:
                    return Task.FromResult("{\"decision\":\"act\",\"thought\":\"first\",\"action\":{\"tool\":\"ApprovalWrite\",\"params\":{\"ordinal\":1}}}");
                case 3:
                    return Task.FromResult("{\"decision\":\"act\",\"thought\":\"second\",\"action\":{\"tool\":\"ApprovalWrite\",\"params\":{\"ordinal\":2}}}");
                default:
                    return Task.FromResult("{\"decision\":\"complete\",\"thought\":\"both writes observed\",\"message\":\"done\",\"evidence\":[\"obs-1/e1\",\"obs-2/e1\"]}");
            }
        }

        if (ModelCalls == 1)
            return Task.FromResult("{\"understanding\":\"reject approval\",\"steps\":[{\"step\":1,\"description\":\"write\",\"toolHint\":\"ApprovalWrite\"}],\"summary\":\"done\",\"capabilityGap\":\"\",\"outcomeContract\":{\"schemaVersion\":\"1.0\",\"requirements\":[{\"id\":\"rejected-write\",\"targetRef\":\"Excel:approval/3\",\"effectType\":\"artifact\",\"operator\":\"exists\",\"criterionIds\":[\"criterion-raw-user-request\"],\"required\":true}]}}");
        return Task.FromResult("{\"decision\":\"act\",\"thought\":\"request write\",\"action\":{\"tool\":\"ApprovalWrite\",\"params\":{\"ordinal\":3}}}");
    }

    private static ToolResult Host(string code, string language, bool preview)
    {
        HostExecutions++;
        var command = JObject.Parse(code);
        var ordinal = command["params"].Value<int>("ordinal");
        var observation = JObject.Parse("{\"kind\":\"write\",\"summary\":\"approval action complete\",\"changed\":true,\"satisfied\":true,\"targetRefs\":[\"Excel:approval/" + ordinal + "\"]}");
        return ToolResult.Succeed("ApprovalWrite", "approval action " + ordinal + " complete", null, observation);
    }
}
'@
}

$approvalKernel = [ShareRibbon.Agent.AgentKernel]::new()
$approvalKernel.Initialize()
$approvalRegistry = [ShareRibbon.Agent.ToolRegistry]$loopField.DeclaringType.GetField(
    "_toolRegistry",
    [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).GetValue($approvalKernel)
$approvalDescriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
$approvalDescriptor.Id = "ApprovalWrite"
$approvalDescriptor.Name = "Approval write"
$approvalDescriptor.AppType = "excel"
$approvalDescriptor.RiskLevel = "risky"
$approvalDescriptor.AccessMode = "write"
$approvalDescriptor.OutcomeEffects.Add("artifact")
$approvalRegistry.RegisterTool($approvalDescriptor)
$approvalKernel.ExecuteCodeWithToolResult = [ApprovalRuntimeProbe]::HostDelegate
$approvalSkill = [ShareRibbon.SkillFileDefinition]::new()
$approvalSkill.Name = "Approval smoke"
$approvalSkill.Description = "Approval state-machine regression"
$approvalSkill.Application = "Excel"
$approvalSkill.AllowedTools = [System.Collections.Generic.List[string]]::new()
$approvalSkill.AllowedTools.Add("ApprovalWrite")
$approvalSkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
$approvalSkills.Add($approvalSkill)
$approvalHarness = [ShareRibbon.Agent.Harness.OfficeHarness]::new($approvalKernel)
[ApprovalRuntimeProbe]::Reset("two")
$approvalKernel.SendAIRequest = [ApprovalRuntimeProbe]::ModelDelegate
$approvalSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$approvalSpec.MutationPolicy = "allow"
$approvalSpec.Goal = "perform two approved writes"
$approvalSpec.RequiredTools.Add("ApprovalWrite")
$approvalTurn = [ShareRibbon.Agent.Harness.UserTurn]::new()
$approvalTurn.AppType = "Excel"
$approvalTurn.Text = "perform two approved writes"
$approvalTurn.TaskSpec = $approvalSpec
$approvalTurn.SelectedSkills = $approvalSkills
$approvalFirst = $approvalHarness.RunAsync(
    $approvalTurn,
    [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
if ($approvalFirst.Status -ne [ShareRibbon.Agent.Harness.HarnessRunStatus]::AwaitingApproval) {
    throw "First approval request was not surfaced by the Harness"
}
$approvalSecond = $approvalHarness.ApproveAsync(
    $approvalFirst.RunId,
    $true,
    [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
if ($approvalSecond.Status -ne [ShareRibbon.Agent.Harness.HarnessRunStatus]::AwaitingApproval) {
    throw "Second approval request was lost after resuming the first"
}
$approvalComplete = $approvalHarness.ApproveAsync(
    $approvalFirst.RunId,
    $true,
    [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
if ($approvalComplete.Status -ne [ShareRibbon.Agent.Harness.HarnessRunStatus]::Succeeded -or
    [ApprovalRuntimeProbe]::HostExecutions -ne 2 -or
    $approvalComplete.TaskFatal -or
    $approvalComplete.SessionFatal) {
    throw "Reusable approval handshake failed: status=$($approvalComplete.Status) hostExecutions=$([ApprovalRuntimeProbe]::HostExecutions)"
}

[ApprovalRuntimeProbe]::Reset("reject")
$rejectSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$rejectSpec.MutationPolicy = "allow"
$rejectSpec.Goal = "request a write and reject it"
$rejectSpec.RequiredTools.Add("ApprovalWrite")
$rejectTurn = [ShareRibbon.Agent.Harness.UserTurn]::new()
$rejectTurn.AppType = "Excel"
$rejectTurn.Text = "request a write and reject it"
$rejectTurn.TaskSpec = $rejectSpec
$rejectTurn.SelectedSkills = $approvalSkills
$rejectFirst = $approvalHarness.RunAsync(
    $rejectTurn,
    [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
$rejectResult = $approvalHarness.ApproveAsync(
    $rejectFirst.RunId,
    $false,
    [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
if ($rejectFirst.Status -ne [ShareRibbon.Agent.Harness.HarnessRunStatus]::AwaitingApproval -or
    $rejectResult.Status -ne [ShareRibbon.Agent.Harness.HarnessRunStatus]::Failed -or
    -not $rejectResult.TaskFatal -or
    $rejectResult.SessionFatal -or
    $rejectResult.ErrorCode -ne "SAFETY_BLOCKED" -or
    [ApprovalRuntimeProbe]::HostExecutions -ne 0) {
    throw "Approval rejection did not propagate task-fatal semantics"
}

# Sandboxed JSON-only Python computation is a controlled compute operation and should
# not require an approval round-trip. The sandbox restrictions themselves remain active.
$pythonSchema = Get-Content -LiteralPath (Join-Path $toolsPath "excel\PythonCompute.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if ($pythonSchema.accessMode -ne "compute" -or $pythonSchema.riskLevel -eq "risky") {
    throw "PythonCompute is still approval-gated instead of controlled compute"
}
$pythonParams = [Newtonsoft.Json.Linq.JObject]::Parse('{"code":"result = sum(input_data)","input":[1,2,3]}')
$pythonTask = [ShareRibbon.Services.Python.PythonComputeService]::ExecuteAsync($pythonParams)
$pythonResult = $pythonTask.GetAwaiter().GetResult()
if (-not $pythonResult.Success -or $pythonResult.Data.ToString() -ne "6") {
    throw "Controlled PythonCompute smoke failed: $($pythonResult.ErrorMessage)"
}

# Python is a separate controlled process, so unlike an in-flight Office COM write it
# can and must be terminated promptly when the user presses Stop.
$pythonCancelParams = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"code":"total = 0\nfor i in range(100000000):\n    total += i\nresult = total","input":null,"timeoutSeconds":60}')
$pythonCancelSource = [System.Threading.CancellationTokenSource]::new()
$pythonCancelWatch = [System.Diagnostics.Stopwatch]::StartNew()
$pythonCancelTask = [ShareRibbon.Services.Python.PythonComputeService]::ExecuteAsync(
    $pythonCancelParams,
    $pythonCancelSource.Token)
[System.Threading.Thread]::Sleep(100)
$pythonCancelSource.Cancel()
$pythonWasCancelled = $false
try {
    $null = $pythonCancelTask.GetAwaiter().GetResult()
} catch [System.OperationCanceledException] {
    $pythonWasCancelled = $true
}
$pythonCancelWatch.Stop()
$pythonCancelSource.Dispose()
if (-not $pythonWasCancelled -or $pythonCancelWatch.ElapsedMilliseconds -gt 3000) {
    throw "PythonCompute did not terminate promptly on cancellation: cancelled=$pythonWasCancelled elapsed=$($pythonCancelWatch.ElapsedMilliseconds)ms"
}

# json is already the transport format and is loaded by the sandbox wrapper itself. A
# user program may safely use that standard-library module without gaining file, network,
# process, or Office access.
$jsonPythonParams = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"code":"import json\nresult = json.loads(\"{\\\"value\\\":7}\")[\"value\"]","input":null}')
$jsonPythonResult = [ShareRibbon.Services.Python.PythonComputeService]::ExecuteAsync(
    $jsonPythonParams).GetAwaiter().GetResult()
if (-not $jsonPythonResult.Success -or $jsonPythonResult.Data.ToString() -ne "7") {
    throw "Safe Python json import is still blocked: $($jsonPythonResult.ErrorMessage)"
}

# Some providers double-escape tool string arguments, leaving literal backslash-n pairs
# after JSON parsing. The compute boundary must normalize that transport representation
# before validating imports or parsing Python source.
$escapedPythonParams = [Newtonsoft.Json.Linq.JObject]::new()
$escapedPythonParams["code"] = [Newtonsoft.Json.Linq.JValue]::new(
    "import json\nresult = json.loads('{`"value`":9}')[`"value`"]")
$escapedPythonParams["input"] = [Newtonsoft.Json.Linq.JValue]::CreateNull()
$escapedPythonResult = [ShareRibbon.Services.Python.PythonComputeService]::ExecuteAsync(
    $escapedPythonParams).GetAwaiter().GetResult()
if (-not $escapedPythonResult.Success -or $escapedPythonResult.Data.ToString() -ne "9") {
    throw "Double-escaped Python source was not normalized at the transport boundary: $($escapedPythonResult.ErrorMessage)"
}

# If generated source is syntactically invalid, Python can exit before consuming stdin.
# The broken input pipe is secondary; callers must receive the Python SyntaxError so the
# adaptive loop can correct the code instead of diagnosing a fake filesystem problem.
$invalidPythonParams = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"code":"result = []; for row in input_data: result.append(row)","input":[1,2,3]}')
$invalidPythonResult = [ShareRibbon.Services.Python.PythonComputeService]::ExecuteAsync(
    $invalidPythonParams).GetAwaiter().GetResult()
if ($invalidPythonResult.Success -or
    $invalidPythonResult.ErrorCode -eq "IO_ERROR" -or
    $invalidPythonResult.DebugDetail -notmatch "SyntaxError") {
    throw "Python syntax failure is still masked by a secondary stdin pipe error: code=$($invalidPythonResult.ErrorCode) detail=$($invalidPythonResult.DebugDetail)"
}

# The planner cannot know the rows returned by a future ReadRange call. The runtime
# dataflow seam must replace preview/sample payloads with actual successful tool output.
$dataflow = [ShareRibbon.Agent.AgentToolDataflow]::new()
$readData = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"workbook":"book.xlsx","sheet":"SalesData","address":"A1:B5","rowCount":5,"columnCount":2,"values":[["\u533a\u57df","\u9500\u552e\u989d"],["\u534e\u4e1c",1000],["\u534e\u4e1c",2000],["\u534e\u5317",1500],["\u534e\u5317",2500]]}')
$dataflow.RecordSuccess([ShareRibbon.Agent.ToolResult]::Succeed("ReadRange", "read", $readData))

$computeCall = [ShareRibbon.Agent.ToolCall]::new()
$computeCall.ToolId = "PythonCompute"
$computeCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"code":"result = input_data[\"rows\"]","input":{"headers":["\u533a\u57df","\u9500\u552e\u989d"],"rows":[["preview",1]]}}')
$dataflow.BindInputs($computeCall)
$boundInput = $computeCall.Parameters["input"].ToString([Newtonsoft.Json.Formatting]::None)
$expectedInput = $readData.ToString([Newtonsoft.Json.Formatting]::None)
if ($boundInput -ne $expectedInput) {
    throw "PythonCompute did not receive the full ReadRange result: $boundInput"
}

$computedRows = [Newtonsoft.Json.Linq.JArray]::Parse(
    '[["\u533a\u57df","\u5e73\u5747\u9500\u552e\u989d"],["\u534e\u4e1c",1500.0],["\u534e\u5317",2000.0]]')
$dataflow.RecordSuccess([ShareRibbon.Agent.ToolResult]::Succeed("PythonCompute", "computed", $computedRows))
$writeCall = [ShareRibbon.Agent.ToolCall]::new()
$writeCall.ToolId = "WriteData"
$writeCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"targetRange":"PythonAverage!A1","data":[["wrong",-1]]}')
$dataflow.BindInputs($writeCall)
$writtenJson = $writeCall.Parameters["data"].ToString([Newtonsoft.Json.Formatting]::None)
if ($writtenJson -ne $computedRows.ToString([Newtonsoft.Json.Formatting]::None)) {
    throw "WriteData did not receive the PythonCompute result: $writtenJson"
}

# Python often returns a JSON object for named scalar statistics.  Runtime binding must
# project that object to a rectangular Excel matrix; passing JObject through as a single
# Value2 payload reaches COM as an unsupported value and fails with 0x800A03EC.
$computedStatistics = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"\u5e73\u5747\u503c":1904.6666666666667,"\u4e2d\u4f4d\u6570":1768,"\u6700\u5927\u503c":3192,"\u6700\u5c0f\u503c":1000}')
$statisticsDataflow = [ShareRibbon.Agent.AgentToolDataflow]::new()
$statisticsDataflow.RecordSuccess(
    [ShareRibbon.Agent.ToolResult]::Succeed("PythonCompute", "computed statistics", $computedStatistics))
$statisticsWriteCall = [ShareRibbon.Agent.ToolCall]::new()
$statisticsWriteCall.ToolId = "WriteData"
$statisticsWriteCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"targetRange":"PythonStats!A1","data":[["preview",-1]]}')
$statisticsDataflow.BindInputs($statisticsWriteCall)
$statisticsWritten = $statisticsWriteCall.Parameters["data"]
$expectedStatistics = [Newtonsoft.Json.Linq.JArray]::Parse(
    '[["\u5b57\u6bb5","\u503c"],["\u5e73\u5747\u503c",1904.6666666666667],["\u4e2d\u4f4d\u6570",1768],["\u6700\u5927\u503c",3192],["\u6700\u5c0f\u503c",1000]]')
if ($statisticsWritten.Type -ne [Newtonsoft.Json.Linq.JTokenType]::Array -or
    $statisticsWritten.ToString([Newtonsoft.Json.Formatting]::None) -ne
        $expectedStatistics.ToString([Newtonsoft.Json.Formatting]::None)) {
    throw "Python scalar-object result was not projected to a rectangular WriteData matrix: $($statisticsWritten.ToString([Newtonsoft.Json.Formatting]::None))"
}

Write-Host "PASS: agent runtime safety, code execution, and latency policies"
