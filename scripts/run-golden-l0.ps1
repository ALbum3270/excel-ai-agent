param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$assemblyPath = Join-Path $repoRoot "ShareRibbon\bin\$Configuration\ShareRibbon.dll"
$catalogPath = Join-Path $repoRoot "tests\golden\l0\catalog.json"
$toolsPath = Join-Path $repoRoot "ShareRibbon\Tools"

if (-not (Test-Path -LiteralPath $assemblyPath)) {
    throw "ShareRibbon assembly not found: $assemblyPath. Build the code projects first."
}
if (-not (Test-Path -LiteralPath $catalogPath)) {
    throw "Golden L0 catalog not found: $catalogPath"
}

Add-Type -Path $assemblyPath
$catalog = Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
$results = @()

$buildStructuredExcelIntent = {
    param(
        [string]$text,
        [string]$interactionMode = "execute",
        [ShareRibbon.OfficeIntentType]$officeIntent = [ShareRibbon.OfficeIntentType]::GENERAL_QUERY,
        [ShareRibbon.ExcelIntentType]$excelIntent = [ShareRibbon.ExcelIntentType]::GENERAL_QUERY
    )

    $structuredIntent = [ShareRibbon.IntentResult]::new()
    $structuredIntent.OfficeIntent = $officeIntent
    $structuredIntent.IntentType = $excelIntent
    $structuredIntent.Confidence = 0.95
    $structuredIntent.ResponseMode = $interactionMode
    $candidate = [ShareRibbon.Agent.Goals.CandidateGoalContract]::new()
    $clause = [ShareRibbon.Agent.Goals.CandidateGoalSourceClause]::new()
    $clause.Id = "clause-1"
    $clause.Text = $text
    $clause.IsExplicit = $true
    $clause.SourceStart = 0
    $candidate.SourceClauses.Add($clause)
    $criterion = [ShareRibbon.Agent.Goals.CandidateGoalCriterion]::new()
    $criterion.Id = "criterion-1"
    $criterion.Statement = $text
    $criterion.Kind = "semantic"
    $criterion.SourceClauseIds.Add("clause-1")
    $criterion.Required = $true
    $candidate.Criteria.Add($criterion)
    $payload = [ShareRibbon.Agent.Goals.GoalInterpretationPayload]::new()
    $payload.Candidate = $candidate
    $structuredIntent.GoalInterpretation = $payload
    return $structuredIntent
}

foreach ($case in $catalog.cases) {
    if ($case.kind -eq "contract") {
        if ($case.contractType -eq "context_pack") {
            $officeContext = New-Object ShareRibbon.Agent.Context.OfficeContext
            $officeContext.AppType = "Word"
            $officeContext.Selection = New-Object ShareRibbon.Agent.Context.SelectionInfo
            $officeContext.Selection.Address = "Selection"
            $officeContext.Selection.ItemCount = 1
            $officeContext.Selection.DataType = "text"
            $officeContext.Selection.Preview = ("x" * 7000)
            $pack = [ShareRibbon.Agent.Context.ContextPack]::FromOfficeContext($officeContext, ("y" * 9000), 4000)
            if ($pack.SchemaVersion -ne $case.expectedSchemaVersion -or
                $pack.Budget.Strategy -ne $case.expectedBudgetStrategy -or
                -not $pack.Budget.Truncated -or
                $pack.Budget.UsedChars -gt $pack.Budget.MaxChars -or
                [string]::IsNullOrWhiteSpace($pack.ToJson())) {
                throw "Golden case failed: $($case.id)"
            }
        }
        elseif ($case.contractType -eq "excel_table_region_context") {
            $officeContext = New-Object ShareRibbon.Agent.Context.OfficeContext
            $officeContext.AppType = "Excel"
            $officeContext.Selection = New-Object ShareRibbon.Agent.Context.SelectionInfo
            $officeContext.Selection.Address = "Sheet1!A1:C4"
            $officeContext.Selection.ItemCount = 12
            $officeContext.Selection.DataType = "table"
            $officeContext.Selection.Preview = "Month`tSales`tRegion"
            $officeContext.HostData = [Newtonsoft.Json.Linq.JObject]::Parse(
                '{"tables":[{"sheet":"Sheet1","address":"A1:C4","hasHeader":true,"rowCount":3,"colCount":3,"source":"selection"}]}'
            )

            $pack = [ShareRibbon.Agent.Context.ContextPack]::FromOfficeContext($officeContext, "", 4000)
            if ($null -eq $pack.Host -or
                $null -eq $pack.Host["tables"] -or
                $pack.Host["tables"].Count -ne 1 -or
                $pack.Host["tables"][0]["address"].ToString() -ne "A1:C4" -or
                $pack.ToPromptText() -notmatch "Host context" -or
                $pack.Budget.UsedChars -gt $pack.Budget.MaxChars) {
                throw "Golden case failed: $($case.id)"
            }

            $providerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\Context\ExcelContextProvider.vb") -Raw -Encoding UTF8
            $detectorSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\Context\ExcelTableRegion.vb") -Raw -Encoding UTF8
            if ($providerSource -notmatch 'ExcelTableRegionDetector' -or
                $providerSource -notmatch 'HostData\("tables"\)' -or
                $providerSource -notmatch '识别数据区域预览' -or
                $providerSource -notmatch 'recommendedRange\s*=\s*GetRangeAddress\(usedRange\)' -or
                $detectorSource -notmatch 'MaxSampleRows\s+As\s+Integer\s*=\s*200' -or
                $detectorSource -notmatch '"list_object"' -or
                $detectorSource -notmatch '"current_region"') {
                throw "Golden case $($case.id): Excel TableRegion detection is not connected to ContextPack"
            }
        }
        elseif ($case.contractType -eq "python_compute_contract") {
            $toolPath = Join-Path $repoRoot "ShareRibbon\Tools\excel\PythonCompute.json"
            if (-not (Test-Path -LiteralPath $toolPath)) {
                throw "Golden case $($case.id): PythonCompute tool schema is missing"
            }
            $toolSchema = Get-Content -LiteralPath $toolPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($toolSchema.id -ne "PythonCompute" -or
                $toolSchema.appType -ne "excel" -or
                $toolSchema.riskLevel -eq "risky" -or
                $toolSchema.accessMode -ne "compute") {
                throw "Golden case $($case.id): PythonCompute must be controlled compute without approval"
            }

            $serviceSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Services\Python\PythonComputeService.vb") -Raw -Encoding UTF8
            $registrySource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\ToolRegistry.vb") -Raw -Encoding UTF8
            $skillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw -Encoding UTF8
            if ($serviceSource -notmatch 'MaxTimeoutSeconds\s+As\s+Integer\s*=\s*60' -or
                $serviceSource -notmatch 'StandardInput\.BaseStream\.WriteAsync\(inputBytes' -or
                $serviceSource -notmatch 'OFFICE_AI_PYTHON_PATH' -or
                $serviceSource -notmatch 'PythonCompute\s+不允许导入模块' -or
                $serviceSource -notmatch 'ast\.parse' -or
                $serviceSource -notmatch 'safe_builtins' -or
                $registrySource -notmatch 'PythonComputeService\.ExecuteAsync' -or
                $skillSource -notmatch 'allowed-tools:[^\r\n]*PythonCompute') {
                throw "Golden case $($case.id): PythonCompute safety or routing contract is incomplete"
            }
        }
        elseif ($case.contractType -eq "read_range_contract") {
            $toolPath = Join-Path $repoRoot "ShareRibbon\Tools\excel\ReadRange.json"
            if (-not (Test-Path -LiteralPath $toolPath)) {
                throw "Golden case $($case.id): ReadRange tool schema is missing"
            }
            $toolSchema = Get-Content -LiteralPath $toolPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($toolSchema.id -ne "ReadRange" -or
                $toolSchema.appType -ne "excel" -or
                $toolSchema.riskLevel -ne "safe") {
                throw "Golden case $($case.id): ReadRange must be a safe Excel read tool"
            }

            $excelChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw -Encoding UTF8
            $readAdapterSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\Runtime\ExcelReadRangeAdapter.vb") -Raw -Encoding UTF8
            $skillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw -Encoding UTF8
            if ($excelChatSource -notmatch 'ExcelReadRangeAdapter\.Execute' -or
                $readAdapterSource -notmatch 'Math\.Min\(20000,\s*maxCells\)' -or
                $readAdapterSource -notmatch '\{"kind",\s*"read"\}' -or
                $readAdapterSource -notmatch 'ExcelMatrixToJson' -or
                $skillSource -notmatch 'allowed-tools:[^\r\n]*ReadRange' -or
                $skillSource -notmatch 'First use `ReadRange`') {
                throw "Golden case $($case.id): ReadRange routing or skill contract is incomplete"
            }
        }
        elseif ($case.contractType -eq "excel_basic_required_tools") {
            $requiredTools = @("FormatRange", "CreateChart", "CreateSheet")
            $skillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw -Encoding UTF8
            $schemaSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelJsonCommandSchema.vb") -Raw -Encoding UTF8
            foreach ($requiredTool in $requiredTools) {
                $toolPath = Join-Path $repoRoot "ShareRibbon\Tools\excel\$requiredTool.json"
                if (-not (Test-Path -LiteralPath $toolPath)) {
                    throw "Golden case $($case.id): required Excel tool is missing: $requiredTool"
                }
                $toolSchema = Get-Content -LiteralPath $toolPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($toolSchema.id -ne $requiredTool -or $toolSchema.appType -ne "excel") {
                    throw "Golden case $($case.id): invalid required Excel tool: $requiredTool"
                }
                if ($skillSource -notmatch "allowed-tools:[^\r\n]*$requiredTool" -or
                    $schemaSource -notmatch "`"$requiredTool`"") {
                    throw "Golden case $($case.id): required Excel tool is not routed by Skill/schema: $requiredTool"
                }
            }

            $promptManagerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Config\PromptManager.vb") -Raw -Encoding UTF8
            if ($promptManagerSource -notmatch 'excelApp\.JsonSchemaConstraint\s*=\s*GetExcelJsonSchemaConstraintDefault\(\)' -or
                $promptManagerSource -notmatch '【Excel command类型】' -or
                $promptManagerSource -notmatch 'Not\s+constraint\.Contains\("CreateSheet"\)') {
                throw "Golden case $($case.id): Excel prompt default or legacy five-command migration is incomplete"
            }

            $effectiveConstraint = [ShareRibbon.PromptManager]::Instance.GetJsonSchemaConstraint("Excel")
            if ($effectiveConstraint -notmatch '【Excel支持的25个命令】' -or
                $effectiveConstraint -notmatch 'CreateSheet') {
                throw "Golden case $($case.id): effective Excel prompt does not expose CreateSheet"
            }
        }
        elseif ($case.contractType -eq "office_action_followup_routing") {
            $decisionMethod = [ShareRibbon.ChatRoutingOrchestrator].GetMethod("DecidePostAnalysisRoute")
            if ($null -eq $decisionMethod) {
                throw "Golden case $($case.id): post-analysis route decision seam is missing"
            }

            $executeIntent = [ShareRibbon.IntentResult]::new()
            $executeIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::FORMAT_STYLE
            $executeIntent.ResponseMode = "execute"
            $executeDecision = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute($true, $executeIntent, $false)
            if ($executeDecision.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): an explicit follow-up action was routed to $executeDecision"
            }

            $openEndedIntent = [ShareRibbon.IntentResult]::new()
            $openEndedIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::GENERAL_QUERY
            $openEndedIntent.ResponseMode = "execute"
            $openEndedDecision = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute($true, $openEndedIntent, $true)
            if ($openEndedDecision.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): an open-ended Office action was routed to $openEndedDecision"
            }

            $answerIntent = [ShareRibbon.IntentResult]::new()
            $answerIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::GENERAL_QUERY
            $answerIntent.ResponseMode = "answer"
            $answerDecision = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute($true, $answerIntent, $false)
            if ($answerDecision.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): a conversational follow-up bypassed the unified adaptive Agent"
            }

            $legacyPreCheckMethod = [ShareRibbon.ChatRoutingOrchestrator].GetMethod("ShouldRunLegacyPreCheck")
            if ($null -eq $legacyPreCheckMethod) {
                throw "Golden case $($case.id): legacy formatting pre-check has no host-aware decision seam"
            }
            $excelFormatIntent = [ShareRibbon.IntentResult]::new()
            $excelFormatIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::FORMAT_STYLE
            $wordFormatIntent = [ShareRibbon.IntentResult]::new()
            $wordFormatIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::FORMAT_STYLE
            if ([bool]$legacyPreCheckMethod.Invoke($null, @("Excel", $excelFormatIntent)) -or
                -not [bool]$legacyPreCheckMethod.Invoke($null, @("Word", $wordFormatIntent))) {
                throw "Golden case $($case.id): Excel formatting can still be blocked by the Word-only legacy pre-check"
            }

            $routingInterface = [ShareRibbon.IChatRoutingHost]
            $blockedCompletionMethod = $routingInterface.GetMethod("CompleteBlockedRequest")
            $routingSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb") -Raw -Encoding UTF8
            if ($null -eq $blockedCompletionMethod -or
                $routingSource -notmatch '_host\.CompleteBlockedRequest\(') {
                throw "Golden case $($case.id): a blocked route can still finish without a visible terminal response"
            }

            $intentService = [ShareRibbon.IntentRecognitionService]::new("Excel")
            $mergeIntentMethod = [ShareRibbon.IntentRecognitionService].GetMethod(
                "MergeLlmIntentResult",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            $keywordIntent = [ShareRibbon.IntentResult]::new()
            $modelIntent = [ShareRibbon.IntentResult]::new()
            $modelIntent.Confidence = 0.9
            $modelIntent.ResponseMode = "execute"
            $modelIntent.RequestedOutputs.Add("worksheet")
            $mergeIntentMethod.Invoke($null, @($keywordIntent, $modelIntent)) | Out-Null
            if ($null -eq $mergeIntentMethod -or
                $keywordIntent.ResponseMode -ne "execute" -or
                -not $keywordIntent.RequestedOutputs.Contains("worksheet")) {
                throw "Golden case $($case.id): parsed interaction mode is discarded before routing"
            }
            # The accepted model result is a snapshot, not an alias to mutable model output.
            $modelIntent.RequestedOutputs.Add("chart")
            if ($keywordIntent.RequestedOutputs.Contains("chart")) {
                throw "Golden case $($case.id): accepted requestedOutputs still alias the model candidate"
            }

            # Candidate trust is atomic. A low-confidence answer must not partially overwrite a
            # deterministic execute result while leaving its semantic intent behind.
            $deterministicIntent = [ShareRibbon.IntentResult]::new()
            $deterministicIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::FORMAT_STYLE
            $deterministicIntent.IntentType = [ShareRibbon.ExcelIntentType]::FORMAT_STYLE
            $deterministicIntent.Confidence = 0.8
            $deterministicIntent.ResponseMode = "execute"
            $deterministicIntent.UserFriendlyDescription = "deterministic"
            $deterministicIntent.RequestedOutputs.Add("formatted_content")
            $lowConfidenceAnswer = [ShareRibbon.IntentResult]::new()
            $lowConfidenceAnswer.OfficeIntent = [ShareRibbon.OfficeIntentType]::GENERAL_QUERY
            $lowConfidenceAnswer.IntentType = [ShareRibbon.ExcelIntentType]::GENERAL_QUERY
            $lowConfidenceAnswer.Confidence = 0.3
            $lowConfidenceAnswer.ResponseMode = "answer"
            $lowConfidenceAnswer.UserFriendlyDescription = "untrusted"
            $mergeIntentMethod.Invoke($null, @($deterministicIntent, $lowConfidenceAnswer)) | Out-Null
            if ($deterministicIntent.OfficeIntent -ne [ShareRibbon.OfficeIntentType]::FORMAT_STYLE -or
                $deterministicIntent.IntentType -ne [ShareRibbon.ExcelIntentType]::FORMAT_STYLE -or
                $deterministicIntent.Confidence -ne 0.8 -or
                $deterministicIntent.ResponseMode -ne "execute" -or
                $deterministicIntent.UserFriendlyDescription -ne "deterministic" -or
                $deterministicIntent.RequestedOutputs.Count -ne 1 -or
                -not $deterministicIntent.RequestedOutputs.Contains("formatted_content")) {
                throw "Golden case $($case.id): low-confidence model routing was partially accepted"
            }

            # A route candidate whose answer/clarify mode conflicts with execution artifacts is
            # rejected as a whole instead of allowing requestedOutputs to force AgentKernel.
            $inconsistentAnswer = [ShareRibbon.IntentResult]::new()
            $inconsistentAnswer.Confidence = 0.95
            $inconsistentAnswer.ResponseMode = "answer"
            $inconsistentAnswer.RequestedOutputs.Add("worksheet")
            $mergeIntentMethod.Invoke($null, @($deterministicIntent, $inconsistentAnswer)) | Out-Null
            if ($deterministicIntent.ResponseMode -ne "execute" -or
                $deterministicIntent.RequestedOutputs.Count -ne 1 -or
                -not $deterministicIntent.RequestedOutputs.Contains("formatted_content")) {
                throw "Golden case $($case.id): answer plus requestedOutputs was partially accepted"
            }

            $invalidModeIntent = [ShareRibbon.IntentResult]::new()
            $invalidModeIntent.Confidence = 0.95
            $invalidModeIntent.ResponseMode = "run"
            $invalidModeIntent.RequestedOutputs.Add("worksheet")
            $mergeIntentMethod.Invoke($null, @($deterministicIntent, $invalidModeIntent)) | Out-Null
            if ($deterministicIntent.ResponseMode -ne "execute" -or
                $deterministicIntent.RequestedOutputs.Count -ne 1) {
                throw "Golden case $($case.id): illegal interactionMode was accepted"
            }
            $promptMethod = [ShareRibbon.IntentRecognitionService].GetMethod(
                "GetEnhancedExcelIntentRecognitionPrompt",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $intentPrompt = [string]$promptMethod.Invoke($intentService, @())
            if ($intentPrompt -notmatch 'interactionMode' -or
                $intentPrompt -notmatch 'requestedOutputs' -or
                $intentPrompt -notmatch '新增.*工作表') {
                throw "Golden case $($case.id): Excel intent prompt cannot preserve open-ended execution requests"
            }

            $officeContext = [ShareRibbon.Agent.Context.OfficeContext]::new()
            $officeContext.AppType = "Excel"
            $officeContext.DocStructure = [ShareRibbon.Agent.Context.DocumentStructure]::new()
            $officeContext.DocStructure.Summary = "当前工作表: 销售数据`r`n表区域: 销售数据!A1:D25; headers=[日期, 区域, 销售员, 销售额]"
            $officeContext.HostData["tables"] = [Newtonsoft.Json.Linq.JArray]::Parse('[{"sheet":"销售数据","address":"A1:D25","headers":["日期","区域","销售员","销售额"]}]')

            $intentContextMethod = [ShareRibbon.IntentRecognitionService].GetMethod(
                "BuildIntentContextInfo",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            if ($null -eq $intentContextMethod) {
                throw "Golden case $($case.id): intent recognition has no testable Office-context bridge"
            }
            $intentSnapshot = [Newtonsoft.Json.Linq.JObject]::new()
            $intentSnapshot["selectionAddress"] = [Newtonsoft.Json.Linq.JValue]::new("I13")
            $intentSnapshot["officeContext"] = [Newtonsoft.Json.Linq.JValue]::new($officeContext.ToPromptText())
            $intentContextArgs = [object[]]::new(1)
            $intentContextArgs[0] = $intentSnapshot
            $intentContextText = [string]$intentContextMethod.Invoke($null, $intentContextArgs)
            if ($intentContextText -notmatch '销售数据!A1:D25' -or $intentContextText -notmatch '销售额') {
                throw "Golden case $($case.id): structured Office context is discarded before intent recognition"
            }

            $chartIntent = [ShareRibbon.IntentResult]::new()
            $chartIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::CHART_GEN
            $chartIntent.IntentType = [ShareRibbon.ExcelIntentType]::CHART_GEN
            $chartIntent.ResponseMode = "clarify"
            $normalizeIntentMethod = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
                "NormalizeIntentForObservedContext",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            if ($null -eq $normalizeIntentMethod) {
                throw "Golden case $($case.id): observed Excel tables cannot deterministically prevent unnecessary chart clarification"
            }
            $normalizeIntentMethod.Invoke($null, [object[]]@($chartIntent, "根据日期和销售额生成折线图", $officeContext, "Excel")) | Out-Null
            if ($chartIntent.ResponseMode -ne "execute" -or -not $chartIntent.RequestedOutputs.Contains("chart")) {
                throw "Golden case $($case.id): an explicit chart request with an observed table still asks the user for ranges"
            }
            $chartAuthorityIntent = & $buildStructuredExcelIntent `
                "根据日期和销售额生成折线图" `
                "execute" `
                ([ShareRibbon.OfficeIntentType]::CHART_GEN) `
                ([ShareRibbon.ExcelIntentType]::CHART_GEN)
            $chartIntent.Confidence = $chartAuthorityIntent.Confidence
            $chartIntent.GoalInterpretation = $chartAuthorityIntent.GoalInterpretation

            $runtime = [ShareRibbon.Agent.AiNativeRuntime]::new([ShareRibbon.Agent.ToolRegistry]::new($null))
            $buildTaskSpec = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
                "BuildTaskSpec",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)

            # Defensive projection: even a directly constructed inconsistent IntentResult cannot
            # turn answer verification metadata into executable TaskSpec work.
            $routingTools = [System.Collections.Generic.List[ShareRibbon.Agent.ToolDescriptor]]::new()
            $routingSkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
            $answerArtifactRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $answerArtifactRequest.UserInput = "请直接回答这个问题"
            $answerArtifactRequest.AppType = "Excel"
            $answerArtifactIntent = [ShareRibbon.IntentResult]::new()
            $answerArtifactIntent.Confidence = 0.95
            $answerArtifactIntent.ResponseMode = "answer"
            $answerArtifactIntent.RequestedOutputs.Add("worksheet")
            $answerArtifactSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($answerArtifactRequest, $answerArtifactIntent, $routingTools, $routingSkills, "Excel"))
            if ($answerArtifactSpec.ExpectedOutputs.Count -ne 0 -or
                $answerArtifactSpec.RequiredTools.Count -ne 0 -or
                [ShareRibbon.IntentAcceptancePolicy]::HasExecutableToolRequirement($answerArtifactSpec) -or
                $answerArtifactSpec.MutationPolicy -ne "read_only") {
                throw "Golden case $($case.id): answer requestedOutputs still create executable TaskSpec work"
            }
            $answerArtifactDecision = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute(
                $false,
                $answerArtifactIntent,
                [ShareRibbon.IntentAcceptancePolicy]::HasExecutableToolRequirement($answerArtifactSpec))
            if ($answerArtifactDecision.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): direct answers still bypass the unified adaptive Agent"
            }

            $invalidRoutingIntent = [ShareRibbon.IntentResult]::new()
            $invalidRoutingIntent.Confidence = 0.95
            $invalidRoutingIntent.ResponseMode = "run"
            $invalidRoutingIntent.RequestedOutputs.Add("worksheet")
            $invalidRoutingSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($answerArtifactRequest, $invalidRoutingIntent, $routingTools, $routingSkills, "Excel"))
            $invalidRoutingDecision = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute(
                $false,
                $invalidRoutingIntent,
                [ShareRibbon.IntentAcceptancePolicy]::HasExecutableToolRequirement($invalidRoutingSpec))
            if ($invalidRoutingSpec.ExpectedOutputs.Count -ne 0 -or
                $invalidRoutingSpec.RequiredTools.Count -ne 0 -or
                $invalidRoutingDecision.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): invalid/clarification output still leaves the unified Agent path"
            }

            # A legitimate answer can still enter AgentKernel when exact workbook evidence needs
            # an actual read tool; the tool requirement, not requestedOutputs, is the authority.
            $readRequirementSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
            $readRequirementSpec.RequiredTools.Add("ReadRange")
            if (-not [ShareRibbon.IntentAcceptancePolicy]::HasExecutableToolRequirement($readRequirementSpec) -or
                [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute(
                    $false,
                    $answerArtifactIntent,
                    $true).ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): a real read requirement can no longer support an exact answer"
            }
            $chartRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $chartRequest.UserInput = "根据日期和销售额生成折线图"
            $chartRequest.AppType = "Excel"
            $chartRequest.OfficeContext = $officeContext
            $chartSkill = [ShareRibbon.SkillFileDefinition]::new()
            $chartSkill.Name = "图表生成"
            $chartSkill.AllowedTools = [System.Collections.Generic.List[string]]::new()
            $chartSkill.AllowedTools.Add("CreateChart")
            $chartSkill.AllowedTools.Add("GenerateReport")
            $chartSkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
            $chartSkills.Add($chartSkill)
            $chartTools = [System.Collections.Generic.List[ShareRibbon.Agent.ToolDescriptor]]::new()
            $chartSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke($runtime, @($chartRequest, $chartIntent, $chartTools, $chartSkills, "Excel"))
            if ($chartSpec.TargetObject -notmatch '销售数据!A1:D25' -or
                $chartSpec.RequiredTools.Count -ne 1 -or
                $chartSpec.RequiredTools[0] -ne "CreateChart") {
                throw "Golden case $($case.id): observed chart task does not target the table with a CreateChart-only gate"
            }

            $routerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb") -Raw -Encoding UTF8
            if ($routerSource -match '_host\.SendChatMessage(?:WithIntent)?\(' -or
                $routerSource -match 'BuildContextAwareChatMessage\(') {
                throw "Golden case $($case.id): smart mode still has a legacy plain/follow-up chat execution branch"
            }

            $answerPromptIntent = [ShareRibbon.IntentResult]::new()
            $answerPromptIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::GENERAL_QUERY
            $answerPromptIntent.ResponseMode = "answer"
            $answerPromptIntent.Confidence = 0.99
            $answerPromptContext = [ShareRibbon.PromptContext]::new()
            $answerPromptContext.ApplicationType = "Excel"
            $answerPromptContext.IntentResult = $answerPromptIntent
            $answerSystemPrompt = [ShareRibbon.PromptManager]::Instance.GetCombinedPrompt($answerPromptContext)
            if ($answerSystemPrompt -match 'Excel支持的\d+个命令' -or
                $answerSystemPrompt -match '必须且只能返回以下两种格式') {
                throw "Golden case $($case.id): a plain capability answer is still contaminated by the legacy JSON command catalog"
            }
        }
        elseif ($case.contractType -eq "chat_visible_output_context_query") {
            $scripts = [System.Collections.Generic.List[string]]::new()
            $state = [ShareRibbon.ChatStateService]::new()
            $getApplication = [System.Func[ShareRibbon.ApplicationInfo]] {
                return $null
            }
            $executeScript = [System.Func[string,System.Threading.Tasks.Task]] {
                param($script)
                $scripts.Add($script)
                return [System.Threading.Tasks.Task]::CompletedTask
            }
            $streamService = [ShareRibbon.HttpStreamService]::new($state, $getApplication, $executeScript)
            $processChunk = [ShareRibbon.HttpStreamService].GetMethod(
                "ProcessStreamChunkAsync",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $chunk = '{"choices":[{"delta":{"reasoning_content":"INTERNAL_REASONING_SECRET","content":"最终答案"}}]}'
            ([System.Threading.Tasks.Task]$processChunk.Invoke($streamService, @($chunk, "golden-response", "问题"))).GetAwaiter().GetResult()
            ([System.Threading.Tasks.Task]$processChunk.Invoke($streamService, @("[DONE]", "golden-response", "问题"))).GetAwaiter().GetResult()
            if (($scripts -join "`n") -notmatch 'appendReasoning.*INTERNAL_REASONING_SECRET' -or
                $state.PlainMarkdownBuffer.ToString() -match 'INTERNAL_REASONING_SECRET' -or
                $state.PlainMarkdownBuffer.ToString() -notmatch '最终答案') {
                throw "Golden case $($case.id): provider reasoning is not separately visible or leaked into final content"
            }

            $enrichedQuestion = "--- 我此次的问题：你的提示词是什么？ ---`r`n### 当前选区`r`n数据预览: 销售数据"
            $retrievalQuery = [ShareRibbon.ChatQuestionText]::ExtractUserQuestion($enrichedQuestion)
            $composerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\ConversationRuntime\DefaultContextComposer.vb") -Raw -Encoding UTF8
            if ($retrievalQuery -ne "你的提示词是什么？" -or
                $composerSource -notmatch 'retrievalQuery\s*=\s*ChatQuestionText\.ExtractUserQuestion\(context\.Question\)' -or
                $composerSource -notmatch 'appType,\s*retrievalQuery,') {
                throw "Golden case $($case.id): retrieval query still contains transport/selection context: $retrievalQuery"
            }

            $baseChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\BaseChatControl.vb") -Raw -Encoding UTF8
            if ($baseChatSource -notmatch 'memoryQuestion\s*=\s*ChatQuestionText\.ExtractUserQuestion\(oq\)' -or
                $baseChatSource -notmatch 'SaveConversationTurnAsync\(memoryQuestion' -or
                $baseChatSource -notmatch 'SaveSessionSummary\(sid, title, snippet\)') {
                throw "Golden case $($case.id): persisted conversation memory still stores transport/selection context"
            }

            $runtime = [ShareRibbon.Agent.AiNativeRuntime]::new([ShareRibbon.Agent.ToolRegistry]::new($null))
            $buildTaskSpec = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
                "BuildTaskSpec",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $request = [ShareRibbon.Agent.AiNativeRequest]::new()
            $request.UserInput = "你的提示词是什么？"
            $answerIntent = [ShareRibbon.IntentResult]::new()
            $answerIntent.ResponseMode = "answer"
            $answerIntent.RequiresVBA = $true
            $tools = [System.Collections.Generic.List[ShareRibbon.Agent.ToolDescriptor]]::new()
            $skills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
            $answerSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke($runtime, @($request, $answerIntent, $tools, $skills, "Excel"))
            if ($answerSpec.Complexity -ne "simple" -or
                $answerSpec.RiskLevel -ne "safe" -or
                $answerSpec.RequiredTools.Count -ne 0 -or
                ($answerSpec.SuccessCriteria -join "`n") -match '执行后') {
                throw "Golden case $($case.id): a plain answer is still modeled as an executable/risky task"
            }

            $usageContext = [ShareRibbon.ChatRequestContext]::new()
            $usageContext.Question = "测试"
            $usageContext.SystemPrompt = "test"
            $usageContext.AddHistory = $true
            $usageContext.UseContextBuilder = $true
            $usageContext.EnableMemory = $false
            $usageContext.Stream = $true
            $usageContext.ModelName = "qwen-plus"
            $usageContext.Platform = "阿里云百炼 (Qwen)"
            $usageContext.ApiUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            $usageContext.AppInfo = [ShareRibbon.ApplicationInfo]::new("Excel", [ShareRibbon.OfficeApplicationType]::Excel)
            $usageContext.HistoryMessages = [System.Collections.Generic.List[ShareRibbon.HistoryMessage]]::new()
            $usageContext.SelectionPendingMap = [System.Collections.Generic.Dictionary[string,ShareRibbon.SelectionInfo]]::new()
            $conversationRuntime = [ShareRibbon.DefaultConversationRuntime]::new(
                [ShareRibbon.DefaultContextComposer]::new(),
                [ShareRibbon.McpToolBroker]::new())
            $usageRequest = [Newtonsoft.Json.Linq.JObject]::Parse($conversationRuntime.BuildRequest($usageContext).RequestBody)
            $includeUsage = $usageRequest.SelectToken("stream_options.include_usage")
            if ($null -eq $includeUsage -or $includeUsage.ToString().ToLowerInvariant() -ne "true") {
                throw "Golden case $($case.id): Qwen streaming usage is not requested, token count will remain zero"
            }
        }
        elseif ($case.contractType -eq "agent_stream_progress") {
            $baseChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\BaseChatControl.vb") -Raw -Encoding UTF8
            $agentServiceSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\AgentKernelService.vb") -Raw -Encoding UTF8
            $agentCardSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Resources\js\agent-card.js") -Raw -Encoding UTF8

            if ($baseChatSource -notmatch 'SendAndGetStreamingResponseAsync' -or
                $baseChatSource -notmatch 'HttpCompletionOption\.ResponseHeadersRead' -or
                $baseChatSource -notmatch 'requestObj\("stream"\)\s*=\s*True' -or
                $baseChatSource -match 'While\s+Not\s+reader\.EndOfStream') {
                throw "Golden case $($case.id): Agent model calls are not using true HTTP streaming"
            }
            if ($agentServiceSource -notmatch 'Await\s+ShowThinkingStatusAsync\(' -or
                $agentServiceSource -notmatch 'QueueUiScriptAsync' -or
                $agentServiceSource -notmatch 'appendAgentStreamDelta') {
                throw "Golden case $($case.id): Agent progress UI is not created and updated in order"
            }
            if ($agentCardSource -notmatch 'function\s+startAgentProgress' -or
                $agentCardSource -notmatch 'function\s+appendAgentStreamDelta' -or
                $agentCardSource -notmatch 'agent-progress-elapsed') {
                throw "Golden case $($case.id): visible Agent progress/heartbeat functions are missing"
            }

            $parserType = [ShareRibbon.AgentStreamChunkParser]
            $parseMethod = $parserType.GetMethod("ParseDataLine", [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Static)
            $delta = $parseMethod.Invoke($null, @('data: {"choices":[{"delta":{"reasoning_content":"正在分析","content":"{"}}]}'))
            if ($delta.ReasoningDelta -ne "正在分析" -or $delta.ContentDelta -ne "{" -or $delta.Done) {
                throw "Golden case $($case.id): Qwen reasoning/content SSE delta parsing failed"
            }
            $done = $parseMethod.Invoke($null, @('data: [DONE]'))
            if (-not $done.Done) {
                throw "Golden case $($case.id): SSE completion marker parsing failed"
            }
        }
        elseif ($case.contractType -eq "approval_api") {
            $interfaceType = [ShareRibbon.Agent.Harness.IOfficeHarness]
            $approve = $interfaceType.GetMethod("ApproveAsync")
            $cancel = $interfaceType.GetMethod("CancelAsync")
            $resume = $interfaceType.GetMethod("ResumeAsync")
            $awaiting = [System.Enum]::Parse([ShareRibbon.Agent.Harness.HarnessRunStatus], "AwaitingApproval")
            if ($null -eq $approve -or $null -eq $cancel -or $null -eq $resume -or $null -eq $awaiting) {
                throw "Golden case failed: $($case.id)"
            }
        }
        elseif ($case.contractType -eq "tool_result_only") {
            $serviceType = [ShareRibbon.CodeExecutionService]
            $legacyMethods = @("ExecuteCodeWithResult", "ExecuteJsonCommand")
            foreach ($methodName in $legacyMethods) {
                if ($null -ne $serviceType.GetMethod($methodName)) {
                    throw "Golden case $($case.id): legacy method still exists: $methodName"
                }
            }
            if ($null -ne $serviceType.GetProperty("JsonCommandExecutor")) {
                throw "Golden case $($case.id): legacy JsonCommandExecutor still exists"
            }

            $sourceFiles = @(
                "ShareRibbon\Controls\BaseChatControl.vb",
                "WordAi\ChatControl.vb",
                "ExcelAi\ChatControl.vb",
                "PowerPointAi\ChatControl.vb"
            )
            foreach ($sourceFile in $sourceFiles) {
                $source = Get-Content -LiteralPath (Join-Path $repoRoot $sourceFile) -Raw -Encoding UTF8
                if ($source -match 'Protected\s+(Overridable|Overrides)\s+Function\s+ExecuteJsonCommand\s*\(') {
                    throw "Golden case $($case.id): Boolean JSON override found in $sourceFile"
                }
            }
        }
        elseif ($case.contractType -eq "powerpoint_create_fallback") {
            $intentService = [ShareRibbon.IntentRecognitionService]::new("PowerPoint")
            $intent = $intentService.IdentifyIntent("ppt")
            if ($intent.OfficeIntent.ToString() -ne "SLIDE_CREATE" -or $intent.Confidence -lt 0.5) {
                throw "Golden case $($case.id): explicit PPT creation request fell back to $($intent.OfficeIntent)"
            }

            $parseMethod = [ShareRibbon.IntentRecognitionService].GetMethod(
                "ParseLLMIntentResponse",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $malformedResult = $parseMethod.Invoke($intentService, @('{"intentType":', 'ppt'))
            if ($null -ne $malformedResult) {
                throw "Golden case $($case.id): malformed LLM JSON returned a default intent and would overwrite the fallback"
            }

            $fallbackIntent = [ShareRibbon.IntentResult]::new()
            $fallbackIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::SLIDE_CREATE
            $fallbackDecision = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute(
                $false,
                $fallbackIntent,
                $true)
            if ($fallbackDecision.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): TaskSpec execution fallback is missing"
            }
        }
        elseif ($case.contractType -eq "powerpoint_com_compat") {
            $classified = [ShareRibbon.ExceptionClassifier]::Classify(
                [InvalidCastException]::new("QueryInterface E_NOINTERFACE 不支持此接口"))
            if ($classified.ErrorCode -ne "COM_ERROR" -or $classified.Recoverable) {
                throw "Golden case $($case.id): E_NOINTERFACE must be terminal"
            }

            $rendererSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\Design\PowerPointSceneRenderer.vb") -Raw -Encoding UTF8
            if ($rendererSource -match 'As\s+(FillFormat|LineFormat|ShadowFormat|ColorFormat|Adjustments|PictureFormat|TextFrame2|TextRange2|Font2|ParagraphFormat2)') {
                throw "Golden case $($case.id): renderer reintroduced a strongly typed Office style interface"
            }

            $pptChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\ChatControl.vb") -Raw -Encoding UTF8
            if ($pptChatSource -notmatch 'ProfessionalDeckExecutor\.ExecuteAsToolResult\(\s*GetPowerPointEnvelopeParams\(envelope\),\s*preview\)') {
                throw "Golden case $($case.id): CreateSlides preview flag is not propagated"
            }
        }
        elseif ($case.contractType -eq "excel_observer_compat") {
            $excelChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw -Encoding UTF8
            if ($excelChatSource -match 'BuildExcelRangePreview\(target\s+As\s+Microsoft\.Office\.Interop\.Excel\.Range' -or
                $excelChatSource -match 'CountExcelFormulaErrors\(target\s+As\s+Microsoft\.Office\.Interop\.Excel\.Range' -or
                $excelChatSource -match 'TryCast\(app\.Selection,\s*Microsoft\.Office\.Interop\.Excel\.Range\)') {
                throw "Golden case $($case.id): Excel observer reintroduced a strongly typed Range cast"
            }
            if ($excelChatSource -notmatch 'CodeObservationFailed' -or
                $excelChatSource -notmatch 'recoverable:=False') {
                throw "Golden case $($case.id): observer failures can still repeat an uncertain write"
            }
            if ($excelChatSource -notmatch 'sheetCountDelta' -or
                $excelChatSource -notmatch 'chartCountDelta' -or
                $excelChatSource -notmatch 'formulaErrorDelta') {
                throw "Golden case $($case.id): Excel observer is missing basic workbook diff metrics"
            }
            if ($excelChatSource -match 'snapshot\("workbook"\)\s*=\s*If\(workbook\.Name' -or
                $excelChatSource -match 'snapshot\("worksheet"\)\s*=\s*If\(worksheet\.Name' -or
                $excelChatSource -notmatch 'CStr\(If\(workbook\.Name' -or
                $excelChatSource -notmatch 'CStr\(If\(worksheet\.Name') {
                throw "Golden case $($case.id): late-bound Excel names can still be cast from String to JToken during observation"
            }
            if ($excelChatSource -match 'target\s*=\s*worksheet\.Range\(targetAddress\)' -or
                $excelChatSource -notmatch 'snapshot\("targetWorksheet"\)' -or
                $excelChatSource -notmatch 'targetWorksheet\.Range\(targetCellAddress\)') {
                throw "Golden case $($case.id): cross-sheet target observation still uses the active worksheet and can raise 0x800A03EC"
            }
            if ($excelChatSource -notmatch 'ExpandExcelObservationTarget' -or
                $excelChatSource -notmatch 'targetAnchor\.Resize\(rows\.Count, firstRow\.Count\)' -or
                $excelChatSource -notmatch 'targetAnchor\.CurrentRegion' -or
                $excelChatSource -notmatch 'params\?\("targetSheet"\)') {
                throw "Golden case $($case.id): output-anchor observation does not cover cross-sheet WriteData/DataAnalysis result ranges"
            }

            $agentPromptSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\PromptManager.vb") -Raw -Encoding UTF8
            $excelSkillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw -Encoding UTF8
            $toolRegistrySource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\ToolRegistry.vb") -Raw -Encoding UTF8
            if ($agentPromptSource -notmatch '纯新建.*CreateSheet' -or
                $agentPromptSource -notmatch '工作表名称.*不代表.*报表' -or
                $excelSkillSource -notmatch 'worksheet name.*does not imply.*report') {
                throw "Golden case $($case.id): plain worksheet creation is not protected from GenerateReport overreach"
            }
            if ($toolRegistrySource -notmatch 'CreateSheet.*sheetName' -or
                $agentPromptSource -notmatch '显式指定.*Python.*不得替换' -or
                $excelSkillSource -notmatch 'explicitly requests.*PythonCompute.*must not substitute') {
                throw "Golden case $($case.id): CreateSheet aliases or explicit Python execution semantics are not protected"
            }

            $normalizingRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
            $normalizingRegistry.LoadFromDirectory($toolsPath)
            $sheetNameAliasCall = [ShareRibbon.Agent.ToolCall]::new()
            $sheetNameAliasCall.ToolId = "CreateSheet"
            $sheetNameAliasCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse('{"sheetName":"验证汇总0821"}')
            $normalizationMessage = ""
            if (-not $normalizingRegistry.TryNormalizeToolCall("Excel", $sheetNameAliasCall, [ref]$normalizationMessage) -or
                $sheetNameAliasCall.Parameters["name"].ToString() -ne "验证汇总0821") {
                throw "Golden case $($case.id): CreateSheet sheetName alias is not normalized to the executable name parameter"
            }

            $runtime = [ShareRibbon.Agent.AiNativeRuntime]::new([ShareRibbon.Agent.ToolRegistry]::new($null))
            $buildTaskSpec = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
                "BuildTaskSpec",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $destinationCorrectionMethod = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
                "IsDestinationOnlyCorrection",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::NonPublic)
            if ([bool]$destinationCorrectionMethod.Invoke($null, @('在工作表“ABC”中写入数据。'))) {
                throw "Golden case $($case.id): a standalone write request is still inheriting the previous Python task"
            }
            if ([bool]$destinationCorrectionMethod.Invoke($null, @('在之前创建的ABC工作表中写入数据。'))) {
                throw "Golden case $($case.id): referring to an existing sheet is still mistaken for correcting a previous task"
            }
            if (-not [bool]$destinationCorrectionMethod.Invoke($null, @('我刚刚说的新工作表改为之前创建的验证汇总0821'))) {
                throw "Golden case $($case.id): an explicit destination correction is no longer recognized"
            }

            $requiredWorksheetTools = @("CreateSheet", "DeleteSheet", "RenameSheet", "CopySheet", "InsertRowCol", "DeleteRowCol", "HideRowCol", "ProtectSheet")
            foreach ($requiredWorksheetTool in $requiredWorksheetTools) {
                if ($excelSkillSource -notmatch ('allowed-tools:[^\r\n]*' + [regex]::Escape($requiredWorksheetTool))) {
                    throw "Golden case $($case.id): Excel skill does not expose migrated worksheet tool $requiredWorksheetTool"
                }
            }
            if ($agentPromptSource -notmatch 'OfficeObjectOperation.*DiscoverOfficeCapability' -or
                $excelSkillSource -notmatch 'long-tail Excel object operation not covered by a high-level tool') {
                throw "Golden case $($case.id): generic Office operations are not constrained to discovered long-tail capabilities"
            }

            $sheetRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $sheetRequest.UserInput = "新增一个名为汇总的工作表"
            $sheetIntent = & $buildStructuredExcelIntent $sheetRequest.UserInput
            $sheetIntent.RequestedOutputs.Add("worksheet")
            $skill = [ShareRibbon.SkillFileDefinition]::new()
            $skill.Name = "Excel Table Agent"
            $skill.AllowedTools = [System.Collections.Generic.List[string]]::new()
            $skill.AllowedTools.Add("CreateSheet")
            $skill.AllowedTools.Add("GenerateReport")
            $skills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
            $skills.Add($skill)
            $tools = [System.Collections.Generic.List[ShareRibbon.Agent.ToolDescriptor]]::new()
            $sheetSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke($runtime, @($sheetRequest, $sheetIntent, $tools, $skills, "Excel"))
            if ($sheetSpec.RequiredTools.Count -ne 1 -or $sheetSpec.RequiredTools[0] -ne "CreateSheet") {
                throw "Golden case $($case.id): plain worksheet creation is not deterministically scoped to CreateSheet"
            }

            $session = [ShareRibbon.Agent.AgentSession]::new($sheetRequest.UserInput, "Excel", "")
            $session.Spec = $sheetSpec
            $agentSkill = [ShareRibbon.Agent.AgentSkill]::new()
            $agentSkill.Name = "Excel Table Agent"
            $agentSkill.RequiredTools.Add("CreateSheet")
            $agentSkill.RequiredTools.Add("GenerateReport")
            $toolExecutionContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession($session, $agentSkill)
            if (-not $toolExecutionContext.AllowedTools.Contains("CreateSheet") -or
                -not $toolExecutionContext.AllowedTools.Contains("GenerateReport")) {
                throw "Golden case $($case.id): TaskSpec planner hints are still being mistaken for the Skill authorization boundary"
            }

            # A read-only answer that depends on live workbook values is still an agent task.
            # The semantic intent and observed table drive this contract; the particular wording
            # of the question must not become a production routing branch.
            $readOnlyRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $readOnlyRequest.UserInput = "请根据当前数据给出精确统计结果，只回答且不要修改工作簿"
            $readOnlyRequest.AppType = "Excel"
            $readOnlyRequest.OfficeContext = [ShareRibbon.Agent.Context.OfficeContext]::new()
            $readOnlyRequest.OfficeContext.AppType = "Excel"
            $readOnlyRequest.OfficeContext.HostData["tables"] = [Newtonsoft.Json.Linq.JArray]::Parse('[{"sheet":"任意数据表","address":"A1:D25","headers":["字段一","字段二","字段三","数值"]}]')
            $readOnlyIntent = [ShareRibbon.IntentResult]::new()
            $readOnlyIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::DATA_ANALYSIS
            $readOnlyIntent.IntentType = [ShareRibbon.ExcelIntentType]::DATA_ANALYSIS
            $readOnlyIntent.ResponseMode = "answer"
            $readOnlySkill = [ShareRibbon.SkillFileDefinition]::new()
            $readOnlySkill.Name = "Excel Table Agent"
            $readOnlySkill.AllowedTools = [System.Collections.Generic.List[string]]::new()
            foreach ($toolId in @("ReadRange", "WriteData", "DataAnalysis", "PythonCompute")) {
                $readOnlySkill.AllowedTools.Add($toolId)
            }
            $readOnlySkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
            $readOnlySkills.Add($readOnlySkill)
            $readOnlySpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke($runtime, @($readOnlyRequest, $readOnlyIntent, $tools, $readOnlySkills, "Excel"))
            if ($readOnlySpec.RequiredTools.Count -ne 1 -or $readOnlySpec.RequiredTools[0] -ne "ReadRange" -or
                $readOnlySpec.MutationPolicy -ne "read_only" -or
                ($readOnlySpec.SuccessCriteria -join "`n") -notmatch '禁止估算') {
                throw "Golden case $($case.id): read-only data answers are not constrained to exact ReadRange evidence"
            }

            # Interaction mode is a model hint, not authority to mutate a workbook.  A data
            # question with no explicit write request remains read-only even if the intent
            # model incorrectly labels it as execute.  Conversely an explicit destination
            # keeps write authority.
            $misclassifiedQuestion = [ShareRibbon.Agent.AiNativeRequest]::new()
            $misclassifiedQuestion.UserInput = "分别有多少个人的销售记录？"
            $misclassifiedQuestion.AppType = "Excel"
            $misclassifiedQuestion.OfficeContext = $readOnlyRequest.OfficeContext
            $misclassifiedIntent = [ShareRibbon.IntentResult]::new()
            $misclassifiedIntent.OfficeIntent = [ShareRibbon.OfficeIntentType]::DATA_ANALYSIS
            $misclassifiedIntent.IntentType = [ShareRibbon.ExcelIntentType]::DATA_ANALYSIS
            $misclassifiedIntent.ResponseMode = "execute"
            $misclassifiedSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($misclassifiedQuestion, $misclassifiedIntent, $tools, $readOnlySkills, "Excel"))
            if ($misclassifiedSpec.MutationPolicy -ne "read_only" -or
                -not $misclassifiedSpec.RequiredTools.Contains("ReadRange")) {
                throw "Golden case $($case.id): an intent-model execute misclassification granted workbook mutation authority to a data question"
            }

            $explicitWriteRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $explicitWriteRequest.UserInput = "按销售员统计记录数，并把结果写入当前工作表M1"
            $explicitWriteRequest.AppType = "Excel"
            $explicitWriteRequest.OfficeContext = $readOnlyRequest.OfficeContext
            $unstructuredWriteSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($explicitWriteRequest, $misclassifiedIntent, $tools, $readOnlySkills, "Excel"))
            if ($unstructuredWriteSpec.MutationPolicy -ne "read_only") {
                throw "Golden case $($case.id): raw command wording bypassed the structured Goal authority boundary"
            }

            $explicitWriteIntent = & $buildStructuredExcelIntent `
                $explicitWriteRequest.UserInput `
                "execute" `
                ([ShareRibbon.OfficeIntentType]::DATA_ANALYSIS) `
                ([ShareRibbon.ExcelIntentType]::DATA_ANALYSIS)
            $explicitWriteSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($explicitWriteRequest, $explicitWriteIntent, $tools, $readOnlySkills, "Excel"))
            if ($explicitWriteSpec.MutationPolicy -eq "read_only" -or
                -not $explicitWriteSpec.RequiresHostExecution) {
                throw "Golden case $($case.id): a validated execute Goal did not grant task-level mutation authority"
            }

            # A terse formatting directive is still explicit user authority when it is grounded
            # in the active Excel selection and the structured intent preserves the exact text.
            # The model cannot grant mutation authority to a question or a negated directive.
            $contextualFormatContext = [ShareRibbon.Agent.Context.OfficeContext]::new()
            $contextualFormatContext.AppType = "Excel"
            $contextualFormatContext.Selection = [ShareRibbon.Agent.Context.SelectionInfo]::new()
            $contextualFormatContext.Selection.Address = "T2"
            $contextualFormatContext.Selection.ItemCount = 1
            $contextualFormatContext.HostData["tables"] = [Newtonsoft.Json.Linq.JArray]::Parse('[{"sheet":"销售数据","address":"A1:T43","headers":["毛利率"]}]')

            $contextualFormatSpec = $null
            foreach ($contextualMutationText in @("用百分比的格式", "照上一列那样")) {
                $contextualFormatRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
                $contextualFormatRequest.UserInput = $contextualMutationText
                $contextualFormatRequest.AppType = "Excel"
                $contextualFormatRequest.OfficeContext = $contextualFormatContext
                $contextualFormatIntent = & $buildStructuredExcelIntent `
                    $contextualMutationText `
                    "execute" `
                    ([ShareRibbon.OfficeIntentType]::FORMAT_STYLE) `
                    ([ShareRibbon.ExcelIntentType]::FORMAT_STYLE)
                $contextualFormatSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                    $runtime,
                    @($contextualFormatRequest, $contextualFormatIntent, $tools, $readOnlySkills, "Excel"))
                if ($contextualFormatSpec.MutationPolicy -eq "read_only" -or
                    -not $contextualFormatSpec.RequiresHostExecution) {
                    throw "Golden case $($case.id): a grounded execute Goal was downgraded to read-only: $contextualMutationText"
                }
            }

            $contextualFormatSession = [ShareRibbon.Agent.AgentSession]::new("照上一列那样", "Excel", "")
            $contextualFormatSession.Spec = $contextualFormatSpec
            $contextualFormatSkill = [ShareRibbon.Agent.AgentSkill]::new()
            $contextualFormatSkill.Name = "Excel Table Agent"
            foreach ($toolId in $readOnlySkill.AllowedTools) {
                $contextualFormatSkill.RequiredTools.Add($toolId)
            }
            $contextualFormatExecutionContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession(
                $contextualFormatSession,
                $contextualFormatSkill)
            $contextualVisibleToolIds = @($normalizingRegistry.GetVisibleTools("Excel", $contextualFormatExecutionContext) | ForEach-Object { $_.Id })
            if ($contextualVisibleToolIds -notcontains "FormatRange" -or
                $contextualVisibleToolIds -notcontains "OfficeObjectOperation") {
                throw "Golden case $($case.id): an authorized formatting directive cannot see equivalent registered mutation capabilities"
            }

            $partialGoalRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $partialGoalRequest.UserInput = "新增一列毛利率，并计算数据填充"
            $partialGoalRequest.AppType = "Excel"
            $partialGoalRequest.OfficeContext = $contextualFormatContext
            $partialGoalIntent = & $buildStructuredExcelIntent `
                $partialGoalRequest.UserInput `
                "execute" `
                ([ShareRibbon.OfficeIntentType]::FORMULA_CALC) `
                ([ShareRibbon.ExcelIntentType]::FORMULA_CALC)
            $partialGoalIntent.GoalInterpretation.Candidate.SourceClauses[0].Text = "新增一列毛利率"
            $partialGoalIntent.GoalInterpretation.Candidate.Criteria[0].Statement = "新增一列毛利率"
            $partialGoalSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($partialGoalRequest, $partialGoalIntent, $tools, $readOnlySkills, "Excel"))
            if ($partialGoalSpec.MutationPolicy -eq "read_only" -or
                -not $partialGoalSpec.RequiresHostExecution -or
                [string]::IsNullOrWhiteSpace($partialGoalSpec.GoalInterpretationFallbackReason)) {
                throw "Golden case $($case.id): rejected structured Goal did not fall back to the complete raw execute request before authorization"
            }

            foreach ($nonMutationText in @("百分比格式是什么意思？", "如何设置百分比格式？", "不要用百分比格式")) {
                $nonMutationRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
                $nonMutationRequest.UserInput = $nonMutationText
                $nonMutationRequest.AppType = "Excel"
                $nonMutationRequest.OfficeContext = $contextualFormatContext
                $nonMutationIntent = & $buildStructuredExcelIntent `
                    $nonMutationText `
                    "answer" `
                    ([ShareRibbon.OfficeIntentType]::FORMAT_STYLE) `
                    ([ShareRibbon.ExcelIntentType]::FORMAT_STYLE)
                $nonMutationSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                    $runtime,
                    @($nonMutationRequest, $nonMutationIntent, $tools, $readOnlySkills, "Excel"))
                if ($nonMutationSpec.MutationPolicy -ne "read_only") {
                    throw "Golden case $($case.id): a question or negated formatting phrase gained workbook mutation authority: $nonMutationText"
                }
            }

            $readOnlyFormatSession = [ShareRibbon.Agent.AgentSession]::new("百分比格式是什么意思？", "Excel", "")
            $readOnlyFormatSession.Spec = $nonMutationSpec
            $readOnlyFormatExecutionContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession(
                $readOnlyFormatSession,
                $contextualFormatSkill)
            $deniedFormatCall = [ShareRibbon.Agent.ToolCall]::new()
            $deniedFormatCall.ToolId = "FormatRange"
            $deniedFormatMessage = ""
            $deniedFormatCode = ""
            if ($normalizingRegistry.TryNormalizeToolCall(
                    "Excel",
                    $deniedFormatCall,
                    $readOnlyFormatExecutionContext,
                    [ref]$deniedFormatMessage,
                    [ref]$deniedFormatCode) -or
                $deniedFormatCode -ne "TOOL_NOT_ALLOWED" -or
                $deniedFormatMessage -notmatch "已注册") {
                throw "Golden case $($case.id): a policy-hidden registered tool is still reported as NOT_FOUND"
            }

            # A compound request is still one adaptive Agent goal. Its exact structured Goal,
            # not a keyword batcher, supplies the task-level mutation authority.
            $compoundMutationRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $compoundMutationRequest.UserInput = "按销售额从高到低排序整个数据表。`n只显示区域为华东的记录。`n删除完全重复的数据行。`n替换销售员姓名。`n清理多余空格。"
            $compoundMutationRequest.AppType = "Excel"
            $compoundMutationRequest.OfficeContext = $readOnlyRequest.OfficeContext
            $compoundMutationIntent = & $buildStructuredExcelIntent `
                $compoundMutationRequest.UserInput `
                "execute" `
                ([ShareRibbon.OfficeIntentType]::DATA_CLEANING) `
                ([ShareRibbon.ExcelIntentType]::DATA_CLEANING)
            $compoundMutationSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke(
                $runtime,
                @($compoundMutationRequest, $compoundMutationIntent, $tools, $readOnlySkills, "Excel"))
            $compoundMutationRequiresAgent = [ShareRibbon.IntentAcceptancePolicy]::HasExecutableToolRequirement($compoundMutationSpec)
            $compoundMutationRoute = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute(
                $false,
                $compoundMutationIntent,
                $compoundMutationRequiresAgent)
            if (-not $compoundMutationRequiresAgent -or $compoundMutationRoute.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): explicit compound workbook mutations were routed to legacy chat batching"
            }
            $readOnlyRoute = [ShareRibbon.ChatRoutingOrchestrator]::DecidePostAnalysisRoute($false, $readOnlyIntent, $true)
            if ($readOnlyRoute.ToString() -ne "AgentKernel") {
                throw "Golden case $($case.id): a workbook-data answer requiring ReadRange was routed to plain chat"
            }

            $loopSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\LoopEngine.vb") -Raw -Encoding UTF8
            $agentServiceSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\AgentKernelService.vb") -Raw -Encoding UTF8
            if ($loopSource -match 'TryCompleteObservedGoal\(' -or
                $loopSource -notmatch 'Dim finalOutput As String = completionMessage' -or
                $agentServiceSource -notmatch 'result\?\.FinalOutput') {
                throw "Golden case $($case.id): a tool result can still bypass the next model turn or the model's final response is not delivered"
            }
            $appendEvidence = [ShareRibbon.Agent.LoopEngine].GetMethod(
                "AppendReadOnlyEvidence",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            $evidence = [Newtonsoft.Json.Linq.JArray]::new()
            $readCall = [ShareRibbon.Agent.ToolCall]::new()
            $readCall.ToolId = "ReadRange"
            $readCall.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse('{"range":"任意数据表!A1:B4"}')
            $readData = [Newtonsoft.Json.Linq.JObject]::Parse('{"sheet":"任意数据表","address":"A1:B4","rowCount":4,"columnCount":2,"values":[["类别","数量"],["甲",1],["乙",2],["甲",3]]}')
            $readObservation = [Newtonsoft.Json.Linq.JObject]::Parse('{"kind":"read","summary":"read","changed":false}')
            $readResult = [ShareRibbon.Agent.ToolResult]::Succeed("ReadRange", "read", $readData, $readObservation)
            $evidenceError = [string]$appendEvidence.Invoke($null, @($evidence, $readCall, $readResult))
            if (-not [string]::IsNullOrWhiteSpace($evidenceError) -or
                $evidence.Count -ne 1 -or
                $evidence[0]["data"]["values"].Count -ne 4) {
                throw "Golden case $($case.id): full ReadRange values are not preserved for final answer synthesis"
            }
            $loop = [ShareRibbon.Agent.LoopEngine]::new(
                [ShareRibbon.Agent.ToolRegistry]::new($null),
                [ShareRibbon.Agent.AgentMemory]::new(),
                [ShareRibbon.Agent.PromptManager]::new((Join-Path $repoRoot "ShareRibbon\Prompts")))

            # Conversational delivery is not an Office host-state outcome. Once required
            # read/compute capabilities succeeded and structured evidence was captured, a
            # model completion decision must be accepted without inventing an OutcomeContract.
            $completionSession = [ShareRibbon.Agent.AgentSession]::new("所以几个人呀？", "Excel", "")
            $completionSession.Spec = $readOnlySpec
            $successfulReadIteration = [ShareRibbon.Agent.ReActIteration]::new()
            $successfulReadIteration.Action = $readCall
            $successfulReadIteration.Explanation = [ShareRibbon.Agent.ExecutionExplanation]::new()
            $successfulReadIteration.Explanation.Success = $true
            $completionSession.Iterations.Add($successfulReadIteration)
            $parseDecision = [ShareRibbon.Agent.LoopEngine].GetMethod(
                "ParseReactDecision",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $completionDecision = $parseDecision.Invoke($loop, @('{"decision":"complete","message":"已得到精确计数"}'))
            $acceptCompletion = [ShareRibbon.Agent.LoopEngine].GetMethod(
                "TryAcceptCompletionDecision",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $completionArgs = [object[]]@($completionSession, $completionDecision, 1, "")
            $completionAccepted = [bool]$acceptCompletion.Invoke($loop, $completionArgs)
            if (-not $completionAccepted -or $completionArgs[3] -ne "已得到精确计数") {
                throw "Golden case $($case.id): a read-only answer with exact successful evidence was forced through Office OutcomeContract verification"
            }

            # The unified Agent also owns ordinary conversational turns. Such a turn may
            # complete directly without fabricating Office evidence, while remaining unable
            # to mutate the workbook.
            $directAnswerSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
            $directAnswerSpec.Goal = "你好"
            $directAnswerSpec.MutationPolicy = "read_only"
            $directAnswerSpec.RequiresHostExecution = $false
            $directAnswerSession = [ShareRibbon.Agent.AgentSession]::new("你好", "Excel", "")
            $directAnswerSession.Spec = $directAnswerSpec
            $directAnswerDecision = $parseDecision.Invoke($loop, @('{"decision":"complete","message":"你好，请问需要我处理什么？"}'))
            $directAnswerArgs = [object[]]@($directAnswerSession, $directAnswerDecision, 0, "")
            if (-not [bool]$acceptCompletion.Invoke($loop, $directAnswerArgs) -or
                $directAnswerArgs[3] -ne "你好，请问需要我处理什么？") {
                throw "Golden case $($case.id): a direct answer cannot complete inside the unified adaptive Agent"
            }

            $emptyCompletionDecision = $parseDecision.Invoke($loop, @('{"decision":"complete","thought":"calculation finished","message":""}'))
            $emptyCompletionArgs = [object[]]@($completionSession, $emptyCompletionDecision, 1, "")
            if ([bool]$acceptCompletion.Invoke($loop, $emptyCompletionArgs)) {
                throw "Golden case $($case.id): an empty control-only completion was accepted without a user-facing answer"
            }

            $genericContext = [ShareRibbon.Agent.ToolExecutionContext]::new()
            $genericContext.AllowedTools = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $genericContext.AllowedTools.Add("DiscoverOfficeCapability") | Out-Null
            $genericContext.AllowedTools.Add("OfficeObjectOperation") | Out-Null
            $genericContext.EnforceAllowedTools = $true
            if ($genericContext.IsOfficeObjectOperationReady()) {
                throw "Golden case $($case.id): OfficeObjectOperation is executable before capability discovery"
            }
            $genericContext.RecordSuccessfulTool("DiscoverOfficeCapability")
            if (-not $genericContext.IsOfficeObjectOperationReady()) {
                throw "Golden case $($case.id): successful capability discovery did not unlock OfficeObjectOperation"
            }

            $pythonRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $pythonRequest.UserInput = "按区域汇总销售额，使用 Python 计算，并把结果写入新的工作表"
            $pythonRequest.AppType = "Excel"
            $pythonRequest.OfficeContext = [ShareRibbon.Agent.Context.OfficeContext]::new()
            $pythonRequest.OfficeContext.AppType = "Excel"
            $pythonRequest.OfficeContext.HostData["tables"] = [Newtonsoft.Json.Linq.JArray]::Parse('[{"sheet":"销售数据","address":"A1:D25","headers":["日期","区域","销售员","销售额"]}]')
            $pythonIntent = & $buildStructuredExcelIntent `
                $pythonRequest.UserInput `
                "execute" `
                ([ShareRibbon.OfficeIntentType]::DATA_ANALYSIS) `
                ([ShareRibbon.ExcelIntentType]::DATA_ANALYSIS)
            $pythonIntent.RequestedOutputs.Add("worksheet")
            $pythonSkill = [ShareRibbon.SkillFileDefinition]::new()
            $pythonSkill.Name = "Excel Table Agent"
            $pythonSkill.AllowedTools = [System.Collections.Generic.List[string]]::new()
            foreach ($toolId in @("ReadRange", "PythonCompute", "CreateSheet", "WriteData", "DataAnalysis", "GenerateReport")) {
                $pythonSkill.AllowedTools.Add($toolId)
            }
            $pythonSkills = [System.Collections.Generic.List[ShareRibbon.SkillFileDefinition]]::new()
            $pythonSkills.Add($pythonSkill)
            $pythonSpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke($runtime, @($pythonRequest, $pythonIntent, $tools, $pythonSkills, "Excel"))
            $expectedPythonTools = @("ReadRange", "PythonCompute", "CreateSheet", "WriteData")
            if ($pythonSpec.RequiredTools.Count -ne $expectedPythonTools.Count -or
                @($expectedPythonTools | Where-Object { -not $pythonSpec.RequiredTools.Contains($_) }).Count -gt 0 -or
                $pythonSpec.RequiredTools.Contains("DataAnalysis") -or
                $pythonSpec.RequiredTools.Contains("GenerateReport")) {
                throw "Golden case $($case.id): explicit Python task can still be silently replaced by native analysis/report tools"
            }

            $continuityRequest = [ShareRibbon.Agent.AiNativeRequest]::new()
            $continuityRequest.UserInput = "额，我刚刚说的新的工作表指的是之前创建的验证汇总0821这个"
            $continuityRequest.AppType = "Excel"
            $continuityRequest.PreviousTaskSpec = $pythonSpec
            $continuityRequest.OfficeContext = [ShareRibbon.Agent.Context.OfficeContext]::new()
            $continuityRequest.OfficeContext.AppType = "Excel"
            $continuityRequest.OfficeContext.HostData["tables"] = [Newtonsoft.Json.Linq.JArray]::Parse('[{"sheet":"区域销售额汇总","address":"A1:D25","headers":["分组汇总","区域","销售员","销售额"]}]')
            $continuityIntent = [ShareRibbon.IntentResult]::new()
            $continuityIntent.ResponseMode = "clarify"
            $continuityMethod = [ShareRibbon.Agent.AiNativeRuntime].GetMethod(
                "NormalizeIntentForTaskContinuity",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            $continuityMethod.Invoke($null, [object[]]@($continuityIntent, $continuityRequest, "Excel")) | Out-Null
            if ($continuityIntent.ResponseMode -ne "execute") {
                throw "Golden case $($case.id): destination-only Python follow-up can still fall back to clarification/chat"
            }
            $continuitySpec = [ShareRibbon.Agent.AgentTaskSpec]$buildTaskSpec.Invoke($runtime, @($continuityRequest, $continuityIntent, $tools, $pythonSkills, "Excel"))
            $expectedContinuityTools = @("ReadRange", "PythonCompute", "WriteData")
            if ($continuitySpec.TargetObject -notmatch '销售数据!A1:D25' -or
                $continuitySpec.RequiredTools.Count -ne $expectedContinuityTools.Count -or
                @($expectedContinuityTools | Where-Object { -not $continuitySpec.RequiredTools.Contains($_) }).Count -gt 0 -or
                ($continuitySpec.Constraints -join "`n") -notmatch '不得把当前输出表改为数据源') {
                throw "Golden case $($case.id): destination correction does not preserve the prior Python source/method or still creates another sheet. target=$($continuitySpec.TargetObject); tools=$($continuitySpec.RequiredTools -join ','); constraints=$($continuitySpec.Constraints -join ' | ')"
            }
        }
        else {
            $observation = [Newtonsoft.Json.Linq.JObject]::Parse('{"kind":"write","summary":"changed","changed":true}')
            $result = [ShareRibbon.Agent.ToolResult]::Succeed("contract.test", "ok", $null, $observation)
            if (-not $result.Success -or $null -eq $result.Observation -or [string]::IsNullOrWhiteSpace($result.ToObserveSummary())) {
                throw "Golden case failed: $($case.id)"
            }
        }
        $results += [pscustomobject]@{ Id = $case.id; Status = "PASS"; ErrorCode = ""; HostCalls = 0 }
        continue
    }

    $registry = [ShareRibbon.Agent.ToolRegistry]::new($null)
    $registry.LoadFromDirectory($toolsPath)
    $hostCalls = [System.Collections.Generic.List[string]]::new()
    $registry.ExecuteCodeWithToolResult = [System.Func[string,string,bool,ShareRibbon.Agent.ToolResult]] {
        param($code, $language, $preview)
        $hostCalls.Add($code)
        return [ShareRibbon.Agent.ToolResult]::Succeed("fake-host", "executed")
    }

    $context = New-Object ShareRibbon.Agent.ToolExecutionContext
    $context.AppType = $case.appType
    $context.RunId = $case.id
    $context.CorrelationId = $case.id
    $context.PrimarySkillName = "golden-l0"
    $context.AllowedTools = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($tool in $case.allowedTools) {
        [void]$context.AllowedTools.Add([string]$tool)
    }
    $context.EnforceAllowedTools = $true
    if ($null -ne $case.successfulPriorTools) {
        foreach ($tool in $case.successfulPriorTools) {
            $context.RecordSuccessfulTool([string]$tool)
        }
    }
    if ($null -ne $case.approvedTools) {
        foreach ($tool in $case.approvedTools) {
            $context.ApproveTool([string]$tool)
        }
    }

    $params = New-Object Newtonsoft.Json.Linq.JObject
    if ($null -ne $case.params) {
        # JObject.FromObject(PSCustomObject) can serialize PowerShell's CliXml adapter
        # instead of the JSON properties. Round-trip through JSON so the runtime receives
        # the same parameter shape that the model/provider sends in production.
        $paramsJson = $case.params | ConvertTo-Json -Depth 30 -Compress
        $params = [Newtonsoft.Json.Linq.JObject]::Parse($paramsJson)
    }

    $task = $registry.ExecuteToolAsync($context, [string]$case.toolId, $params)
    $result = $task.GetAwaiter().GetResult()
    $expectedSuccess = ($case.expectedSuccess -eq $true)
    if ($result.Success -ne $expectedSuccess) {
        throw "Golden case $($case.id) expected success=$expectedSuccess, got $($result.Success)"
    }
    if (-not $expectedSuccess -and $result.ErrorCode -ne $case.expectedErrorCode) {
        throw "Golden case $($case.id) expected $($case.expectedErrorCode), got $($result.ErrorCode)"
    }
    if ($null -ne $case.expectedRecoverable -and $result.Recoverable -ne [bool]$case.expectedRecoverable) {
        throw "Golden case $($case.id) expected recoverable=$($case.expectedRecoverable), got $($result.Recoverable)"
    }
    if ($hostCalls.Count -ne [int]$case.expectedHostCalls) {
        throw "Golden case $($case.id) expected host calls $($case.expectedHostCalls), got $($hostCalls.Count)"
    }
    if ($case.id -eq "S-approval-resume") {
        $second = $registry.ExecuteToolAsync($context, [string]$case.toolId, $params).GetAwaiter().GetResult()
        if ($second.Success -or $second.ErrorCode -ne "SAFETY_NEEDS_APPROVAL" -or $hostCalls.Count -ne 1) {
            throw "Golden case $($case.id) approval token was not single-use"
        }
    }

    $results += [pscustomobject]@{
        Id = $case.id
        Status = "PASS"
        ErrorCode = $result.ErrorCode
        HostCalls = $hostCalls.Count
    }
}

$results | Format-Table -AutoSize
Write-Host "Golden L0 passed: $($results.Count) cases"
