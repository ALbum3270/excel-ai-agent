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

# Ordered workflows are contracts as well as sets. A plan containing every tool in the
# wrong dependency order must be rejected before touching the workbook.
$orderedSpec = [ShareRibbon.Agent.AgentTaskSpec]::new()
foreach ($id in @("ReadRange", "PythonCompute", "WriteData")) {
    $orderedSpec.MandatoryTools.Add($id)
    $orderedSpec.MandatoryToolSequence.Add($id)
}
$wrongOrderPlan = [ShareRibbon.Agent.ExecutionPlan]::new()
foreach ($id in @("PythonCompute", "ReadRange", "WriteData")) {
    $wrongOrderPlan.Steps.Add([ShareRibbon.Agent.PlanStep]@{
        Code = "{`"command`":`"$id`",`"params`":{}}"
        Language = "json"
    })
}
if ([string]::IsNullOrWhiteSpace(
    [ShareRibbon.Agent.AgentExecutionContract]::ValidatePlan($orderedSpec, $wrongOrderPlan))) {
    throw "Mandatory tool sequence accepted an invalid dependency order"
}

$session = [ShareRibbon.Agent.AgentSession]::new("compute", "Excel", "")
$session.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$session.Spec.RequiredTools.Add("ReadRange")
$session.Spec.RequiredTools.Add("PythonCompute")
$session.Spec.RequiredTools.Add("WriteData")
$toolContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession($session, $skill)
$hiddenByTaskSpec = @($availableExcelToolIds | Where-Object { -not $toolContext.IsToolAllowed($_) })
if ($hiddenByTaskSpec.Count -gt 0) {
    throw "TaskSpec still hides Skill-authorized Excel tools: $($hiddenByTaskSpec -join ', ')"
}

# Read-only mutation policy remains a hard boundary and may narrow the Skill to the
# explicitly requested read tools.
$session.Spec.MutationPolicy = "read_only"
$session.Spec.RequiredTools.Clear()
$session.Spec.RequiredTools.Add("ReadRange")
$readOnlyContext = [ShareRibbon.Agent.ToolExecutionContext]::FromSession($session, $skill)
if (-not $readOnlyContext.IsToolAllowed("ReadRange") -or
    $readOnlyContext.IsToolAllowed("WriteData") -or
    $readOnlyContext.IsToolAllowed("CreateSheet")) {
    throw "Read-only mutation policy does not remain a hard tool boundary"
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

# A generated plan already contains executable command JSON. The loop must execute it
# directly instead of paying for a second model call per step.
$testTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$testTool.Id = "TestAction"
$testTool.Name = "Test action"
$testTool.AppType = "excel"
$testTool.RiskLevel = "safe"
$testTool.AccessMode = "write"
$registry.RegisterTool($testTool)
$registry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"changed","changed":true,"satisfied":true,"targetRefs":["Excel:test"]}')
    return [ShareRibbon.Agent.ToolResult]::Succeed("TestAction", "ok", $null, $observation)
}

$promptManager = [ShareRibbon.Agent.PromptManager]::new((Join-Path $repoRoot "ShareRibbon\Prompts"))

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

# A mandatory task contract is self-contained. The planner should not receive an entire
# broad Skill handbook when the semantic runtime has already selected the exact tools.
$contractSession = [ShareRibbon.Agent.AgentSession]::new("create a worksheet", "Excel", "")
$contractSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$contractSession.Spec.Goal = "create a worksheet"
$contractSession.Spec.RequiredTools.Add("CreateSheet")
$contractSession.Spec.MandatoryTools.Add("CreateSheet")
$contractSkill = [ShareRibbon.Agent.AgentSkill]::new()
$contractSkill.Name = "Broad Excel skill"
$contractSkill.Description = "general spreadsheet operations"
$contractSkill.RequiredTools.Add("CreateSheet")
$contractSkill.PromptTemplate = "SKILL-HANDBOOK-MARKER-" + ("x" * 7000)
$contractPrompt = $promptManager.BuildPlanningPrompt($contractSession, "", $contractSkill)
if ($contractPrompt.Contains("SKILL-HANDBOOK-MARKER") -or
    -not $contractPrompt.Contains("CreateSheet")) {
    throw "Planner prompt does not use the compact mandatory-tool contract"
}

$memory = [ShareRibbon.Agent.AgentMemory]::new()
$loop = [ShareRibbon.Agent.LoopEngine]::new($registry, $memory, $promptManager)
$script:modelCalls = 0
$loop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:modelCalls += 1
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"understanding":"test","steps":[{"step":1,"description":"run","code":"{\"command\":\"TestAction\",\"params\":{}}","language":"json"}],"summary":"done","capabilityGap":""}')
}
$planSession = [ShareRibbon.Agent.AgentSession]::new("run", "Excel", "")
$planSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$planSession.Spec.Goal = "run"
$planSession.Spec.RequiredTools.Add("TestAction")
$planSkill = [ShareRibbon.Agent.AgentSkill]::new()
$planSkill.RequiredTools.Add("TestAction")
$runTask = [System.Threading.Tasks.Task[ShareRibbon.Agent.AgentResult]]$loop.RunAsync(
    $planSession,
    "system",
    $planSkill)
$runResult = $runTask.GetAwaiter().GetResult()
if (-not $runResult.Success -or $script:modelCalls -ne 1) {
    throw "Plan execution made $script:modelCalls model calls; expected exactly one planning call"
}

# Incomplete plans get one bounded correction response and must never execute. This is
# a set-level task contract: no wording- or golden-case-specific branch is involved.
$coverageRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($id in @("TestAction", "RequiredAction")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $id
    $descriptor.Name = $id
    $descriptor.AppType = "excel"
    $descriptor.RiskLevel = "safe"
    $descriptor.AccessMode = "write"
    $coverageRegistry.RegisterTool($descriptor)
}
$script:coverageExecutions = 0
$coverageRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:coverageExecutions += 1
    return [ShareRibbon.Agent.ToolResult]::Succeed("TestAction", "unexpected")
}
$coverageLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $coverageRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:coverageModelCalls = 0
$coverageLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:coverageModelCalls += 1
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"understanding":"incomplete","steps":[{"step":1,"description":"wrong","code":"{\"command\":\"TestAction\",\"params\":{}}","language":"json"}],"summary":"wrong","capabilityGap":""}')
}
$coverageSession = [ShareRibbon.Agent.AgentSession]::new("contract", "Excel", "")
$coverageSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$coverageSession.Spec.Goal = "contract"
$coverageSession.Spec.RequiredTools.Add("TestAction")
$coverageSession.Spec.RequiredTools.Add("RequiredAction")
$coverageSession.Spec.MandatoryTools.Add("RequiredAction")
$coverageSkill = [ShareRibbon.Agent.AgentSkill]::new()
$coverageSkill.RequiredTools.Add("TestAction")
$coverageSkill.RequiredTools.Add("RequiredAction")
$coverageResult = $coverageLoop.RunAsync($coverageSession, "system", $coverageSkill).GetAwaiter().GetResult()
if ($coverageResult.Success -or $script:coverageExecutions -ne 0 -or $script:coverageModelCalls -ne 2) {
    throw "Incomplete mandatory-tool plan was executed or not bounded: calls=$script:coverageModelCalls executions=$script:coverageExecutions"
}

# Repair may change parameters for a mandatory engine, but may not silently replace it
# with a different tool and then claim the task succeeded.
$repairRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
foreach ($id in @("MandatoryAction", "FallbackAction")) {
    $descriptor = [ShareRibbon.Agent.ToolDescriptor]::new()
    $descriptor.Id = $id
    $descriptor.Name = $id
    $descriptor.AppType = "excel"
    $descriptor.RiskLevel = "safe"
    $descriptor.AccessMode = "write"
    $repairRegistry.RegisterTool($descriptor)
}
$script:fallbackExecutions = 0
$repairRegistry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $commandObject = [Newtonsoft.Json.Linq.JObject]::Parse($code)
    $command = $commandObject["command"].ToString()
    if ($command -eq "FallbackAction") {
        $script:fallbackExecutions += 1
        return [ShareRibbon.Agent.ToolResult]::Succeed("FallbackAction", "unexpected")
    }
    $failure = [ShareRibbon.Agent.ToolResult]::new()
    $failure.Success = $false
    $failure.ToolId = "MandatoryAction"
    $failure.Message = "repair parameters"
    $failure.UserMessage = "repair parameters"
    $failure.DebugDetail = "repair parameters"
    $failure.ErrorCode = "TEST_RECOVERABLE"
    $failure.Recoverable = $true
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
    if ($script:repairModelCalls -eq 1) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"understanding":"repair","steps":[{"step":1,"description":"run","code":"{\"command\":\"MandatoryAction\",\"params\":{}}","language":"json"}],"summary":"done","capabilityGap":""}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"toolId":"FallbackAction","parameters":{}}')
}
$repairSession = [ShareRibbon.Agent.AgentSession]::new("repair", "Excel", "")
$repairSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$repairSession.Spec.Goal = "repair"
$repairSession.Spec.RequiredTools.Add("MandatoryAction")
$repairSession.Spec.RequiredTools.Add("FallbackAction")
$repairSession.Spec.MandatoryTools.Add("MandatoryAction")
$repairSkill = [ShareRibbon.Agent.AgentSkill]::new()
$repairSkill.RequiredTools.Add("MandatoryAction")
$repairSkill.RequiredTools.Add("FallbackAction")
$repairResult = $repairLoop.RunAsync($repairSession, "system", $repairSkill).GetAwaiter().GetResult()
if ($repairResult.Success -or $script:fallbackExecutions -ne 0 -or $script:repairModelCalls -ne 2) {
    throw "Mandatory tool was substituted during repair: calls=$script:repairModelCalls fallbackExecutions=$script:fallbackExecutions"
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

# If generated source is syntactically invalid, Python can exit before consuming stdin.
# The broken input pipe is secondary; callers must receive the Python SyntaxError so the
# repair loop can correct the code instead of diagnosing a fake filesystem problem.
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
$expectedInput = [Newtonsoft.Json.Linq.JObject]::Parse(
    '{"workbook":"book.xlsx","sheet":"SalesData","address":"A1:B5","headers":["\u533a\u57df","\u9500\u552e\u989d"],"rows":[["\u534e\u4e1c",1000],["\u534e\u4e1c",2000],["\u534e\u5317",1500],["\u534e\u5317",2500]]}').ToString([Newtonsoft.Json.Formatting]::None)
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

Write-Host "PASS: agent runtime safety, code execution, and latency policies"
