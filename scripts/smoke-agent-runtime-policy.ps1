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

# A generated plan is only a high-level skeleton. Every action must be selected after the
# latest observation, so planned command JSON must never be executed directly.
$testTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$testTool.Id = "TestAction"
$testTool.Name = "Test action"
$testTool.AppType = "excel"
$testTool.RiskLevel = "safe"
$testTool.AccessMode = "write"
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
    $observation = [Newtonsoft.Json.Linq.JObject]::Parse(
        '{"kind":"write","summary":"changed","changed":true,"satisfied":true,"targetRefs":["Excel:test"]}')
    $observation["summary"] = [Newtonsoft.Json.Linq.JValue]::CreateString($message)
    return [ShareRibbon.Agent.ToolResult]::Succeed("TestAction", $message, $null, $observation)
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
    public static readonly Func<string, string, List<HistoryMessage>, Task<string>> PythonModelDelegate = PythonModel;
    public static readonly Func<string, string, bool, ToolResult> PythonHostDelegate = PythonHost;

    public static void Reset()
    {
        Captures = 0;
        PythonModelCalls = 0;
        PythonHostExecutions = 0;
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
                return Task.FromResult("{\"understanding\":\"python workflow\",\"steps\":[{\"step\":1,\"description\":\"read full data\",\"toolHint\":\"ReadRange\"},{\"step\":2,\"description\":\"compute averages\",\"toolHint\":\"PythonCompute\"},{\"step\":3,\"description\":\"create destination\",\"toolHint\":\"CreateSheet\"},{\"step\":4,\"description\":\"write result\",\"toolHint\":\"WriteData\"}],\"summary\":\"written\",\"capabilityGap\":\"\"}");
            case 2:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"read first\",\"action\":{\"tool\":\"ReadRange\",\"params\":{\"range\":\"SalesData!A1:B5\"}}}");
            case 3:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"compute observed rows\",\"action\":{\"tool\":\"PythonCompute\",\"params\":{\"code\":\"groups = {}\\nfor row in input_data['rows']:\\n    key = row[0]\\n    groups.setdefault(key, []).append(float(row[1]))\\nresult = [['Region', 'Average']] + [[key, sum(values) / len(values)] for key, values in sorted(groups.items())]\",\"input\":{\"preview\":true}}}}");
            case 4:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"create destination\",\"action\":{\"tool\":\"CreateSheet\",\"params\":{\"name\":\"PythonAverage\"}}}");
            case 5:
                return Task.FromResult("{\"decision\":\"act\",\"thought\":\"write computed rows\",\"action\":{\"tool\":\"WriteData\",\"params\":{\"targetRange\":\"PythonAverage!A1\",\"data\":[[\"wrong\",-1]]}}}");
            default:
                return Task.FromResult("{\"decision\":\"complete\",\"thought\":\"all four verified observations exist\",\"message\":\"done\"}");
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
            if (!written.Contains("1500.0") || !written.Contains("2000.0"))
                throw new InvalidOperationException("WriteData did not receive PythonCompute output: " + written);
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
            '{"understanding":"test","steps":[{"step":1,"description":"first","toolHint":"TestAction"},{"step":2,"description":"second","toolHint":"TestAction"}],"summary":"done","capabilityGap":""}')
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
    if ($script:modelCalls -eq 4) {
        if (-not $prompt.Contains("second-observation")) {
            throw "Completion decision did not receive the final tool observation"
        }
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"decision":"complete","thought":"verified by observations","message":"done"}')
    }
    return [System.Threading.Tasks.Task[string]]::FromResult(
        '{"decision":"act","thought":"choose from current facts","action":{"tool":"TestAction","params":{}}}')
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
if (-not $runResult.Success -or $script:modelCalls -ne 4 -or
    $script:adaptiveExecutions -ne 2 -or [AgentRuntimeContextProbe]::Captures -ne 3) {
    throw "Adaptive plan execution did not close the observation loop: calls=$script:modelCalls executions=$script:adaptiveExecutions contexts=$([AgentRuntimeContextProbe]::Captures) result=$($runResult.Message)"
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
$acceptRegistry.RegisterTool($acceptTool)
$supportTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$supportTool.Id = "SupportAction"
$supportTool.Name = "Support action"
$supportTool.AppType = "excel"
$supportTool.RiskLevel = "safe"
$supportTool.AccessMode = "write"
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
                '{"understanding":"acceptance","steps":[{"step":1,"description":"perform required action","toolHint":"RequiredAction"}],"summary":"done","capabilityGap":""}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"too early","message":"done"}') }
        3 {
            if (-not $prompt.Contains("Completion was rejected by deterministic acceptance")) {
                throw "Premature completion rejection was not fed back into ReAct"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"supporting action first","action":{"tool":"SupportAction","params":{}}}')
        }
        4 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"support action is insufficient","message":"done"}') }
        5 {
            if (-not $prompt.Contains("Completion was rejected by deterministic acceptance")) {
                throw "An unrelated successful tool incorrectly completed the plan milestone"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"close contract gap","action":{"tool":"RequiredAction","params":{}}}')
        }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"accepted evidence exists","message":"done"}') }
    }
}
$acceptSession = [ShareRibbon.Agent.AgentSession]::new("accept", "Excel", "")
$acceptSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$acceptSession.Spec.Goal = "accept"
$acceptSession.Spec.MandatoryTools.Add("RequiredAction")
$acceptSkill = [ShareRibbon.Agent.AgentSkill]::new()
$acceptSkill.RequiredTools.Add("RequiredAction")
$acceptSkill.RequiredTools.Add("SupportAction")
$acceptResult = $acceptLoop.RunAsync($acceptSession, "system", $acceptSkill).GetAwaiter().GetResult()
if (-not $acceptResult.Success -or $script:acceptExecutions -ne 2 -or $script:acceptModelCalls -ne 6) {
    throw "Deterministic completion gate failed: calls=$script:acceptModelCalls executions=$script:acceptExecutions result=$($acceptResult.Message)"
}

# Replanning is a normal decision after a successful observation, not merely a last-ditch
# error handler. The replacement skeleton must be followed by a fresh current action.
$replanRegistry = [ShareRibbon.Agent.ToolRegistry]::new($null)
$replanTool = [ShareRibbon.Agent.ToolDescriptor]::new()
$replanTool.Id = "ReplanAction"
$replanTool.Name = "Replan action"
$replanTool.AppType = "excel"
$replanTool.RiskLevel = "safe"
$replanTool.AccessMode = "write"
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
$replanLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $replanRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$script:replanModelCalls = 0
$replanLoop.SendAIRequest = {
    param($prompt, $system, $history)
    $script:replanModelCalls += 1
    switch ($script:replanModelCalls) {
        1 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"initial","steps":[{"step":1,"description":"initial action","toolHint":"ReplanAction"}],"summary":"initial","capabilityGap":""}') }
        2 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"first action","action":{"tool":"ReplanAction","params":{}}}') }
        3 {
            if (-not $prompt.Contains("replan-observation-1")) {
                throw "Successful observation was not available to the replan decision"
            }
            return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"replan","thought":"new facts require a revised skeleton","message":"revise"}')
        }
        4 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"understanding":"replanned","steps":[{"step":1,"description":"replanned action","toolHint":"ReplanAction"}],"summary":"replanned","capabilityGap":""}') }
        5 { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"act","thought":"act from replacement skeleton","action":{"tool":"ReplanAction","params":{}}}') }
        default { return [System.Threading.Tasks.Task[string]]::FromResult(
                '{"decision":"complete","thought":"all observations accepted","message":"done"}') }
    }
}
$replanSession = [ShareRibbon.Agent.AgentSession]::new("replan", "Excel", "")
$replanSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$replanSession.Spec.Goal = "replan"
$replanSkill = [ShareRibbon.Agent.AgentSkill]::new()
$replanSkill.RequiredTools.Add("ReplanAction")
$replanResult = $replanLoop.RunAsync($replanSession, "system", $replanSkill).GetAwaiter().GetResult()
if (-not $replanResult.Success -or $script:replanExecutions -ne 2 -or $script:replanModelCalls -ne 6) {
    throw "Success-triggered replan did not execute adaptively: calls=$script:replanModelCalls executions=$script:replanExecutions result=$($replanResult.Message)"
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
$pythonLoop = [ShareRibbon.Agent.LoopEngine]::new(
    $pythonLoopRegistry,
    [ShareRibbon.Agent.AgentMemory]::new(),
    $promptManager)
$pythonLoop.SendAIRequest = [AgentRuntimeContextProbe]::PythonModelDelegate
$pythonLoopSession = [ShareRibbon.Agent.AgentSession]::new("python workflow", "Excel", "")
$pythonLoopSession.Spec = [ShareRibbon.Agent.AgentTaskSpec]::new()
$pythonLoopSession.Spec.Goal = "python workflow"
$pythonLoopSkill = [ShareRibbon.Agent.AgentSkill]::new()
foreach ($toolId in @("ReadRange", "PythonCompute", "CreateSheet", "WriteData")) {
    $pythonLoopSession.Spec.RequiredTools.Add($toolId)
    $pythonLoopSession.Spec.MandatoryTools.Add($toolId)
    $pythonLoopSession.Spec.MandatoryToolSequence.Add($toolId)
    $pythonLoopSkill.RequiredTools.Add($toolId)
}
$pythonLoopResult = $pythonLoop.RunAsync($pythonLoopSession, "system", $pythonLoopSkill).GetAwaiter().GetResult()
if (-not $pythonLoopResult.Success -or $pythonLoopSession.CurrentIteration -ne 4 -or
    [AgentRuntimeContextProbe]::PythonHostExecutions -ne 3 -or
    [AgentRuntimeContextProbe]::PythonModelCalls -ne 6) {
    throw "Adaptive Python workflow failed: calls=$([AgentRuntimeContextProbe]::PythonModelCalls) actions=$($pythonLoopSession.CurrentIteration) hostExecutions=$([AgentRuntimeContextProbe]::PythonHostExecutions) result=$($pythonLoopResult.Message)"
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
    if ($script:repairModelCalls -eq 2) {
        return [System.Threading.Tasks.Task[string]]::FromResult(
            '{"thought":"run mandatory action","action":{"tool":"MandatoryAction","params":{}}}')
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
if ($repairResult.Success -or $script:fallbackExecutions -ne 0 -or $script:repairModelCalls -ne 3) {
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
