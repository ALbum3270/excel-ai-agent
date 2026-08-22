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

$hostCalls = 0
$registry.ExecuteCodeWithToolResult = {
    param($code, $language, $preview)
    $script:hostCalls += 1
    $command = [Newtonsoft.Json.Linq.JObject]::Parse($code)
    if ($command["command"].ToString() -ne "CreateSheet" -or
        $command["params"]["name"].ToString() -ne "Python平均值0821") {
        throw "Unexpected host command: $code"
    }
    return [ShareRibbon.Agent.ToolResult]::Succeed(
        "CreateSheet",
        "created",
        $null,
        [Newtonsoft.Json.Linq.JObject]::Parse('{"kind":"worksheet","changed":true,"summary":"created"}'))
}

$broker = [ShareRibbon.OfficeToolBroker]::new($registry)
$application = [ShareRibbon.ApplicationInfo]::new(
    "Excel",
    [ShareRibbon.OfficeApplicationType]::Excel)
$context = [ShareRibbon.ChatRequestContext]::new()
$context.AppInfo = $application
$toolSchemas = $broker.GetTools($context)
$createSheetSchema = @($toolSchemas | Where-Object {
    $_["function"]["name"].ToString() -eq "CreateSheet"
})
if ($createSheetSchema.Count -ne 1 -or
    $createSheetSchema[0]["function"]["parameters"]["required"][0].ToString() -ne "name") {
    throw "Plain chat request does not expose the registered CreateSheet contract"
}

$nonPublicStatic = [System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::NonPublic
$snapshot = [Newtonsoft.Json.Linq.JObject]::new()
$attachFacts = [ShareRibbon.Agent.AiNativeRuntime].GetMethod("AttachRuntimeToolFacts", $nonPublicStatic)
$buildIntentContext = [ShareRibbon.IntentRecognitionService].GetMethod("BuildIntentContextInfo", $nonPublicStatic)
if ($null -eq $attachFacts -or $null -eq $buildIntentContext) {
    throw "Runtime capability facts are not wired into intent analysis"
}
$attachArgs = [object[]]::new(2)
$attachArgs[0] = $snapshot
$attachArgs[1] = $registry.GetAvailableTools("Excel")
$attachFacts.Invoke($null, $attachArgs) | Out-Null
$contextArgs = [object[]]::new(1)
$contextArgs[0] = $snapshot
$intentContext = $buildIntentContext.Invoke($null, $contextArgs).ToString()
if (-not $intentContext.Contains("CreateSheet") -or -not $intentContext.Contains("available")) {
    throw "Intent model cannot see the authoritative runtime tool registry"
}

$runtime = [ShareRibbon.ChatToolCallRuntime]::new($registry)
$toolCall = $null
$content = '{"tool":"CreateSheet","parameters":{"name":"Python平均值0821"}}'
if (-not [ShareRibbon.ChatToolCallParser]::TryParse($content, [ref]$toolCall)) {
    throw "Structured tool call in assistant content was not recognized"
}

$executionTask = [System.Threading.Tasks.Task[ShareRibbon.Agent.ToolResult]]$runtime.ExecuteAsync(
    "Excel",
    $toolCall.ToolId,
    $toolCall.Parameters)
$result = $executionTask.GetAwaiter().GetResult()
if ($null -eq $result -or -not $result.Success -or $script:hostCalls -ne 1) {
    throw "Structured chat tool call did not reach the shared host executor"
}

$deleteTask = [System.Threading.Tasks.Task[ShareRibbon.Agent.ToolResult]]$runtime.ExecuteAsync(
    "Excel",
    "DeleteSheet",
    [Newtonsoft.Json.Linq.JObject]::Parse('{"sheetName":"must-confirm"}'))
$deleteResult = $deleteTask.GetAwaiter().GetResult()
if ($null -eq $deleteResult -or $deleteResult.Success -or
    $deleteResult.ErrorCode -ne [ShareRibbon.ExceptionClassifier]::CodeSafetyNeedsApproval -or
    $script:hostCalls -ne 1) {
    throw "Unified chat tool execution bypassed SafetyGate for a destructive call"
}

$quoted = 'The following is only an example: {"tool":"CreateSheet","parameters":{"name":"do-not-run"}}'
$ignored = $null
if ([ShareRibbon.ChatToolCallParser]::TryParse($quoted, [ref]$ignored)) {
    throw "Tool-call parser accepted prose containing a quoted example"
}

Write-Output "PASS: chat and agent share registered Office tool schemas and execution"
