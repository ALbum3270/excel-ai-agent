param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$NoBuild
)

$ErrorActionPreference = "Stop"
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

if (-not $NoBuild) {
    & (Join-Path $repoRoot "scripts\build.ps1") -Configuration $Configuration -Clean
}

$forbiddenPaths = @(
    "WordAi", "PowerPointAi", "OfficeAgent", "ShareRibbon", "ExcelAgent.Core\Mcp", "ExcelAgent.Core\Storage",
    "ExcelAgent.Core\Controls", "ExcelAgent.Core\Translate", "ExcelAi\ExcelDirectOperationService.vb",
    "ExcelAi\ExcelDnaFunctions.vb", "ExcelAi\ExcelAi_TemporaryKey.pfx",
    "ExcelAgent.Core\Agent\SkillRegistry.vb", "ExcelAgent.Core\Tools\excel\ExecuteVBA.json"
)
foreach ($relativePath in $forbiddenPaths) {
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $repoRoot $relativePath))) "Legacy path remains: $relativePath"
}

$toolDirectory = Join-Path $repoRoot "ExcelAgent.Core\Tools\excel"
$toolIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($manifestPath in Get-ChildItem -LiteralPath $toolDirectory -Filter *.json -File) {
    $manifest = Get-Content -LiteralPath $manifestPath.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-True (-not [string]::IsNullOrWhiteSpace($manifest.id)) "Tool manifest has no id: $($manifestPath.Name)"
    Assert-True ($toolIds.Add([string]$manifest.id)) "Duplicate tool id: $($manifest.id)"
    Assert-True ($manifest.id -notmatch 'VBA|MCP|memory') "Forbidden tool is registered: $($manifest.id)"
}

$skillText = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAgent.Core\Skills\excel-table-agent\SKILL.md") -Raw -Encoding UTF8
$allowedLine = [regex]::Match($skillText, '(?m)^allowed-tools:\s*(.+)$')
Assert-True $allowedLine.Success "Excel Skill has no allowed-tools declaration"
foreach ($toolId in $allowedLine.Groups[1].Value.Split(',')) {
    $normalized = $toolId.Trim()
    Assert-True ($toolIds.Contains($normalized)) "Excel Skill references an unregistered tool: $normalized"
}
Assert-True ($skillText -notmatch 'ExecuteVBA') "Excel Skill still contains a VBA fallback"

$loopPath = Join-Path $repoRoot "ExcelAgent.Core\Agent\LoopEngine.vb"
$loopSource = Get-Content -LiteralPath $loopPath -Raw -Encoding UTF8
$completionAssignments = [regex]::Matches($loopSource, 'modelDeclaredComplete\s*=\s*True').Count
Assert-True ($completionAssignments -eq 1) "Completion state can be assigned outside the single completion gate"
Assert-True ([regex]::IsMatch($loopSource, 'Case\s+"complete"[\s\S]{0,500}TryAcceptCompletionDecision[\s\S]{0,300}modelDeclaredComplete\s*=\s*True')) "Explicit complete decision is not guarded by completion verification"
Assert-True ([regex]::IsMatch($loopSource, 'Dim\s+failMsg[\s\S]{0,600}decision=complete[\s\S]{0,400}AgentResult\.Failed')) "Iteration exhaustion can still be reported as success"
Assert-True ($loopSource -match 'RequestedMutationPolicy') "Loop does not apply explicit UI mutation authority"

$runnerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAgent.Core\Agent\AgentRunner.vb") -Raw -Encoding UTF8
$chatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw -Encoding UTF8
$gatewaySource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAgent.Core\Services\Ai\AiGateway.vb") -Raw -Encoding UTF8
Assert-True ($chatSource -match 'Case\s+"cancel"[\s\S]{0,100}CancelAsync') "Task pane stop action is not connected"
Assert-True ($runnerSource -match '_harness\.CancelAsync' -and $runnerSource -match '_cancellation\?\.Cancel') "Runner does not cancel both Harness and linked token"
Assert-True ($gatewaySource -match 'CreateLinkedTokenSource[\s\S]{0,200}options\.CancellationToken') "AI HTTP request is not linked to cancellation"
$cancelFunction = [regex]::Match($runnerSource, 'Public Async Function CancelAsync\(\)[\s\S]*?^\s*End Function', [Text.RegularExpressions.RegexOptions]::Multiline)
Assert-True $cancelFunction.Success "Runner CancelAsync implementation was not found"
Assert-True ($cancelFunction.Value -notmatch 'FinishRun') "CancelAsync can emit a duplicate terminal completion event"

$outputDirectory = Join-Path $repoRoot "ExcelAi\bin\$Configuration"
$newtonsoftPath = Join-Path $outputDirectory "Newtonsoft.Json.dll"
$corePath = Join-Path $outputDirectory "ExcelAgent.Core.dll"
$excelPath = Join-Path $outputDirectory "ExcelAi.dll"
foreach ($requiredPath in @($newtonsoftPath, $corePath, $excelPath)) {
    Assert-True (Test-Path -LiteralPath $requiredPath) "Build output is missing: $requiredPath"
}

[void][Reflection.Assembly]::LoadFrom($newtonsoftPath)
[void][Reflection.Assembly]::LoadFrom($corePath)
$registry = New-Object ExcelAgent.Core.Agent.ToolRegistry
$registry.LoadFromDirectory($toolDirectory)
Assert-True ($registry.ToolCount -eq $toolIds.Count) "Runtime tool count does not match Excel manifests"
Assert-True ($null -eq $registry.GetTool("ExecuteVBA")) "ExecuteVBA is visible at runtime"

$skill = New-Object ExcelAgent.Core.Agent.AgentSkill
$skill.Name = "narrow-test-skill"
$skill.RequiredTools.Add("ReadRange")
$executeSpec = New-Object ExcelAgent.Core.Agent.AgentTaskSpec
$executeSpec.MutationPolicy = "allow_mutation"
$executeSession = [ExcelAgent.Core.Agent.AgentSession]::new("test", "Excel", "")
$executeSession.Spec = $executeSpec
$executeContext = [ExcelAgent.Core.Agent.ToolExecutionContext]::FromSession($executeSession, $skill)
Assert-True ($executeContext.IsToolAllowed($registry.GetTool("CreateChart"))) "Skill metadata incorrectly hides a registered Excel host tool"

$readOnlySpec = New-Object ExcelAgent.Core.Agent.AgentTaskSpec
$readOnlySpec.MutationPolicy = "read_only"
$readOnlySession = [ExcelAgent.Core.Agent.AgentSession]::new("test", "Excel", "")
$readOnlySession.Spec = $readOnlySpec
$readOnlyContext = [ExcelAgent.Core.Agent.ToolExecutionContext]::FromSession($readOnlySession, $skill)
Assert-True (-not $readOnlyContext.IsToolAllowed($registry.GetTool("WriteData"))) "Read-only mode exposes a write tool"
Assert-True ($readOnlyContext.IsToolAllowed($registry.GetTool("ReadRange"))) "Read-only mode hides a read tool"

[void][Reflection.Assembly]::LoadFrom($excelPath)
$excelAssembly = [Reflection.Assembly]::LoadFrom($excelPath)
$catalogType = $excelAssembly.GetType("ExcelAi.OfficeRuntime.ExcelApiCatalogProvider", $true)
$searchMethod = $catalogType.GetMethod("SearchAsToolResult", [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::NonPublic)
$searchParams = [Newtonsoft.Json.Linq.JObject]::Parse('{"query":"number format percentage","targetType":"Range","includeReadOnly":true,"maxResults":12}')
$searchArguments = New-Object object[] 1
$searchArguments[0] = $searchParams
$searchResult = $searchMethod.Invoke($null, $searchArguments)
Assert-True $searchResult.Success "Excel capability discovery failed: $($searchResult.Message)"
Assert-True (([Newtonsoft.Json.JsonConvert]::SerializeObject($searchResult.Data)) -match 'NumberFormat') "Excel catalog did not expose Range.NumberFormat"

$unexpectedOutput = Get-ChildItem -LiteralPath $outputDirectory -Recurse -File | Where-Object {
    $_.FullName -match '\\(WordAi|PowerPointAi|Mcp|Storage)\\' -or $_.Name -eq 'ExecuteVBA.json'
}
Assert-True (@($unexpectedOutput).Count -eq 0) "Build output still contains removed host/runtime assets"

Write-Host "PASS: standalone Excel Agent architecture, completion gate, execution mode, cancellation, tools, and catalog"
