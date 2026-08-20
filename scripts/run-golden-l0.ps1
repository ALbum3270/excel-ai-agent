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
$catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
$results = @()

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

            $providerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\Context\ExcelContextProvider.vb") -Raw
            $detectorSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\Context\ExcelTableRegion.vb") -Raw
            if ($providerSource -notmatch 'ExcelTableRegionDetector' -or
                $providerSource -notmatch 'HostData\("tables"\)' -or
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
            $toolSchema = Get-Content -LiteralPath $toolPath -Raw | ConvertFrom-Json
            if ($toolSchema.id -ne "PythonCompute" -or
                $toolSchema.appType -ne "excel" -or
                $toolSchema.riskLevel -ne "risky") {
                throw "Golden case $($case.id): PythonCompute must be an approval-gated Excel tool"
            }

            $serviceSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Services\Python\PythonComputeService.vb") -Raw
            $registrySource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\ToolRegistry.vb") -Raw
            $skillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw
            if ($serviceSource -notmatch 'MaxTimeoutSeconds\s+As\s+Integer\s*=\s*60' -or
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
            $toolSchema = Get-Content -LiteralPath $toolPath -Raw | ConvertFrom-Json
            if ($toolSchema.id -ne "ReadRange" -or
                $toolSchema.appType -ne "excel" -or
                $toolSchema.riskLevel -ne "safe") {
                throw "Golden case $($case.id): ReadRange must be a safe Excel read tool"
            }

            $excelChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw
            $skillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw
            if ($excelChatSource -notmatch 'ReadExcelRangeAsToolResult' -or
                $excelChatSource -notmatch 'Math\.Min\(20000,\s*maxCells\)' -or
                $excelChatSource -notmatch '\{"kind",\s*"read"\}' -or
                $excelChatSource -notmatch 'ExcelMatrixToJson' -or
                $skillSource -notmatch 'allowed-tools:[^\r\n]*ReadRange' -or
                $skillSource -notmatch 'First use `ReadRange`') {
                throw "Golden case $($case.id): ReadRange routing or skill contract is incomplete"
            }
        }
        elseif ($case.contractType -eq "excel_basic_required_tools") {
            $requiredTools = @("FormatRange", "CreateChart", "CreateSheet")
            $skillSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Skills\excel-table-agent\SKILL.md") -Raw
            $schemaSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelJsonCommandSchema.vb") -Raw
            foreach ($requiredTool in $requiredTools) {
                $toolPath = Join-Path $repoRoot "ShareRibbon\Tools\excel\$requiredTool.json"
                if (-not (Test-Path -LiteralPath $toolPath)) {
                    throw "Golden case $($case.id): required Excel tool is missing: $requiredTool"
                }
                $toolSchema = Get-Content -LiteralPath $toolPath -Raw | ConvertFrom-Json
                if ($toolSchema.id -ne $requiredTool -or $toolSchema.appType -ne "excel") {
                    throw "Golden case $($case.id): invalid required Excel tool: $requiredTool"
                }
                if ($skillSource -notmatch "allowed-tools:[^\r\n]*$requiredTool" -or
                    $schemaSource -notmatch "`"$requiredTool`"") {
                    throw "Golden case $($case.id): required Excel tool is not routed by Skill/schema: $requiredTool"
                }
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
                $source = Get-Content -LiteralPath (Join-Path $repoRoot $sourceFile) -Raw
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

            $routeSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb") -Raw
            if ($routeSource -notmatch 'taskSpecRequiresExecution' -or
                $routeSource -notmatch 'Not\s+taskSpecRequiresExecution') {
                throw "Golden case $($case.id): TaskSpec execution fallback is missing"
            }
        }
        elseif ($case.contractType -eq "powerpoint_com_compat") {
            $classified = [ShareRibbon.ExceptionClassifier]::Classify(
                [InvalidCastException]::new("QueryInterface E_NOINTERFACE 不支持此接口"))
            if ($classified.ErrorCode -ne "COM_ERROR" -or $classified.Recoverable) {
                throw "Golden case $($case.id): E_NOINTERFACE must be terminal"
            }

            $rendererSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\Design\PowerPointSceneRenderer.vb") -Raw
            if ($rendererSource -match 'As\s+(FillFormat|LineFormat|ShadowFormat|ColorFormat|Adjustments|PictureFormat|TextFrame2|TextRange2|Font2|ParagraphFormat2)') {
                throw "Golden case $($case.id): renderer reintroduced a strongly typed Office style interface"
            }

            $pptChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\ChatControl.vb") -Raw
            if ($pptChatSource -notmatch 'ProfessionalDeckExecutor\.ExecuteAsToolResult\(\s*GetPowerPointEnvelopeParams\(envelope\),\s*preview\)') {
                throw "Golden case $($case.id): CreateSlides preview flag is not propagated"
            }
        }
        elseif ($case.contractType -eq "excel_observer_compat") {
            $excelChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw
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
    if ($null -ne $case.approvedTools) {
        foreach ($tool in $case.approvedTools) {
            $context.ApproveTool([string]$tool)
        }
    }

    $params = New-Object Newtonsoft.Json.Linq.JObject
    if ($null -ne $case.params) {
        $params = [Newtonsoft.Json.Linq.JObject]::FromObject($case.params)
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
