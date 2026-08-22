param(
    [string]$Configuration = "Debug",
    [switch]$LiveExcel
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$outputDir = Join-Path $repoRoot "ExcelAi\bin\$Configuration"
$excelAssemblyPath = Join-Path $outputDir "ExcelAi.dll"

$requiredFiles = @(
    "ExcelAi\Runtime\ExcelApiCatalogProvider.vb",
    "ExcelAi\Runtime\ExcelObjectResolver.vb",
    "ExcelAi\Runtime\ExcelOperationExecutor.vb",
    "ExcelAi\Runtime\ExcelOperationObserver.vb",
    "ExcelAi\Runtime\ExcelFormatRangeAdapter.vb",
    "ShareRibbon\Tools\excel\DiscoverOfficeCapability.json",
    "ShareRibbon\Tools\excel\OfficeObjectOperation.json"
)
foreach ($relativePath in $requiredFiles) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Excel office-operation runtime is missing: $relativePath"
    }
}

foreach ($toolName in @("DiscoverOfficeCapability", "OfficeObjectOperation")) {
    $excelTool = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Tools\excel\$toolName.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $pptTool = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Tools\ppt\$toolName.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($excelTool.id -ne $pptTool.id -or
        $excelTool.name -ne $pptTool.name -or
        $excelTool.description -ne $pptTool.description -or
        $excelTool.riskLevel -ne $pptTool.riskLevel -or
        (($excelTool.parameters | ConvertTo-Json -Depth 10 -Compress) -ne ($pptTool.parameters | ConvertTo-Json -Depth 10 -Compress))) {
        throw "Merged Office tool has host-dependent interface: $toolName"
    }
}

$projectText = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelAi.vbproj") -Raw -Encoding UTF8
foreach ($fileName in @(
    "ExcelApiCatalogProvider.vb",
    "ExcelObjectResolver.vb",
    "ExcelOperationExecutor.vb",
    "ExcelOperationObserver.vb",
    "ExcelFormatRangeAdapter.vb"
)) {
    if ($projectText -notmatch [regex]::Escape($fileName)) {
        throw "ExcelAi project does not compile runtime module: $fileName"
    }
}

$directOperationSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelDirectOperationService.vb") -Raw -Encoding UTF8
if ($directOperationSource -notmatch 'Case\s+"formatrange"[\s\S]{0,240}ExcelFormatRangeAdapter\.Execute' -or
    $directOperationSource -match 'Return\s+ExecuteFormatRange\(' -or
    $directOperationSource -notmatch 'Obsolete\("Use OfficeRuntime\.ExcelFormatRangeAdapter;[^\r\n]*True\)') {
    throw "Legacy FormatRange execution can bypass the declarative runtime"
}

if (-not (Test-Path -LiteralPath $excelAssemblyPath)) {
    throw "ExcelAi assembly not found: $excelAssemblyPath. Build first."
}

Push-Location $outputDir
try {
    foreach ($dependency in @("Newtonsoft.Json.dll", "ShareRibbon.dll")) {
        [void][Reflection.Assembly]::LoadFrom((Join-Path $outputDir $dependency))
    }
    $assembly = [Reflection.Assembly]::LoadFrom($excelAssemblyPath)
    $catalogType = $assembly.GetType("ExcelAi.OfficeRuntime.ExcelApiCatalogProvider", $true)
    $executorType = $assembly.GetType("ExcelAi.OfficeRuntime.ExcelOperationExecutor", $true)
    $adapterType = $assembly.GetType("ExcelAi.OfficeRuntime.ExcelFormatRangeAdapter", $true)
    $staticFlags = [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::NonPublic

    $searchMethod = $catalogType.GetMethod("SearchAsToolResult", $staticFlags)
    $searchParams = [Newtonsoft.Json.Linq.JObject]::Parse('{"query":"number format thousands separator","targetType":"Range","includeReadOnly":true,"maxResults":12}')
    $searchArguments = New-Object object[] 1
    $searchArguments[0] = $searchParams
    $searchResult = $searchMethod.Invoke($null, $searchArguments)
    if (-not $searchResult.Success) {
        throw "Excel capability discovery failed: $($searchResult.Message)"
    }
    $searchJson = [Newtonsoft.Json.JsonConvert]::SerializeObject($searchResult.Data)
    if ($searchJson -notmatch 'NumberFormat') {
        throw "Excel capability discovery did not expose Range.NumberFormat"
    }

    $blockedParams = [Newtonsoft.Json.Linq.JObject]::Parse('{"query":"export chart","targetType":"Chart","includeReadOnly":true,"maxResults":20}')
    $searchArguments[0] = $blockedParams
    $blockedResult = $searchMethod.Invoke($null, $searchArguments)
    $blockedJson = [Newtonsoft.Json.JsonConvert]::SerializeObject($blockedResult.Data)
    $blockedObject = $blockedJson | ConvertFrom-Json
    $exportMember = @($blockedObject.Members | Where-Object { $_.MemberName -eq "Export" }) | Select-Object -First 1
    if ($null -ne $exportMember -and [bool]$exportMember.Executable) {
        throw "External-output Chart.Export was exposed as executable"
    }

    Write-Host "PASS: Excel capability catalog, runtime types, and tool schemas"

    if ($LiveExcel) {
        $excel = $null
        $workbook = $null
        $sheet = $null
        $range = $null
        try {
            $excelType = [Type]::GetTypeFromProgID("Excel.Application", $true)
            $excel = [Activator]::CreateInstance($excelType)
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Add()
            $sheet = $workbook.Worksheets.Item(1)
            $sheet.Name = "Data"
            $sheet.Range("D1").Value2 = "Amount"
            $sheet.Range("D2").Value2 = 1000
            $sheet.Range("D3").Value2 = 2500
            $sheet.Range("D4").Value2 = 2501
            $sheet.Range("D5").Value2 = 3200

            $adapterMethod = $adapterType.GetMethod("Execute", $staticFlags)
            $cases = @(
                '{"range":"Data!D2:D{lastRow}","numberFormat":"#,##0","bold":false}',
                '{"range":"Data!D2:D5","italic":true,"fontColor":"#FF0000","backgroundColor":"#E2F0D9"}',
                '{"range":"Data!D2:D5","horizontalAlignment":"center","verticalAlignment":"center","wrapText":true}'
            )
            foreach ($caseJson in $cases) {
                $caseParams = [Newtonsoft.Json.Linq.JObject]::Parse($caseJson)
                $adapterArguments = New-Object object[] 2
                $adapterArguments[0] = $excel
                $adapterArguments[1] = $caseParams
                $toolResult = $adapterMethod.Invoke($null, $adapterArguments)
                if (-not $toolResult.Success) {
                    throw "Property-driven FormatRange adapter failed: $($toolResult.ErrorCode) $($toolResult.Message)"
                }
            }

            $range = $sheet.Range("D2:D5")
            if ([string]$range.NumberFormat -ne "#,##0") { throw "NumberFormat was not applied" }
            if ([bool]$range.Font.Bold) { throw "Bold=false was not applied" }
            if (-not [bool]$range.Font.Italic) { throw "Italic=true was not applied" }
            if ([int]$range.Font.Color -ne 255) { throw "Font color was not applied" }
            if ([int]$range.Interior.Color -ne 14282978) { throw "Background color was not applied" }
            if (-not [bool]$range.WrapText) { throw "WrapText=true was not applied" }
            if ([string]$sheet.Range("D4").Text -ne "2,501") { throw "Displayed number format did not change" }

            $findMember = $catalogType.GetMethod("FindMemberId", $staticFlags)
            $findArguments = New-Object object[] 3
            $findArguments[0] = "Range"
            $findArguments[1] = "NumberFormat"
            $findArguments[2] = "property"
            $numberFormatMemberId = [string]$findMember.Invoke($null, $findArguments)
            if ([string]::IsNullOrWhiteSpace($numberFormatMemberId)) { throw "Range.NumberFormat member id was not found" }

            $batchJson = @"
{
  "batch": {
    "schemaVersion": "1.0",
    "appType": "Excel",
    "atomic": true,
    "operations": [{
      "id": "wrong-expectation",
      "targetRef": "Excel:workbooks/active/worksheets/Data/ranges/D2%3AD5",
      "action": "set",
      "memberId": "$numberFormatMemberId",
      "arguments": { "value": "#,##0.00" },
      "expectedEffects": { "NumberFormat": "#,##0" }
    }]
  }
}
"@
            $executorMethod = $executorType.GetMethod("Execute", $staticFlags)
            $executorArguments = New-Object object[] 2
            $executorArguments[0] = $excel
            $executorArguments[1] = [Newtonsoft.Json.Linq.JObject]::Parse($batchJson)
            $verifyFailure = $executorMethod.Invoke($null, $executorArguments)
            if ($verifyFailure.Success -or $verifyFailure.ErrorCode -ne "VERIFY_FAILED") {
                throw "Executor accepted a deliberately false expected effect"
            }

            $unverifiedBatch = [Newtonsoft.Json.Linq.JObject]::Parse($batchJson)
            $unverifiedBatch["batch"]["operations"][0]["expectedEffects"] = [Newtonsoft.Json.Linq.JObject]::new()
            $executorArguments[1] = $unverifiedBatch
            $unverifiedResult = $executorMethod.Invoke($null, $executorArguments)
            if ($unverifiedResult.Success -or $unverifiedResult.ErrorCode -ne "OPERATION_SCHEMA_INVALID") {
                throw "Executor accepted an unverified mutating operation"
            }

            $worksheetNameMemberId = [string]$findMember.Invoke($null, @("Worksheet", "Name", "property"))
            if ([string]::IsNullOrWhiteSpace($worksheetNameMemberId)) { throw "Worksheet.Name member id was not found" }
            $renameJson = @"
{
  "batch": {
    "schemaVersion": "1.0",
    "appType": "Excel",
    "atomic": true,
    "operations": [{
      "id": "rename-sheet",
      "targetRef": "Excel:workbooks/active/worksheets/Data",
      "action": "set",
      "memberId": "$worksheetNameMemberId",
      "arguments": { "value": "Renamed" },
      "expectedEffects": { "Name": "Renamed" }
    }]
  }
}
"@
            $executorArguments[1] = [Newtonsoft.Json.Linq.JObject]::Parse($renameJson)
            $renameResult = $executorMethod.Invoke($null, $executorArguments)
            if (-not $renameResult.Success -or [string]$sheet.Name -ne "Renamed") {
                throw "Identity-changing worksheet operation was not verified: $($renameResult.ErrorCode) $($renameResult.Message)"
            }

            $chartObjects = $null
            $chartObject = $null
            try {
                $chartObjects = $sheet.ChartObjects()
                $chartObject = $chartObjects.Add(400, 20, 300, 180)
                $chartObject.Name = "DeleteMe"
            }
            finally {
                if ($null -ne $chartObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($chartObject) }
                if ($null -ne $chartObjects) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($chartObjects) }
            }

            $chartDeleteMemberId = [string]$findMember.Invoke($null, @("ChartObject", "Delete", "method"))
            if ([string]::IsNullOrWhiteSpace($chartDeleteMemberId)) { throw "ChartObject.Delete member id was not found" }
            $deleteChartJson = @"
{
  "batch": {
    "schemaVersion": "1.0",
    "appType": "Excel",
    "atomic": true,
    "operations": [{
      "id": "delete-chart",
      "targetRef": "Excel:workbooks/active/worksheets/Renamed/chartobjects/DeleteMe",
      "action": "delete",
      "memberId": "$chartDeleteMemberId",
      "arguments": {},
      "expectedEffects": { "exists": false }
    }]
  }
}
"@
            $executorArguments[1] = [Newtonsoft.Json.Linq.JObject]::Parse($deleteChartJson)
            $deleteChartResult = $executorMethod.Invoke($null, $executorArguments)
            if (-not $deleteChartResult.Success -or [int]$sheet.ChartObjects().Count -ne 0) {
                throw "Generic ChartObject.Delete was not verified: $($deleteChartResult.ErrorCode) $($deleteChartResult.Message)"
            }

            $findArguments[0] = "Worksheets"
            $findArguments[1] = "Add"
            $findArguments[2] = "method"
            $worksheetAddMemberId = [string]$findMember.Invoke($null, $findArguments)
            if ([string]::IsNullOrWhiteSpace($worksheetAddMemberId)) { throw "Worksheets.Add member id was not found" }
            $createSheetJson = @"
{
  "batch": {
    "schemaVersion": "1.0",
    "appType": "Excel",
    "atomic": true,
    "operations": [{
      "id": "create-sheet",
      "targetRef": "Excel:workbooks/active/worksheets",
      "action": "create",
      "memberId": "$worksheetAddMemberId",
      "arguments": {},
      "expectedEffects": { "exists": true }
    }]
  }
}
"@
            $executorArguments[1] = [Newtonsoft.Json.Linq.JObject]::Parse($createSheetJson)
            $createSheetResult = $executorMethod.Invoke($null, $executorArguments)
            $createdSheetRef = [string]$createSheetResult.Data["operations"][0]["resultRef"]
            if (-not $createSheetResult.Success -or [string]::IsNullOrWhiteSpace($createdSheetRef) -or [int]$workbook.Worksheets.Count -ne 2) {
                throw "Generic Worksheets.Add did not return a verifiable object ref: $($createSheetResult.ErrorCode) $($createSheetResult.Message)"
            }

            $findArguments[0] = "Worksheet"
            $findArguments[1] = "Delete"
            $findArguments[2] = "method"
            $worksheetDeleteMemberId = [string]$findMember.Invoke($null, $findArguments)
            if ([string]::IsNullOrWhiteSpace($worksheetDeleteMemberId)) { throw "Worksheet.Delete member id was not found" }
            $deleteSheetJson = @"
{
  "batch": {
    "schemaVersion": "1.0",
    "appType": "Excel",
    "atomic": true,
    "operations": [{
      "id": "delete-created-sheet",
      "targetRef": "$createdSheetRef",
      "action": "delete",
      "memberId": "$worksheetDeleteMemberId",
      "arguments": {},
      "expectedEffects": { "exists": false }
    }]
  }
}
"@
            $executorArguments[1] = [Newtonsoft.Json.Linq.JObject]::Parse($deleteSheetJson)
            $deleteSheetResult = $executorMethod.Invoke($null, $executorArguments)
            if (-not $deleteSheetResult.Success -or [int]$workbook.Worksheets.Count -ne 1) {
                throw "Generic Worksheet.Delete was not verified: $($deleteSheetResult.ErrorCode) $($deleteSheetResult.Message)"
            }

            Write-Host "PASS: Live Excel generic execution, observation, semantic verification, and FormatRange adapter"
        }
        finally {
            if ($null -ne $range) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
            if ($null -ne $workbook) { $workbook.Close($false) }
            if ($null -ne $sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
            if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
            if ($null -ne $excel) {
                $excel.Quit()
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
finally {
    Pop-Location
}
