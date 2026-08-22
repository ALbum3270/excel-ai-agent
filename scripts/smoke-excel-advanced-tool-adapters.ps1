param(
    [string]$Configuration = "Debug",
    [switch]$LiveExcel
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$adapterPath = Join-Path $repoRoot "ExcelAi\Runtime\ExcelStandardToolAdapter.vb"
$projectPath = Join-Path $repoRoot "ExcelAi\ExcelAi.vbproj"

$adapterSource = Get-Content -LiteralPath $adapterPath -Raw -Encoding UTF8
$projectSource = Get-Content -LiteralPath $projectPath -Raw -Encoding UTF8
$advancedTools = @(
    "CleanData",
    "ConditionalFormat",
    "CreatePivotTable",
    "DataAnalysis",
    "TransformData",
    "GenerateReport"
)

foreach ($toolId in $advancedTools) {
    if ($adapterSource -notmatch ('"' + [regex]::Escape($toolId) + '"')) {
        throw "Structured adapter manifest is missing $toolId"
    }
}
if ($projectSource -notmatch 'ExcelStandardToolAdapter\.Advanced\.vb') {
    throw "Advanced Excel adapter is not registered in ExcelAi.vbproj"
}

Write-Host "PASS: advanced Excel tools are registered at the structured adapter seam"

if (-not $LiveExcel) {
    return
}

$outputDir = Join-Path $repoRoot "ExcelAi\bin\$Configuration"
$excelAssemblyPath = Join-Path $outputDir "ExcelAi.dll"
if (-not (Test-Path -LiteralPath $excelAssemblyPath)) {
    throw "ExcelAi assembly not found: $excelAssemblyPath. Build first."
}

Push-Location $outputDir
try {
    foreach ($dependency in @("Newtonsoft.Json.dll", "ShareRibbon.dll")) {
        [void][Reflection.Assembly]::LoadFrom((Join-Path $outputDir $dependency))
    }
    $assembly = [Reflection.Assembly]::LoadFrom($excelAssemblyPath)
    $adapterType = $assembly.GetType("ExcelAi.OfficeRuntime.ExcelStandardToolAdapter", $true)
    $tryExecute = $adapterType.GetMethod(
        "TryExecute",
        [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::NonPublic)

    $excel = $null
    $workbook = $null
    $dataSheet = $null
    try {
        $excelType = [Type]::GetTypeFromProgID("Excel.Application", $true)
        $excel = [Activator]::CreateInstance($excelType)
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $dataSheet = $workbook.Worksheets.Item(1)
        $dataSheet.Name = "Data"

        function Invoke-Adapter {
            param([string]$ToolId, [string]$Json)
            $arguments = New-Object object[] 4
            $arguments[0] = $excel
            $arguments[1] = $ToolId
            $arguments[2] = [Newtonsoft.Json.Linq.JObject]::Parse($Json)
            $arguments[3] = $null
            $handled = [bool]$tryExecute.Invoke($null, $arguments)
            $toolResult = $arguments[3]
            if (-not $handled -or $null -eq $toolResult -or -not $toolResult.Success) {
                $observationJson = [Newtonsoft.Json.JsonConvert]::SerializeObject($toolResult.Observation)
                throw "$ToolId failed: handled=$handled code=$($toolResult.ErrorCode) message=$($toolResult.Message) observation=$observationJson"
            }
            if ([string]$toolResult.ToolId -ne $ToolId) {
                throw "$ToolId returned the wrong ToolId: $($toolResult.ToolId)"
            }
            if ($null -eq $toolResult.Observation -or -not [bool]$toolResult.Observation["changed"] -or -not [bool]$toolResult.Observation["satisfied"]) {
                $observationJson = [Newtonsoft.Json.JsonConvert]::SerializeObject($toolResult.Observation)
                throw "$ToolId returned success without a changed and satisfied observation: $observationJson"
            }
            return $toolResult
        }

        $dataSheet.Range("A1").Value2 = "Category"
        $dataSheet.Range("A2").Value2 = " East "
        $dataSheet.Range("A3").Value2 = "West"
        $dataSheet.Range("A4").Value2 = "East"
        $dataSheet.Range("A5").Value2 = "West"
        $dataSheet.Range("B1").Value2 = "Amount"
        $dataSheet.Range("B2").Value2 = 10
        $dataSheet.Range("B3").Value2 = 20
        $dataSheet.Range("B4").Value2 = 20
        $dataSheet.Range("B5").Value2 = 5
        $dataSheet.Range("C1").Value2 = "Person"
        $dataSheet.Range("C2").Value2 = "P1"
        $dataSheet.Range("C3").Value2 = "P1"
        $dataSheet.Range("C4").Value2 = "P2"
        $dataSheet.Range("C5").Value2 = "P2"

        Invoke-Adapter "CleanData" '{"range":"Data!A2:A5","operation":"trim"}' | Out-Null
        if ([string]$dataSheet.Range("A2").Value2 -ne "East") { throw "CleanData did not trim text" }

        $dataSheet.Range("D1").Value2 = "Dirty"
        $dataSheet.Range("D2").ClearContents()
        $dataSheet.Range("D3").Value2 = " alpha beta "
        Invoke-Adapter "CleanData" '{"range":"Data!D2:D2","operation":"fillempty","fillValue":"missing"}' | Out-Null
        Invoke-Adapter "CleanData" '{"range":"Data!D3:D3","operation":"trim"}' | Out-Null
        Invoke-Adapter "CleanData" '{"range":"Data!D3:D3","operation":"replace","findText":"beta","replaceText":"gamma"}' | Out-Null
        if ([string]$dataSheet.Range("D2").Value2 -ne "missing" -or [string]$dataSheet.Range("D3").Value2 -ne "alpha gamma") {
            throw "CleanData fill/trim/replace output is incorrect"
        }

        $dataSheet.Range("E1").Value2 = "Key"
        $dataSheet.Range("F1").Value2 = "Value"
        $dataSheet.Range("E2").Value2 = "x"
        $dataSheet.Range("F2").Value2 = 1
        $dataSheet.Range("E3").Value2 = "x"
        $dataSheet.Range("F3").Value2 = 1
        $dataSheet.Range("E4").Value2 = "y"
        $dataSheet.Range("F4").Value2 = 2
        Invoke-Adapter "CleanData" '{"range":"Data!E1:F4","operation":"removeduplicates","hasHeader":true}' | Out-Null
        if ([string]$dataSheet.Range("E3").Value2 -ne "y" -or $null -ne $dataSheet.Range("E4").Value2) {
            throw "CleanData removeduplicates output is incorrect"
        }

        $summarySheet = $workbook.Worksheets.Add()
        $summarySheet.Name = "Summary"
        Invoke-Adapter "DataAnalysis" '{"sourceRange":"Data!A1:B5","type":"groupby","groupBy":"Category","valueField":"Amount","aggregate":"sum","targetRange":"Summary!A1"}' | Out-Null
        $groupSummaryTitle = [regex]::Unescape('\u5206\u7ec4\u6c47\u603b')
        if ([string]$summarySheet.Range("A1").Value2 -ne $groupSummaryTitle -or [double]$summarySheet.Range("B3").Value2 -ne 30) {
            throw "DataAnalysis groupby output is incorrect"
        }
        Invoke-Adapter "DataAnalysis" '{"sourceRange":"Data!A1:B5","type":"summary","targetRange":"Summary!E1"}' | Out-Null
        Invoke-Adapter "DataAnalysis" '{"sourceRange":"Data!A1:B5","type":"ranking","labelField":"Category","rankBy":"Amount","topN":2,"targetRange":"Summary!H1"}' | Out-Null
        if ([double]$summarySheet.Range("J3").Value2 -ne 20 -or [double]$summarySheet.Range("J4").Value2 -ne 20) {
            throw "DataAnalysis summary/ranking output is incorrect"
        }

        $transformSheet = $workbook.Worksheets.Add()
        $transformSheet.Name = "Transformed"
        Invoke-Adapter "TransformData" '{"sourceRange":"Data!A1:B3","operation":"transpose","targetRange":"Transformed!A1"}' | Out-Null
        if ([string]$transformSheet.Range("A2").Value2 -ne "Amount" -or [double]$transformSheet.Range("C2").Value2 -ne 20) {
            throw "TransformData transpose output is incorrect"
        }
        $transformSheet.Range("F1").Value2 = "A|B"
        $transformSheet.Range("F2").Value2 = "C|D"
        Invoke-Adapter "TransformData" '{"sourceRange":"Transformed!F1:F2","operation":"split","delimiter":"|","targetRange":"Transformed!G1"}' | Out-Null
        Invoke-Adapter "TransformData" '{"sourceRange":"Transformed!G1:H2","operation":"merge","delimiter":"-","targetRange":"Transformed!J1"}' | Out-Null
        if ([string]$transformSheet.Range("H2").Value2 -ne "D" -or [string]$transformSheet.Range("J1").Value2 -ne "A-B") {
            throw "TransformData split/merge output is incorrect"
        }

        Invoke-Adapter "GenerateReport" '{"sourceRange":"Data!A1:B5","targetSheet":"Report","title":"Sales Report","includeChart":false}' | Out-Null
        $reportSheet = $workbook.Worksheets.Item("Report")
        if ([string]$reportSheet.Range("A1").Value2 -ne "Sales Report" -or [string]$reportSheet.Range("A3").Value2 -ne "Category") {
            throw "GenerateReport output is incorrect"
        }
        Invoke-Adapter "GenerateReport" '{"sourceRange":"Data!A1:B5","targetSheet":"ReportWithChart","title":"Sales Chart Report","includeChart":true}' | Out-Null
        if ([int]$workbook.Worksheets.Item("ReportWithChart").ChartObjects().Count -ne 1) {
            throw "GenerateReport includeChart output is incorrect"
        }

        $pivotSheet = $workbook.Worksheets.Add()
        $pivotSheet.Name = "Pivot"
        Invoke-Adapter "CreatePivotTable" '{"sourceRange":"Data!A1:C5","targetCell":"Pivot!A3","rowFields":["Category"],"valueFields":["Amount"],"columnFields":["Person"]}' | Out-Null
        if ([int]$pivotSheet.PivotTables().Count -ne 1) { throw "CreatePivotTable did not create a pivot table at the requested sheet" }

        $analysisPivotSheet = $workbook.Worksheets.Add()
        $analysisPivotSheet.Name = "PivotViaAnalysis"
        Invoke-Adapter "DataAnalysis" '{"sourceRange":"Data!A1:C5","type":"pivot","targetRange":"PivotViaAnalysis!A3","rowFields":["Category"],"valueFields":["Amount"]}' | Out-Null
        if ([int]$analysisPivotSheet.PivotTables().Count -ne 1) { throw "DataAnalysis pivot delegation is incorrect" }

        Invoke-Adapter "ConditionalFormat" '{"range":"Data!B2:B5","rule":"highlight","condition":">15","color":"lightgreen"}' | Out-Null
        Invoke-Adapter "ConditionalFormat" '{"range":"Data!B2:B5","rule":"databar"}' | Out-Null
        if ([int]$dataSheet.Range("B2:B5").FormatConditions.Count -ne 2) { throw "ConditionalFormat did not create both rule types" }

        Write-Host "PASS: live Excel advanced adapters execute and semantically verify all migrated tool families"
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $dataSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($dataSheet) }
        if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($null -ne $excel) {
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
finally {
    Pop-Location
}
