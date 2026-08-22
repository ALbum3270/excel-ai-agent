param(
    [string]$Configuration = "Debug",
    [switch]$LiveExcel
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$adapterPath = Join-Path $repoRoot "ExcelAi\Runtime\ExcelStandardToolAdapter.vb"
$readAdapterPath = Join-Path $repoRoot "ExcelAi\Runtime\ExcelReadRangeAdapter.vb"

if (-not (Test-Path -LiteralPath $adapterPath) -or -not (Test-Path -LiteralPath $readAdapterPath)) {
    throw "Excel standard/read adapter module is missing"
}

$adapterSource = Get-Content -LiteralPath $adapterPath -Raw -Encoding UTF8
$migratedTools = @(
    "WriteData",
    "ApplyFormula",
    "SortData",
    "FilterData",
    "RemoveDuplicates",
    "MergeCells",
    "AutoFit",
    "FindReplace",
    "CreateSheet",
    "DeleteSheet",
    "RenameSheet",
    "CopySheet",
    "InsertRowCol",
    "DeleteRowCol",
    "HideRowCol",
    "ProtectSheet",
    "CreateChart"
)

foreach ($toolId in $migratedTools) {
    if ($adapterSource -notmatch ('"' + [regex]::Escape($toolId) + '"')) {
        throw "Migration manifest is missing $toolId"
    }
}

$chatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw -Encoding UTF8
$directSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelDirectOperationService.vb") -Raw -Encoding UTF8
$projectSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelAi.vbproj") -Raw -Encoding UTF8
if ($chatSource -notmatch 'ExcelStandardToolAdapter\.TryExecute' -or
    $chatSource -notmatch 'ExcelReadRangeAdapter\.Execute' -or
    $directSource -notmatch 'ExcelStandardToolAdapter\.TryExecute' -or
    $projectSource -notmatch 'ExcelStandardToolAdapter\.vb' -or
    $projectSource -notmatch 'ExcelReadRangeAdapter\.vb') {
    throw "Standard Excel tools can still bypass the declarative runtime adapter seam"
}

Write-Host "PASS: standard Excel tools are registered at the declarative adapter seam"

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
    $readAdapterType = $assembly.GetType("ExcelAi.OfficeRuntime.ExcelReadRangeAdapter", $true)
    $readExecute = $readAdapterType.GetMethod(
        "Execute",
        [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::NonPublic)

    $excel = $null
    $workbook = $null
    $sheet = $null
    try {
        $excelType = [Type]::GetTypeFromProgID("Excel.Application", $true)
        $excel = [Activator]::CreateInstance($excelType)
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = "Data"

        function Invoke-Adapter {
            param([string]$ToolId, [string]$Json)
            Write-Host "LIVE: $ToolId"
            $resultBox = $null
            $arguments = New-Object object[] 4
            $arguments[0] = $excel
            $arguments[1] = $ToolId
            $arguments[2] = [Newtonsoft.Json.Linq.JObject]::Parse($Json)
            $arguments[3] = $resultBox
            $handled = [bool]$tryExecute.Invoke($null, $arguments)
            $toolResult = $arguments[3]
            if (-not $handled -or $null -eq $toolResult -or -not $toolResult.Success) {
                $observationJson = [Newtonsoft.Json.JsonConvert]::SerializeObject($toolResult.Observation)
                throw "$ToolId failed: handled=$handled code=$($toolResult.ErrorCode) message=$($toolResult.Message) observation=$observationJson"
            }
            return $toolResult
        }

        Invoke-Adapter "WriteData" '{"targetRange":"Data!A1","data":[["Name","Amount"],["B",20],["A",10],["A",10]]}' | Out-Null
        $readResult = $readExecute.Invoke($null, @($excel, [Newtonsoft.Json.Linq.JObject]::Parse('{"range":"Data!A1:B4"}')))
        if ($null -eq $readResult -or -not $readResult.Success -or
            [int]$readResult.Data["rowCount"] -ne 4 -or
            [int]$readResult.Data["columnCount"] -ne 2 -or
            $readResult.Data["values"].Count -ne 4 -or
            [bool]$readResult.Observation["changed"]) {
            throw "ReadRange did not return the complete structured range without mutation"
        }
        Invoke-Adapter "ApplyFormula" '{"targetRange":"Data!C2:C4","formula":"=B2*2","fillDown":true}' | Out-Null
        if ([double]$sheet.Range("C4").Value2 -ne 20) { throw "ApplyFormula did not fill relative formulas" }

        Invoke-Adapter "SortData" '{"range":"Data!A1:C4","sortColumn":2,"order":"asc","hasHeader":true}' | Out-Null
        if ([double]$sheet.Range("B2").Value2 -ne 10 -or [double]$sheet.Range("B4").Value2 -ne 20) { throw "SortData post-state is not sorted" }

        Invoke-Adapter "RemoveDuplicates" '{"range":"Data!A1:C4","columns":[1,2,3],"hasHeader":true}' | Out-Null
        if ([int]$sheet.UsedRange.Rows.Count -ne 3) { throw "RemoveDuplicates did not remove the duplicate row" }

        Invoke-Adapter "FindReplace" '{"range":"Data!A1:C3","find":"A","replace":"Alpha","matchCase":true,"matchEntireCell":true}' | Out-Null
        if ([string]$sheet.Range("A2").Value2 -ne "Alpha") { throw "FindReplace did not produce the expected value" }

        Invoke-Adapter "FilterData" '{"range":"Data!A1:C3","column":1,"criteria":"Alpha"}' | Out-Null
        if (-not [bool]$sheet.FilterMode) { throw "FilterData did not produce an active filtered view" }
        Invoke-Adapter "FilterData" '{"range":"Data!A1:C3","clearFilter":true}' | Out-Null
        if ([bool]$sheet.AutoFilterMode) { throw "FilterData did not clear the filter" }

        $sheet.Columns("A").ColumnWidth = 2
        Invoke-Adapter "AutoFit" '{"range":"Data!A1:A3","type":"columns"}' | Out-Null
        if ([double]$sheet.Columns("A").ColumnWidth -le 2) { throw "AutoFit did not change the constrained column width" }

        $beforeInsert = [string]$sheet.Range("A2").Value2
        Invoke-Adapter "InsertRowCol" '{"type":"row","position":"2","count":1}' | Out-Null
        if ($null -ne $sheet.Range("A2").Value2 -or [string]$sheet.Range("A3").Value2 -ne $beforeInsert) {
            throw "InsertRowCol did not shift worksheet values"
        }
        Invoke-Adapter "DeleteRowCol" '{"type":"row","position":"2","count":1}' | Out-Null
        if ([string]$sheet.Range("A2").Value2 -ne $beforeInsert) { throw "DeleteRowCol did not restore the worksheet layout" }

        Invoke-Adapter "MergeCells" '{"range":"Data!E1:F1","unmerge":false}' | Out-Null
        if (-not [bool]$sheet.Range("E1:F1").MergeCells) { throw "MergeCells was not observed" }
        Invoke-Adapter "MergeCells" '{"range":"Data!E1:F1","unmerge":true}' | Out-Null

        Invoke-Adapter "HideRowCol" '{"type":"column","position":"F","unhide":false}' | Out-Null
        if (-not [bool]$sheet.Columns("F").Hidden) { throw "HideRowCol was not observed" }
        Invoke-Adapter "HideRowCol" '{"type":"column","position":"F","unhide":true}' | Out-Null

        Invoke-Adapter "CreateSheet" '{"name":"Created"}' | Out-Null
        Invoke-Adapter "RenameSheet" '{"oldName":"Created","newName":"Renamed"}' | Out-Null
        Invoke-Adapter "CopySheet" '{"sourceName":"Data","newName":"DataCopy"}' | Out-Null
        if ([int]$workbook.Worksheets.Count -ne 3) { throw "Worksheet lifecycle count is wrong" }
        Invoke-Adapter "DeleteSheet" '{"name":"Renamed"}' | Out-Null

        Invoke-Adapter "ProtectSheet" '{"sheetName":"DataCopy","password":"smoke","unprotect":false}' | Out-Null
        if (-not [bool]$workbook.Worksheets.Item("DataCopy").ProtectContents) { throw "ProtectSheet was not observed" }
        Invoke-Adapter "ProtectSheet" '{"sheetName":"DataCopy","password":"smoke","unprotect":true}' | Out-Null

        Invoke-Adapter "CreateChart" '{"dataRange":"Data!A1:B3","type":"line","title":"Trend","position":"E3","legendPosition":"right"}' | Out-Null
        $chartObject = $sheet.ChartObjects().Item(1)
        try {
            if (-not [bool]$chartObject.Chart.HasTitle -or [string]$chartObject.Chart.ChartTitle.Text -ne "Trend") {
                throw "CreateChart title was not observed"
            }
            if ([int]$chartObject.Chart.SeriesCollection().Count -lt 1) { throw "CreateChart has no series" }
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($chartObject)
        }

        Write-Host "PASS: live Excel standard adapters execute and verify multiple capability families"
    }
    finally {
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
finally {
    Pop-Location
}
