param(
    [string]$Configuration = "Debug",
    [switch]$LiveExcel
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$outputDir = Join-Path $repoRoot "ExcelAi\bin\$Configuration"
$excelAssemblyPath = Join-Path $outputDir "ExcelAi.dll"

if (-not (Test-Path -LiteralPath $excelAssemblyPath)) {
    throw "ExcelAi assembly not found: $excelAssemblyPath. Build first."
}

Push-Location $outputDir
try {
    foreach ($dependency in @("Newtonsoft.Json.dll", "ShareRibbon.dll")) {
        $dependencyPath = Join-Path $outputDir $dependency
        if (-not (Test-Path -LiteralPath $dependencyPath)) {
            throw "Conditional-format smoke dependency not found: $dependencyPath"
        }
        [void][Reflection.Assembly]::LoadFrom($dependencyPath)
    }

    $assembly = [Reflection.Assembly]::LoadFrom($excelAssemblyPath)
    $contractType = $assembly.GetType("ExcelAi.ExcelConditionalFormatContract", $true)
    $staticFlags = [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
    $instanceFlags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic

    $parseCondition = $contractType.GetMethod("ParseHighlightCondition", $staticFlags)
    $parseColor = $contractType.GetMethod("ParseColor", $staticFlags)
    $evaluate = $contractType.GetMethod("EvaluatePostState", $staticFlags)
    if ($null -eq $parseCondition -or $null -eq $parseColor -or $null -eq $evaluate) {
        throw "Conditional-format contract methods are missing"
    }

    $condition = $parseCondition.Invoke($null, [object[]]@(">2500"))
    $operatorValue = [int]$condition.GetType().GetProperty("ExcelOperator", $instanceFlags).GetValue($condition)
    $operand = [string]$condition.GetType().GetProperty("FormulaOperand", $instanceFlags).GetValue($condition)
    if ($operatorValue -ne 5 -or $operand -ne "2500") {
        throw "Condition normalization failed: operator=$operatorValue operand=$operand"
    }

    $lightGreen = [int]$parseColor.Invoke($null, [object[]]@("lightgreen"))
    $lightGreenHex = [int]$parseColor.Invoke($null, [object[]]@("#90EE90"))
    $lightGreenChineseName = [string]([char]0x6D45) + [char]0x7EFF + [char]0x8272
    $lightGreenChinese = [int]$parseColor.Invoke($null, [object[]]@($lightGreenChineseName))
    if ($lightGreen -ne 9498256 -or $lightGreenHex -ne $lightGreen -or $lightGreenChinese -ne $lightGreen) {
        throw "Color normalization failed: named=$lightGreen hex=$lightGreenHex zh=$lightGreenChinese"
    }

    $invalidRules = [Newtonsoft.Json.Linq.JArray]::Parse('[{"type":1,"operator":5,"formula1":"=\">2500\"","interiorColor":16777215}]')
    $invalidVerification = [Newtonsoft.Json.Linq.JObject]$evaluate.Invoke(
        $null,
        [object[]]@("highlight", ">2500", "lightgreen", $invalidRules))
    if ($invalidVerification["satisfied"].ToObject([bool])) {
        throw "Observer accepted the historical invalid formula/white-fill rule"
    }

    $validRules = [Newtonsoft.Json.Linq.JArray]::Parse('[{"type":1,"operator":5,"formula1":"=2500","interiorColor":9498256}]')
    $validVerification = [Newtonsoft.Json.Linq.JObject]$evaluate.Invoke(
        $null,
        [object[]]@("highlight", ">2500", "lightgreen", $validRules))
    if (-not $validVerification["satisfied"].ToObject([bool])) {
        throw "Observer rejected a normalized greater-than/light-green rule: $validVerification"
    }

    $executorSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ExcelDirectOperationService.vb") -Raw -Encoding UTF8
    $observerSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw -Encoding UTF8
    $loopSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\LoopEngine.Planning.vb") -Raw -Encoding UTF8
    if ($executorSource -notmatch 'ParseHighlightCondition\(condition\)' -or
        $executorSource -notmatch 'ResolveExcelRange\(range') {
        throw "ConditionalFormat executor is not using the shared normalization/range contract"
    }
    if ($observerSource -notmatch 'CaptureExcelConditionalFormats' -or
        $observerSource -notmatch 'EvaluatePostState' -or
        $observerSource -notmatch 'conditionalFormatCountDelta') {
        throw "ConditionalFormat observer is not verifying the real rule state"
    }
    if ($loopSource -notmatch 'hasSemanticVerification' -or
        $loopSource -notmatch 'satisfiedToken\.Value\(Of Boolean\)') {
        throw "Agent loop does not enforce semantic post-state verification"
    }

    Write-Host "PASS: Excel conditional-format normalization and semantic verification contract"

    if ($LiveExcel) {
        $excel = $null
        $workbook = $null
        $sheet = $null
        $range = $null
        $formatCondition = $null
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

            $serviceType = $assembly.GetType("ExcelAi.ExcelDirectOperationService", $true)
            $service = [Activator]::CreateInstance($serviceType, [object[]]@($excel))
            $command = [Newtonsoft.Json.Linq.JObject]::Parse('{"command":"ConditionalFormat","params":{"range":"Data!D2:D5","rule":"highlight","condition":">2500","color":"lightgreen"}}')
            $executeArguments = New-Object object[] 1
            $executeArguments[0] = $command
            $success = [bool]$serviceType.GetMethod("ExecuteCommand").Invoke($service, $executeArguments)

            $range = $sheet.Range("D2:D5")
            $formatCondition = $range.FormatConditions.Item(1)
            $ruleCount = [int]$range.FormatConditions.Count
            $operator = [int]$formatCondition.Operator
            $formula1 = [string]$formatCondition.Formula1
            $fillColor = [int64]$formatCondition.Interior.Color
            $belowThresholdColor = [int64]$sheet.Range("D3").DisplayFormat.Interior.Color
            $aboveThresholdColor1 = [int64]$sheet.Range("D4").DisplayFormat.Interior.Color
            $aboveThresholdColor2 = [int64]$sheet.Range("D5").DisplayFormat.Interior.Color

            if (-not $success -or
                $ruleCount -ne 1 -or
                $operator -ne 5 -or
                $formula1 -ne "=2500" -or
                $fillColor -ne 9498256 -or
                $belowThresholdColor -eq 9498256 -or
                $aboveThresholdColor1 -ne 9498256 -or
                $aboveThresholdColor2 -ne 9498256) {
                throw "Live Excel assertion failed: success=$success count=$ruleCount operator=$operator formula=$formula1 fill=$fillColor below=$belowThresholdColor above1=$aboveThresholdColor1 above2=$aboveThresholdColor2"
            }
            Write-Host "PASS: Live Excel conditional-format execution and displayed effect"
        }
        finally {
            if ($null -ne $formatCondition) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($formatCondition) }
            if ($null -ne $range) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
            if ($null -ne $workbook) { $workbook.Close($false) }
            if ($null -ne $sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
            if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
            if ($null -ne $excel) {
                $excel.Quit()
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            }
            $service = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
finally {
    Pop-Location
}
