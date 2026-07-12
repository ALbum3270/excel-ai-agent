param(
    [int]$LargeFileThreshold = 800
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$sourceRoots = @("ShareRibbon", "WordAi", "ExcelAi", "PowerPointAi")

$allowedLargeFiles = @(
    "WordAi\ChatControl.vb",
    "ShareRibbon\Controls\BaseChatControl.vb",
    "PowerPointAi\ChatControl.vb",
    "ShareRibbon\Config\ConfigApiForm.vb",
    "ShareRibbon\Controls\BaseDataCapturePane.vb",
    "ExcelAi\ExcelDirectOperationService.vb",
    "ShareRibbon\Controls\Services\IntentRecognitionService.vb",
    "ShareRibbon\Mcp\MCPConfigForm.vb",
    "ShareRibbon\Services\Reformat\SmartFormattingOrchestrator.vb",
    "ShareRibbon\Controls\Services\HttpStreamService.vb",
    "WordAi\WordDocumentTranslateService.vb",
    "ExcelAi\ChatControl.vb",
    "ShareRibbon\Services\Reformat\DocumentAnalyzer.vb",
    "ShareRibbon\Controls\BaseDeepseekChat.vb",
    "ShareRibbon\Controls\ReformatTemplateEditorControl.vb",
    "ShareRibbon\Controls\Services\ReformatService.vb",
    "ShareRibbon\Storage\MemoryRepository.vb",
    "ExcelAi\ExcelJsonCommandSchema.vb",
    "ShareRibbon\Config\ConfigPromptForm.vb",
    "ShareRibbon\Agent\ToolRegistry.vb",
    "ShareRibbon\Controls\Services\ChatFormatterAgent.vb",
    "ShareRibbon\Config\PromptManager.vb",
    "ShareRibbon\Config\ReformatTemplateManager.vb",
    "ShareRibbon\Services\SkillsService.vb",
    "ShareRibbon\Config\ReformatTemplateEditorForm.vb",
    "ShareRibbon\Controls\BaseDoubaoChat.vb",
    "ShareRibbon\Storage\OfficeAiDatabase.vb",
    "ShareRibbon\Storage\AgentMemoryRepository.vb"
)

function Get-RelativePath {
    param([string]$Path)
    return $Path.Substring($repoRoot.Path.Length + 1)
}

function Get-SourceFiles {
    Get-ChildItem -Path ($sourceRoots | ForEach-Object { Join-Path $repoRoot $_ }) -Recurse -Filter *.vb -File |
        Where-Object {
            $_.FullName -notmatch "\\(bin|obj)\\" -and
            $_.Name -notlike "*.Designer.vb"
        }
}

function Test-CodeLineMatch {
    param(
        [string]$Path,
        [string]$Pattern
    )

    foreach ($line in Get-Content -LiteralPath $Path) {
        $trimmed = $line.TrimStart()
        if ($trimmed.StartsWith("'")) {
            continue
        }
        if ($line -match $Pattern) {
            return $true
        }
    }
    return $false
}

$failures = New-Object System.Collections.Generic.List[string]

# P0-1 guardrail: no new >800-line VB files outside the baseline list.
$largeFiles = Get-SourceFiles | ForEach-Object {
    $lineCount = (Get-Content -LiteralPath $_.FullName).Count
    if ($lineCount -gt $LargeFileThreshold) {
        [pscustomobject]@{
            Path = Get-RelativePath $_.FullName
            Lines = $lineCount
        }
    }
}

$unexpectedLarge = $largeFiles | Where-Object { $allowedLargeFiles -notcontains $_.Path }
foreach ($item in $unexpectedLarge) {
    $failures.Add("New oversized VB file detected: $($item.Path) ($($item.Lines) lines). Move responsibilities into a service instead of growing a god class.")
}

# P0-2 guardrail: Ralph UI must not be loaded as a product entrypoint.
$htmlFiles = Get-ChildItem -Path (Join-Path $repoRoot "ShareRibbon\Resources") -Recurse -Include *.html -File
foreach ($html in $htmlFiles) {
    $content = Get-Content -LiteralPath $html.FullName -Raw
    if ($content -match '<script[^>]+ralph-loop\.js') {
        $failures.Add("Ralph loop script is still loaded by HTML: $(Get-RelativePath $html.FullName)")
    }
}

# P0-3 guardrail: Task.Run in host projects that also import Office interop must be explicitly reviewed.
$hostProjectPaths = @(
    (Join-Path $repoRoot "WordAi"),
    (Join-Path $repoRoot "ExcelAi"),
    (Join-Path $repoRoot "PowerPointAi")
)

$taskRunInteropFiles = Get-ChildItem -Path $hostProjectPaths -Recurse -Filter *.vb -File |
    Where-Object { $_.FullName -notmatch "\\(bin|obj)\\" -and $_.Name -notlike "*.Designer.vb" } |
    Where-Object {
        (Test-CodeLineMatch -Path $_.FullName -Pattern "Task\.Run") -and
        (Test-CodeLineMatch -Path $_.FullName -Pattern "Microsoft\.Office\.Interop")
    } |
    ForEach-Object { Get-RelativePath $_.FullName }

$allowedTaskRunInteropFiles = @(
    "ExcelAi\ThisAddIn.vb",
    "ExcelAi\ChatControl.vb",
    "ExcelAi\EnhancedPreviewAndConfirm.vb",
    "PowerPointAi\ChatControl.vb",
    "WordAi\ChatControl.vb"
)

foreach ($path in $taskRunInteropFiles) {
    if ($allowedTaskRunInteropFiles -notcontains $path) {
        $failures.Add("Task.Run + Office Interop file needs review: $path")
    }
}

if ($failures.Count -gt 0) {
    Write-Host "P0 guardrail failures:" -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host " - $failure" -ForegroundColor Red
    }
    throw "P0 guardrail audit failed with $($failures.Count) issue(s)."
}

Write-Host "P0 guardrails passed."
Write-Host "Oversized baseline files: $($largeFiles.Count)"
Write-Host "Task.Run + Interop baseline files: $($taskRunInteropFiles.Count)"
