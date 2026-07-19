param(
    [string]$OutputDirectory = "artifacts\professional-design-smoke\scene-preview"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$outputRoot = if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path $repoRoot $OutputDirectory }
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

@(
    "packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll",
    "ShareRibbon\bin\Debug\ShareRibbon.dll",
    "PowerPointAi\bin\Debug\Office.dll",
    "PowerPointAi\bin\Debug\PowerPointAi.dll"
) | ForEach-Object {
    [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot $_)) | Out-Null
}

$json = @'
{"designSystem":"modern-tech","slides":[
{"slideType":"cover","eyebrow":"AI NATIVE STRATEGY","title":"Office Agent: From Chat to Autonomous Execution","subtitle":"Professional design compiler, declarative operations, verified delivery"},
{"slideType":"kpi","title":"The execution loop creates measurable business value","keyMessage":"Efficiency, quality, and safety improve together","metrics":[{"value":"70%","label":"Time saved","description":"Less repetitive document work"},{"value":"3x","label":"Faster delivery","description":"From request to final output"},{"value":"0","label":"Unapproved effects","description":"Safety-gated execution"}]},
{"slideType":"process","title":"The Agent closes execution through five stages","keyMessage":"Every step is observable, repairable, and explainable","items":[{"title":"Context","body":"Read the active document"},{"title":"Plan","body":"Create a structured plan"},{"title":"Act","body":"Execute native tools"},{"title":"Observe","body":"Verify the real result"},{"title":"Repair","body":"Correct deviations"}]},
{"slideType":"architecture","title":"One execution chain carries every dynamic capability","keyMessage":"Shared control plane with host-specific executors","items":[{"title":"Experience","body":"Chat, Ribbon, Shortcut"},{"title":"Agent Core","body":"Harness, Kernel, Loop"},{"title":"Capability","body":"Skill, Tool, Catalog"},{"title":"Execution","body":"Safety, COM, Observer"}]},
{"slideType":"comparison","title":"AI Native Agents deliver outcomes, not suggestions","items":[{"title":"Traditional assistant","body":"Produces advice and text while users still complete the work."},{"title":"AI Native Agent","body":"Reads context, executes operations, observes results, and repairs."}]},
{"slideType":"closing","eyebrow":"NEXT STEP","title":"Make every presentation delivery-ready","subtitle":"Upgrade from demo generation to professional design and verification"}
]}
'@

$spec = [PowerPointAi.Design.DeckDesignSpec]::Parse([Newtonsoft.Json.Linq.JObject]::Parse($json))
$tokens = [PowerPointAi.Design.DesignSystemCatalog]::Resolve($spec.DesignSystem)
$reports = @()

for ($index = 0; $index -lt $spec.Slides.Count; $index++) {
    $plan = [PowerPointAi.Design.SlideLayoutEngine]::Compile(
        $spec.Slides[$index], $tokens, 960, 540, $index, $spec.Slides.Count)
    $report = [PowerPointAi.Design.PowerPointVisualVerifier]::PreflightAndRepair($plan, 960, 540)
    if (-not $report.Passed) { throw "Preflight failed for slide $($index + 1)" }

    $path = Join-Path $outputRoot ("slide-{0:00}-{1}.png" -f ($index + 1), $spec.Slides[$index].SlideType)
    [PowerPointAi.Design.ScenePlanPreviewRenderer]::RenderToPng($plan, $tokens, $path, 960, 540, 2)
    $reports += [pscustomobject]@{
        Slide = $index + 1
        Type = $spec.Slides[$index].SlideType
        Nodes = $plan.Nodes.Count
        Passed = $report.Passed
        Issues = $report.Issues.Count
        Repairs = $report.RepairCount
        Preview = $path
    }
}

$reports | Format-Table -AutoSize
Write-Host "PREVIEW=$outputRoot"
