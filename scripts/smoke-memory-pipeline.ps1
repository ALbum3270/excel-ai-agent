param(
    [switch]$Build,
    [string]$Configuration = "Debug",
    [string]$Platform = "AnyCPU"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

if ($Build) {
    if (-not (Test-Path $msbuild)) {
        throw "MSBuild not found: $msbuild"
    }
    & $msbuild (Join-Path $repoRoot "WordAi\WordAi.vbproj") /p:Configuration=$Configuration /p:Platform=$Platform
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed, exit code: $LASTEXITCODE"
    }
}

$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
if (-not (Test-Path (Join-Path $bin "ShareRibbon.dll"))) {
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-memory-pipeline.ps1 -Build"
}

Push-Location $bin
try {
    Add-Type -Path (Join-Path $bin "ShareRibbon.dll")

    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()
    $dbPath = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()

    $session = "smoke-" + [guid]::NewGuid().ToString("N")

    $userEvent = New-Object ShareRibbon.ConversationEventRecord
    $userEvent.SessionId = $session
    $userEvent.AppType = "Word"
    $userEvent.EventType = "message"
    $userEvent.Role = "user"
    $userEvent.Content = "Remember my default official document body style: FangSong No.3, first-line indent 2 chars, justified alignment."
    $userId = [ShareRibbon.AgentMemoryRepository]::AppendConversationEvent($userEvent)

    $assistantEvent = New-Object ShareRibbon.ConversationEventRecord
    $assistantEvent.SessionId = $session
    $assistantEvent.AppType = "Word"
    $assistantEvent.EventType = "message"
    $assistantEvent.Role = "assistant"
    $assistantEvent.Content = "I have saved your official document body formatting preference."
    $assistantId = [ShareRibbon.AgentMemoryRepository]::AppendConversationEvent($assistantEvent)

    $payload = New-Object Newtonsoft.Json.Linq.JObject
    $payload.Add("user_event_id", [Newtonsoft.Json.Linq.JValue]::CreateString($userId))
    $payload.Add("assistant_event_id", [Newtonsoft.Json.Linq.JValue]::CreateString($assistantId))

    $job = New-Object ShareRibbon.MemoryJobRecord
    $job.JobType = "extract_memory"
    $job.TargetId = $assistantId
    $job.PayloadJson = $payload.ToString([Newtonsoft.Json.Formatting]::None)
    $job.Status = "pending"
    [ShareRibbon.AgentMemoryRepository]::EnqueueJob($job) | Out-Null

    [ShareRibbon.AgentMemoryPipelineService]::ProcessPendingJobsAsync(50).GetAwaiter().GetResult()

    $memories = [ShareRibbon.AgentMemoryRepository]::RetrieveMemoryItems("official document FangSong", "Word", "", "", 10)
    $skills = [ShareRibbon.AgentMemoryRepository]::RetrieveSkillRegistry("", "", "Word", 5)

    $result = [pscustomobject]@{
        Database = $dbPath
        Session = $session
        UserEventId = $userId
        AssistantEventId = $assistantId
        MemoryCount = $memories.Count
        FirstMemory = $(if ($memories.Count -gt 0) { $memories[0].Content } else { "" })
        SkillRegistryRows = $skills.Count
    }

    $result | Format-List

    if ($memories.Count -lt 1) {
        throw "Smoke failed: the inserted structured memory was not retrieved."
    }
}
finally {
    Pop-Location
}
