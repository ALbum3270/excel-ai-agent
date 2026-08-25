param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\build-code-projects.ps1 -Configuration $Configuration -Projects `"ShareRibbon\ShareRibbon.vbproj`""
}

Push-Location $bin
try {
    Add-Type -Path $dll

    $openAiOptions = New-Object ShareRibbon.AiRequestOptions
    $openAiOptions.ApiUrl = "https://api.openai.example/v1/chat/completions"
    $openAiOptions.ApiKey = "smoke-key"
    $openAiOptions.ModelName = "gpt-smoke"
    $openAiOptions.Platform = "OpenAI"
    $openAiOptions.SystemPrompt = "system prompt"
    $openAiOptions.UserPrompt = "hello"
    $openAiOptions.Temperature = 0.2
    $openAiOptions.MaxTokens = 123

    $openAiRequest = [ShareRibbon.AiGateway]::BuildProviderRequest($openAiOptions)
    if ($openAiRequest["model"].ToString() -ne "gpt-smoke") { throw "OpenAI model was not preserved." }
    if ($openAiRequest["stream"].ToString().ToLowerInvariant() -ne "false") { throw "OpenAI stream must be false." }
    if ($openAiRequest["messages"].Count -ne 2) { throw "OpenAI messages count mismatch." }
    if ($openAiRequest["temperature"].ToString() -ne "0.2") { throw "OpenAI temperature mismatch." }
    if ($openAiRequest["max_tokens"].ToString() -ne "123") { throw "OpenAI max_tokens mismatch." }

    $anthropicOptions = New-Object ShareRibbon.AiRequestOptions
    $anthropicOptions.ApiUrl = "https://api.anthropic.com/v1/messages"
    $anthropicOptions.ApiKey = "smoke-key"
    $anthropicOptions.ModelName = "claude-smoke"
    $anthropicOptions.Platform = "Anthropic"
    $anthropicOptions.Messages = [Newtonsoft.Json.Linq.JArray]::Parse('[{"role":"system","content":"anthropic system"},{"role":"user","content":"hello"}]')
    $anthropicOptions.MaxTokens = 321

    $anthropicRequest = [ShareRibbon.AiGateway]::BuildProviderRequest($anthropicOptions)
    if ($anthropicRequest["model"].ToString() -ne "claude-smoke") { throw "Anthropic model was not preserved." }
    if ($anthropicRequest["stream"].ToString().ToLowerInvariant() -ne "false") { throw "Anthropic stream must be false." }
    if ($anthropicRequest["max_tokens"].ToString() -ne "321") { throw "Anthropic max_tokens mismatch." }
    if ($anthropicRequest["system"].ToString() -ne "anthropic system") { throw "Anthropic system message was not lifted." }
    if ($anthropicRequest["messages"].Count -ne 1) { throw "Anthropic user messages count mismatch." }
    if ($anthropicRequest["messages"][0]["role"].ToString() -ne "user") { throw "Anthropic user role mismatch." }

    $anthropicResponse = '{"content":[{"type":"text","text":"hello from claude"}]}'
    $extracted = [ShareRibbon.AiGateway]::ExtractAssistantContent($anthropicResponse, $true)
    if ($extracted -ne "hello from claude") { throw "Anthropic response extraction mismatch." }

    $arrearageBody = '{"error":{"message":"Access denied because the account is in arrears.","type":"Arrearage","code":"Arrearage"},"request_id":"smoke-request"}'
    $providerException = [ShareRibbon.AiGateway]::CreateProviderHttpException(
        400,
        "Bad Request",
        $arrearageBody)
    $classifiedProviderError = [ShareRibbon.ExceptionClassifier]::Classify($providerException)
    if ($providerException.ProviderErrorCode -ne "Arrearage") { throw "Provider error code was discarded." }
    if ($classifiedProviderError.ErrorCode -ne ([ShareRibbon.ExceptionClassifier]::CodeProviderAccount)) { throw "Provider account error was misclassified: $($classifiedProviderError.ErrorCode)" }
    if ($classifiedProviderError.ErrorCode -eq ([ShareRibbon.ExceptionClassifier]::CodeNetwork)) { throw "Provider account error was flattened into NETWORK_ERROR." }
    if ($classifiedProviderError.Retryable) { throw "Provider account error was marked retryable." }
    if ([string]::IsNullOrWhiteSpace($classifiedProviderError.UserMessage) -or $classifiedProviderError.UserMessage.Length -lt 20) { throw "Provider account error lost its actionable user message." }

    $authException = [ShareRibbon.AiGateway]::CreateProviderHttpException(
        401,
        "Unauthorized",
        '{"error":{"message":"Invalid API key","code":"invalid_api_key"}}')
    $authError = [ShareRibbon.ExceptionClassifier]::Classify($authException)
    if ($authError.ErrorCode -ne ([ShareRibbon.ExceptionClassifier]::CodeProviderAuth) -or $authError.Retryable) {
        throw "Provider authentication failure classification is incorrect."
    }

    $rateException = [ShareRibbon.AiGateway]::CreateProviderHttpException(
        429,
        "Too Many Requests",
        '{"error":{"message":"Rate limit reached","code":"rate_limit_exceeded"}}')
    $rateError = [ShareRibbon.ExceptionClassifier]::Classify($rateException)
    if ($rateError.ErrorCode -ne ([ShareRibbon.ExceptionClassifier]::CodeProviderRateLimited) -or -not $rateError.Retryable) {
        throw "Provider rate-limit classification is incorrect."
    }

    [pscustomobject]@{
        OpenAiMessages = $openAiRequest["messages"].Count
        AnthropicMessages = $anthropicRequest["messages"].Count
        AnthropicSystem = $anthropicRequest["system"].ToString()
        Extracted = $extracted
        ProviderError = $classifiedProviderError.ErrorCode
    } | Format-List
}
finally {
    Pop-Location
}
