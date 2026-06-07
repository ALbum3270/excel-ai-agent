# Global Architecture Hardening Execution

## 2026-06-07 Phase 1

### Scope

- Added `ShareRibbon/Services/Ai/AiGateway.vb`.
- Registered `Services\Ai\AiGateway.vb` in `ShareRibbon/ShareRibbon.vbproj`.
- Migrated `ShareRibbon/Services/Memory/LlmMemoryExtractor.vb` to call `AiGateway.SendChatAsync`.
- Centralized non-streaming OpenAI-compatible request building, Anthropic `/v1/messages` conversion, reasoning option application, timeout handling, and response text extraction.
- Removed the LLM memory extractor's Anthropic hard block; provider adaptation now belongs to `AiGateway`.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `git diff --check`: only existing line-ending warnings.
- VB syntax scan: no new C# syntax in the new VB code; matches are existing embedded JavaScript strings.

### Follow-Up

- Phase 2 should add a shared UI dispatcher before further WebView2 or Office Interop changes.
- Phase 3 should keep converging old memory and skills usage paths into the structured SQLite sidecar.

## 2026-06-07 Phase 2

### Scope

- Added `ShareRibbon/Services/Ui/UiDispatcher.vb`.
- Registered `Services\Ui\UiDispatcher.vb` in `ShareRibbon/ShareRibbon.vbproj`.
- Migrated high-risk WebView2/UI entry points to `UiDispatcher`:
  - `BaseChatControl.InitializeWebView2`
  - `BaseChatControl.ExecuteJavaScriptAsyncJS`
  - `BaseChatControl.ExecuteJavaScriptAndWaitAsync`
  - `BaseChatControl.WaitForRendererMapAsync`
  - `BaseChatControl.LoadLocalHtmlFile`
  - `BaseChatControl.GetFullHtmlContentAsync`
  - `ConfigApiForm.InitializeSkillsWebView2`
  - `ConfigApiForm.RunAfterHandleCreated`
- Replaced service callback lambdas using direct `Me.Invoke` with `BaseChatControl.RunUiAction`.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- Targeted VB syntax scan for new/changed service files: no C# syntax matches.
- `git diff --check`: only existing line-ending warnings.

### Remaining Work

- `WebViewService.vb` has no current instantiation found by `rg`; keep it compiling, but consider deleting or merging it after a broader dead-code review.

## 2026-06-07 Phase 2 Follow-Up

### Scope

- Replaced remaining direct `Me.Invoke` usages in `BaseChatControl` with `RunUiActionSync`.
- Replaced `BaseChatControl.OnWebViewNavigationCompleted` direct `ChatBrowser.Invoke` with `UiDispatcher.InvokeAsync`.
- Replaced `ConfigApiForm.SkillsTab_SelectedIndexChanged` delayed `BeginInvoke` with `RunAfterHandleCreated`.
- Migrated `WebViewService.InitializeAsync`, `ExecuteScriptAsync`, and `InvokeIfRequired` to `UiDispatcher`.
- Confirmed `WebViewService` currently has no instantiation references in `ShareRibbon`, `WordAi`, `ExcelAi`, or `PowerPointAi`.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- Targeted scan for direct `.Invoke(` / `BeginInvoke(` in `BaseChatControl`, `ConfigApiForm`, and `WebViewService`: no matches outside `UiDispatcher`.
- Targeted VB syntax scan for `Services\Ui` and `WebViewService`: no C# syntax matches.
- `git diff --check`: only existing line-ending warnings.

### Remaining Work

- Run a real Word UI smoke test for opening chat, opening the API config form, opening the Skills tab, sending a message, and attaching a file.
- Phase 3 can now focus on Memory/Skills convergence and migration-file extraction.

## 2026-06-07 Phase 3 Skills Usage Convergence

### Scope

- Made `skills_registry` the primary source for Skill usage stats during catalog loading.
- Updated `SkillsService.RecordSkillUsage` to write:
  - primary: `skills_registry`
  - compatibility: legacy `skills_usage` SQLite table
  - compatibility/cache: existing `skills_usage.json` batch cache
- Updated `EnhancedSkillsService` to read usage stats from `skills_registry` first, then fall back to legacy `skills_usage`.
- Updated `EnhancedSkillsService.RecordSkillFeedback` to route through `SkillsService.RecordSkillUsage`.
- Hardened `AgentMemoryRepository.RecordSkillRegistryUsage` so first usage inserts a minimal registry row if the Skill has not been indexed yet.
- Added `scripts/smoke-skills-registry.ps1`.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-skills-registry.ps1`: passed.
  - `RegistryUsage = 1`
  - `RegistrySuccess = 1`
  - `LegacyUsage = 1`
  - `LegacySuccess = 1`
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- Targeted VB syntax scan for changed Skills code: no C# syntax matches.
- `git diff --check`: only existing line-ending warnings.

### Remaining Work

- Add a one-time migration to copy existing `skills_usage` and `skills_usage.json` counts into `skills_registry` for old users.
- Consider deprecating `skills_usage.json` after one or two release cycles once registry migration is proven.

## 2026-06-07 Phase 3 Follow-Up: Legacy Skills Usage Migration

### Scope

- Added `data_migration_marker` in `OfficeAiDatabase.RunSchemaHealthChecks`.
- Added one-time migration from legacy SQLite `skills_usage` into `skills_registry`.
- Added one-time migration from legacy `skills_usage.json` into `skills_registry` when the JSON file exists and parses successfully.
- Migrations merge counters by taking the higher existing/source value instead of adding both sources, avoiding double-counting when JSON and SQLite contain duplicate historical stats.
- Added `scripts/smoke-skills-usage-migration.ps1` to verify table migration and startup idempotency.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-skills-usage-migration.ps1`: passed.
  - `RegistryUsage = 7`
  - `RegistrySuccess = 5`
  - `Idempotent = True`
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `git diff --check`: only existing line-ending warnings.

### Remaining Work

- Keep writing compatibility stats to `skills_usage` and `skills_usage.json` for now.
- After one or two stable release cycles, remove the JSON cache write path first, then consider removing the legacy `skills_usage` read fallback.

## 2026-06-07 Phase 4 Start: Database Baseline SQL Extraction

### Scope

- Added `ShareRibbon/Storage/Migrations/001_baseline.sql` as the external version-1 database baseline.
- Updated `OfficeAiDatabase.GetMigrationSql()` to prefer external SQL from the build output before falling back to the inline baseline string.
- Kept the inline baseline string as a runtime fallback so missing copied SQL does not break startup.
- Registered `Storage\Migrations\001_baseline.sql` and `Storage\OfficeAiDbSchema.current.sql` in `ShareRibbon.vbproj` with `CopyToOutputDirectory=PreserveNewest`.
- Updated the current schema snapshot to include `data_migration_marker`.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- Build copied SQL artifacts to:
  - `ShareRibbon\bin\Debug\Storage\Migrations\001_baseline.sql`
  - `WordAi\bin\Debug\Storage\Migrations\001_baseline.sql`
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-skills-usage-migration.ps1`: passed, `Idempotent = True`.

### Remaining Work

- Continue extracting versioned migration SQL (`002` through `010`) or move them behind a structured migration-provider abstraction.
- Add a dedicated empty-database initialization smoke that creates a temporary isolated DB instead of using the normal Debug user database.

## 2026-06-08 Phase 4 Follow-Up: Versioned Migration SQL Files

### Scope

- Added external versioned SQL migrations:
  - `ShareRibbon/Storage/Migrations/002_atomic_memory_app_type.sql`
  - `ShareRibbon/Storage/Migrations/003_atomic_memory_embedding.sql`
  - `ShareRibbon/Storage/Migrations/004_format_templates.sql`
  - `ShareRibbon/Storage/Migrations/005_prompt_template_metadata.sql`
  - `ShareRibbon/Storage/Migrations/006_atomic_memory_type.sql`
  - `ShareRibbon/Storage/Migrations/007_skills_usage.sql`
  - `ShareRibbon/Storage/Migrations/008_memory_schema_v2.sql`
  - `ShareRibbon/Storage/Migrations/009_atomic_memory_embedding_metadata.sql`
  - `ShareRibbon/Storage/Migrations/010_agent_memory_and_skills_registry.sql`
- Updated `OfficeAiDatabase.RunVersionedMigrations` to read `Storage\Migrations\NNN_*.sql` first and fall back to the existing inline SQL per version.
- Updated `ShareRibbon.vbproj` to copy `Storage\Migrations\*.sql`, so future migration files do not require per-file project edits.
- Added `scripts/smoke-empty-db-initialization.ps1`.
  - The script temporarily moves the Debug DB aside.
  - It initializes an empty DB in a child PowerShell process so SQLite file handles are released before restore.
  - It verifies `schema_version = 10` and required core tables.
  - It restores the original Debug DB after the smoke.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- Build copied `001` through `010` SQL files to `ShareRibbon\bin\Debug\Storage\Migrations` and `WordAi\bin\Debug\Storage\Migrations`.
- `scripts/smoke-empty-db-initialization.ps1`: passed.
  - `SchemaVersion = 10`
  - `RequiredTables = 8`
  - `EmptyInitialization = True`
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-skills-usage-migration.ps1`: passed, `Idempotent = True`.

### Remaining Work

- Inline version SQL can now be reduced to documented fallback constants, or eventually removed after packaged deployment proves external SQL files are always present.
- Add a DB schema snapshot comparison smoke if we want to detect drift between `OfficeAiDbSchema.current.sql`, external migrations, and health checks.

## 2026-06-08 Phase 4 Follow-Up: Schema Drift Smoke

### Scope

- Added `scripts/smoke-db-schema-drift.ps1`.
- The smoke creates two temporary SQLite databases:
  - one from the external migration chain `001` through `010`
  - one from `OfficeAiDbSchema.current.sql`
- It compares:
  - user tables
  - column name/type/not-null/primary-key shape
  - explicit index names
- It intentionally does not compare column default values yet, because legacy SQL contains encoded historical text defaults that should be normalized separately from structural drift checks.

### Validation

- `scripts/smoke-db-schema-drift.ps1`: passed.
  - `Tables = 18`
  - `Columns = 168`
  - `Indexes = 35`
  - `Drift = False`

### Remaining Work

- Decide whether to normalize historical encoded SQL defaults, then optionally extend drift checks to include default values.
- Consider cleaning up or archiving old inline migration strings once external SQL deployment is proven in an installed VSTO package, not only Debug build output.

## 2026-06-08 Phase 5 Start: WebView Command Router

### Scope

- Added `ShareRibbon/Controls/Services/ChatCommandRouter.vb`.
- Registered `Controls\Services\ChatCommandRouter.vb` in `ShareRibbon.vbproj`.
- Added lazy `CommandRouter` initialization in `BaseChatControl`.
- Moved WebView message type dispatch from a large `Select Case` block to table-driven command registration via `RegisterWebViewCommandHandlers`.
- Simplified `BaseChatControl.WebView2_WebMessageReceived` to:
  - read raw WebView JSON
  - parse `JObject`
  - dispatch through `ChatCommandRouter`
- Removed the old `Select Case` dispatch block after compile validation.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-db-schema-drift.ps1`: passed, `Drift = False`.

### Remaining Work

- Run a real Word UI smoke because WebView command routing cannot be fully exercised by current PowerShell smokes.
- Next Phase 5 slice should extract WebView script execution/navigation into a `WebViewBridge` while keeping `BaseChatControl` handlers intact.

## 2026-06-08 Phase 5 Follow-Up: WebView Bridge

### Scope

- Added `ShareRibbon/Controls/Services/WebViewBridge.vb`.
- Registered `Controls\Services\WebViewBridge.vb` in `ShareRibbon.vbproj`.
- Added lazy `WebViewBridge` initialization in `BaseChatControl`.
- Kept existing `BaseChatControl` method signatures for compatibility while delegating implementation to `WebViewBridge`:
  - `ExecuteJavaScriptAsyncJS`
  - `ExecuteJavaScriptAndWaitAsync`
  - `WaitForRendererMapAsync`
  - `LoadLocalHtmlFile`
  - `GetFullHtmlContentAsync`
- Moved full-page HTML extraction and cleanup logic from `BaseChatControl` into `WebViewBridge`.
- Removed direct `Regex` dependency from `BaseChatControl`.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-db-schema-drift.ps1`: passed, `Drift = False`.

### Remaining Work

- Run real Word UI smoke for:
  - opening the chat pane
  - sending a message
  - opening API settings
  - loading/saving a chat HTML snapshot
  - opening reformat/proofread flows
- Next slice can extract `Send` request orchestration or move WebView initialization itself out of `BaseChatControl`.

## 2026-06-08 Phase 5 Follow-Up: System Prompt Resolver

### Scope

- Added `ShareRibbon/Controls/Services/ChatSystemPromptResolver.vb`.
- Registered `Controls\Services\ChatSystemPromptResolver.vb` in `ShareRibbon.vbproj`.
- Added lazy `SystemPromptResolver` initialization in `BaseChatControl`.
- Moved `Send` system prompt fallback selection out of `BaseChatControl`.
- Preserved the existing prompt resolution order:
  - caller-provided system prompt
  - `PromptManager.Instance.GetCombinedPrompt(context)`
  - `ConfigSettings.propmtContent`
  - final Office AI assistant fallback text

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-db-schema-drift.ps1`: passed, `Drift = False`.
- `git diff --check`: no whitespace errors; only existing LF/CRLF conversion warnings.

### Remaining Work

- Run a real Word UI smoke for prompt selection paths:
  - normal chat with default prompt
  - chat after custom prompt configuration
  - reformat/proofread flows that pass explicit prompts
- Continue extracting `Send` request validation and request orchestration so `BaseChatControl` keeps shrinking toward UI lifecycle/event binding only.

## 2026-06-08 Phase 5 Follow-Up: Send Validation

### Scope

- Added `ShareRibbon/Controls/Services/ChatSendValidator.vb`.
- Registered `Controls\Services\ChatSendValidator.vb` in `ShareRibbon.vbproj`.
- Added lazy `SendValidator` initialization in `BaseChatControl`.
- Replaced the three inline `Send` preflight branches for missing API key, missing API URL, and empty question with one validator call.
- Kept user-facing behavior unchanged:
  - same warning text
  - same `changeSendButton()` recovery call
  - same early return behavior before request UUID creation

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-db-schema-drift.ps1`: passed, `Drift = False`.
- `git diff --check`: no whitespace errors; only existing LF/CRLF conversion warnings.

### Remaining Work

- Expand `ChatSendValidator` later to validate model/provider capability combinations, especially reasoning-mode settings and provider-specific required fields.
- Continue Phase 5 by extracting request-body creation or finalize-history persistence from `Send`.

## 2026-06-08 Phase 5 Follow-Up: Memory Turn Recorder

### Scope

- Added `ShareRibbon/Controls/Services/MemoryTurnRecorder.vb`.
- Registered `Controls\Services\MemoryTurnRecorder.vb` in `ShareRibbon.vbproj`.
- Added lazy `MemoryTurnRecorder` initialization in `BaseChatControl`.
- Moved Agent memory turn persistence out of `BaseChatControl`:
  - appends user and assistant `conversation_event` rows
  - enqueues an `extract_memory` job
  - kicks off pending memory jobs
- Preserved the old asynchronous fire-and-forget behavior and captured values before background execution.

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-db-schema-drift.ps1`: passed, `Drift = False`.
- `git diff --check`: no whitespace errors; only existing LF/CRLF conversion warnings.
- Targeted scan confirms `PersistAgentMemoryTurnAsync` was removed from `BaseChatControl` and only `MemoryTurnRecorder.RecordConversationTurn` remains.

### Remaining Work

- Consider moving `MemoryTurnRecorder` under `Services\Memory` once the Phase 5 UI-control split is complete; it currently sits near chat services because it is called by the stream finalization path.
- Add a dedicated small smoke for `MemoryTurnRecorder` if we want to validate it without running the full memory pipeline smoke.

## 2026-06-08 Phase 5 Follow-Up: WebView Protocol Audit

### Scope

- Audited `ShareRibbon/Resources/chat-template-refactored.html` and `ShareRibbon/Resources/js/*.js` for WebView message protocol usage.
- Confirmed this refactor added no new html/js/css files, so `ResourceExtractor` and virtual-server static resource registration do not need updates.
- Added compatibility router registrations for legacy frontend protocol entries found during the audit:
  - `startLoop`
  - `continueLoop`
  - `replanLoop`
  - `cancelLoop`
  - `getCurrentAppInfo`
- Added `HandleLegacyStartLoop`, mapping legacy `/loop` requests to the existing Agent start flow.
- Added `HandleGetCurrentAppInfo`, refreshing `window.currentOfficeAppName` for the frontend fallback path.
- Added detailed landing document:
  - `openspec/changes/phase5-chat-control-refactor-and-webview-audit.md`

### Validation

- `MSBuild WordAi\WordAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `MSBuild ExcelAi\ExcelAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- `MSBuild PowerPointAi\PowerPointAi.vbproj /p:Configuration=Debug /p:Platform=AnyCPU`: passed, 0 warnings, 0 errors.
- WebView frontend message type scan filtered by `postMessage/sendMessageToServer/sendMessageToVB` context: no remaining missing router registrations.
- `scripts/smoke-memory-pipeline.ps1`: passed, `MemoryCount = 1`.
- `scripts/smoke-skills-registry.ps1`: passed.
- `scripts/smoke-db-schema-drift.ps1`: passed, `Drift = False`.
- `git diff --check`: no whitespace errors; only existing LF/CRLF conversion warnings.

### Notes

- A parallel three-project build hit a transient `ShareRibbon.dll` copy lock for PowerPoint. Re-running `PowerPointAi.vbproj` alone passed; this was a build artifact contention issue, not a compile failure.
