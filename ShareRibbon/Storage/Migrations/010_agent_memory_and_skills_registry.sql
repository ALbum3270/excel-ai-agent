CREATE TABLE IF NOT EXISTS conversation_event (
    event_id TEXT PRIMARY KEY,
    session_id TEXT NOT NULL,
    app_type TEXT DEFAULT '',
    document_id TEXT DEFAULT '',
    event_type TEXT NOT NULL,
    role TEXT DEFAULT '',
    content TEXT,
    metadata_json TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    processed_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_conversation_event_session ON conversation_event(session_id, created_at);
CREATE INDEX IF NOT EXISTS idx_conversation_event_scope ON conversation_event(app_type, document_id, event_type);
CREATE INDEX IF NOT EXISTS idx_conversation_event_processed ON conversation_event(processed_at);

CREATE TABLE IF NOT EXISTS memory_item (
    memory_id TEXT PRIMARY KEY,
    source_event_id TEXT,
    scope TEXT NOT NULL DEFAULT 'user',
    app_type TEXT DEFAULT '',
    document_id TEXT DEFAULT '',
    project_id TEXT DEFAULT '',
    memory_type TEXT NOT NULL DEFAULT 'fact',
    content TEXT NOT NULL,
    summary TEXT,
    confidence REAL DEFAULT 0.5,
    importance REAL DEFAULT 0.5,
    status TEXT NOT NULL DEFAULT 'active',
    expires_at TEXT,
    last_verified_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (source_event_id) REFERENCES conversation_event(event_id)
);

CREATE INDEX IF NOT EXISTS idx_memory_item_scope ON memory_item(scope, app_type, document_id, status);
CREATE INDEX IF NOT EXISTS idx_memory_item_type ON memory_item(memory_type, status, importance);
CREATE INDEX IF NOT EXISTS idx_memory_item_source ON memory_item(source_event_id);

CREATE TABLE IF NOT EXISTS memory_embedding (
    embedding_id TEXT PRIMARY KEY,
    memory_id TEXT NOT NULL,
    embedding_model TEXT NOT NULL,
    embedding_dim INTEGER DEFAULT 0,
    embedding_json TEXT,
    vector_status TEXT NOT NULL DEFAULT 'pending',
    last_error TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (memory_id) REFERENCES memory_item(memory_id) ON DELETE CASCADE
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_memory_embedding_unique ON memory_embedding(memory_id, embedding_model);
CREATE INDEX IF NOT EXISTS idx_memory_embedding_status ON memory_embedding(vector_status, embedding_model);

CREATE TABLE IF NOT EXISTS memory_job (
    job_id TEXT PRIMARY KEY,
    job_type TEXT NOT NULL,
    target_id TEXT,
    payload_json TEXT,
    status TEXT NOT NULL DEFAULT 'pending',
    attempt_count INTEGER DEFAULT 0,
    last_error TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    next_run_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_memory_job_pending ON memory_job(status, next_run_at, created_at);
CREATE INDEX IF NOT EXISTS idx_memory_job_type ON memory_job(job_type, status);

CREATE TABLE IF NOT EXISTS skills_registry (
    skill_name TEXT PRIMARY KEY,
    file_path TEXT NOT NULL,
    app_scope TEXT DEFAULT '',
    intent_types TEXT DEFAULT '',
    trigger_keywords TEXT DEFAULT '',
    description TEXT,
    embedding_json TEXT,
    usage_count INTEGER DEFAULT 0,
    success_count INTEGER DEFAULT 0,
    last_indexed_at TEXT,
    enabled INTEGER NOT NULL DEFAULT 1
);

CREATE INDEX IF NOT EXISTS idx_skills_registry_enabled ON skills_registry(enabled, app_scope);
CREATE INDEX IF NOT EXISTS idx_skills_registry_usage ON skills_registry(usage_count, success_count);
UPDATE schema_version SET version = 10;
