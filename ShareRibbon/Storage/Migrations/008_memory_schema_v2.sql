CREATE TABLE IF NOT EXISTS atomic_memory_new (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp INTEGER NOT NULL,
    content TEXT NOT NULL,
    tags TEXT,
    session_id TEXT,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    app_type TEXT DEFAULT '',
    embedding TEXT,
    memory_type TEXT DEFAULT 'short_term',
    importance REAL DEFAULT 0.5,
    access_count INTEGER DEFAULT 0,
    last_access TEXT,
    source_type TEXT DEFAULT 'general',
    linked_memories TEXT
);

INSERT OR IGNORE INTO atomic_memory_new
    (id, timestamp, content, tags, session_id, create_time, app_type, embedding, memory_type)
SELECT id, timestamp, content, tags, session_id, create_time, app_type, embedding, memory_type
FROM atomic_memory;

DROP TABLE atomic_memory;
ALTER TABLE atomic_memory_new RENAME TO atomic_memory;

CREATE INDEX IF NOT EXISTS idx_atomic_memory_content ON atomic_memory(content);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_timestamp ON atomic_memory(timestamp);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_app_type ON atomic_memory(app_type);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_memory_type ON atomic_memory(memory_type);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_importance ON atomic_memory(importance);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_source_type ON atomic_memory(source_type);

CREATE TABLE IF NOT EXISTS memory_graph (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER NOT NULL,
    target_id INTEGER NOT NULL,
    relation_type TEXT NOT NULL DEFAULT 'similar',
    weight REAL DEFAULT 1.0,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_memory_graph_source ON memory_graph(source_id);
CREATE INDEX IF NOT EXISTS idx_memory_graph_target ON memory_graph(target_id);
CREATE INDEX IF NOT EXISTS idx_memory_graph_relation ON memory_graph(relation_type);

CREATE TABLE IF NOT EXISTS user_profile_new (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key TEXT NOT NULL UNIQUE,
    value TEXT NOT NULL,
    category TEXT DEFAULT 'preference',
    confidence REAL DEFAULT 0.5,
    last_updated TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    observation_count INTEGER DEFAULT 1
);

INSERT OR IGNORE INTO user_profile_new (key, value)
SELECT 'legacy_data', content FROM user_profile WHERE content IS NOT NULL LIMIT 1;

DROP TABLE IF EXISTS user_profile;
ALTER TABLE user_profile_new RENAME TO user_profile;

CREATE INDEX IF NOT EXISTS idx_user_profile_category ON user_profile(category);
CREATE INDEX IF NOT EXISTS idx_user_profile_key ON user_profile(key);

CREATE TABLE IF NOT EXISTS conversation_branch (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    conversation_id INTEGER NOT NULL,
    parent_message_id INTEGER,
    branch_name TEXT,
    is_active INTEGER DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_conversation_branch_conv ON conversation_branch(conversation_id);
CREATE INDEX IF NOT EXISTS idx_conversation_branch_active ON conversation_branch(is_active);
UPDATE schema_version SET version = 8;
