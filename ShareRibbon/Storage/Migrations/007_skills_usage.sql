CREATE TABLE IF NOT EXISTS skills_usage (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    skill_name TEXT NOT NULL,
    usage_count INTEGER DEFAULT 0,
    success_count INTEGER DEFAULT 0,
    total_tokens INTEGER DEFAULT 0,
    last_used_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_skills_usage_name ON skills_usage(skill_name);
UPDATE schema_version SET version = 7;
