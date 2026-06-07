ALTER TABLE atomic_memory ADD COLUMN embedding_model TEXT DEFAULT '';
ALTER TABLE atomic_memory ADD COLUMN embedding_dim INTEGER DEFAULT 0;
ALTER TABLE atomic_memory ADD COLUMN embedding_updated_at TEXT;
CREATE INDEX IF NOT EXISTS idx_atomic_memory_rag ON atomic_memory(memory_type, app_type, timestamp);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_embedding_present ON atomic_memory(memory_type, app_type) WHERE embedding IS NOT NULL AND embedding != '';
UPDATE schema_version SET version = 9;
