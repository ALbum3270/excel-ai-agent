ALTER TABLE atomic_memory ADD COLUMN memory_type TEXT DEFAULT 'short_term';
CREATE INDEX IF NOT EXISTS idx_atomic_memory_memory_type ON atomic_memory(memory_type);
UPDATE schema_version SET version = 6;
