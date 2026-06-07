ALTER TABLE atomic_memory ADD COLUMN app_type TEXT DEFAULT '';
CREATE INDEX IF NOT EXISTS idx_atomic_memory_app_type ON atomic_memory(app_type);
UPDATE schema_version SET version = 2;
