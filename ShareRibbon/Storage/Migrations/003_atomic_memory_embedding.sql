ALTER TABLE atomic_memory ADD COLUMN embedding TEXT DEFAULT NULL;
UPDATE schema_version SET version = 3;
