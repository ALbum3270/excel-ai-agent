ALTER TABLE prompt_template ADD COLUMN description TEXT DEFAULT '';
ALTER TABLE prompt_template ADD COLUMN keywords TEXT DEFAULT '';
ALTER TABLE prompt_template ADD COLUMN category TEXT DEFAULT '';
ALTER TABLE prompt_template ADD COLUMN priority INTEGER DEFAULT 50;
ALTER TABLE prompt_template ADD COLUMN enabled INTEGER DEFAULT 1;
ALTER TABLE prompt_template ADD COLUMN parameters_json TEXT DEFAULT '';
UPDATE schema_version SET version = 5;
