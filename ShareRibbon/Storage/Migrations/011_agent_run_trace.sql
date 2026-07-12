CREATE TABLE IF NOT EXISTS agent_run (
  run_id TEXT PRIMARY KEY,
  turn_id TEXT,
  session_id TEXT,
  app_type TEXT,
  status TEXT,
  user_text TEXT,
  started_at TEXT,
  finished_at TEXT,
  final_message TEXT,
  error_code TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_session ON agent_run(session_id, started_at);
CREATE INDEX IF NOT EXISTS idx_agent_run_status ON agent_run(status, started_at);

CREATE TABLE IF NOT EXISTS agent_run_step (
  step_id TEXT PRIMARY KEY,
  run_id TEXT NOT NULL,
  seq INTEGER NOT NULL,
  tool_id TEXT,
  status TEXT,
  message TEXT,
  error_code TEXT,
  observation_json TEXT,
  started_at TEXT,
  finished_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_step_run ON agent_run_step(run_id, seq);
CREATE INDEX IF NOT EXISTS idx_agent_run_step_status ON agent_run_step(status, error_code);

UPDATE schema_version SET version = 11;
