CREATE TABLE IF NOT EXISTS format_template (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    description TEXT,
    category TEXT DEFAULT '通用',
    target_app TEXT DEFAULT 'Word',
    is_preset INTEGER NOT NULL DEFAULT 0,
    template_source TEXT DEFAULT 'manual',
    source_file_name TEXT,
    source_file_content TEXT,
    source_file_blob BLOB,
    layout_json TEXT,
    style_rules_json TEXT,
    page_settings_json TEXT,
    ai_guidance TEXT,
    thumbnail_base64 TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    last_modified TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_format_template_target_app ON format_template(target_app);
CREATE INDEX IF NOT EXISTS idx_format_template_category ON format_template(category);

CREATE TABLE IF NOT EXISTS format_element (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id TEXT NOT NULL,
    element_name TEXT NOT NULL,
    element_type TEXT NOT NULL,
    default_value TEXT,
    is_required INTEGER DEFAULT 1,
    sort_order INTEGER DEFAULT 0,
    font_config_json TEXT,
    paragraph_config_json TEXT,
    color_config_json TEXT,
    special_props_json TEXT,
    placeholder_content TEXT,
    is_enabled INTEGER DEFAULT 1,
    FOREIGN KEY (template_id) REFERENCES format_template(template_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_format_element_template ON format_element(template_id);

CREATE TABLE IF NOT EXISTS format_style_rule (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id TEXT NOT NULL,
    rule_name TEXT NOT NULL,
    match_condition TEXT,
    sort_order INTEGER DEFAULT 0,
    font_config_json TEXT,
    paragraph_config_json TEXT,
    color_config_json TEXT,
    is_enabled INTEGER DEFAULT 1,
    FOREIGN KEY (template_id) REFERENCES format_template(template_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_format_style_rule_template ON format_style_rule(template_id);
UPDATE schema_version SET version = 4;
