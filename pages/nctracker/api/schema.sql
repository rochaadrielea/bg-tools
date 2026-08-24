-- NC Tracker schema. Runs once, on an empty database.

CREATE TABLE IF NOT EXISTS nc (
  id                SERIAL PRIMARY KEY,
  system            TEXT DEFAULT '',
  id_blackout       TEXT DEFAULT '',
  nc_type           TEXT DEFAULT '',
  tc_id             TEXT DEFAULT '',
  migrated          TEXT DEFAULT '',
  project           TEXT DEFAULT '',
  flight_unit       TEXT DEFAULT '',
  detection         TEXT DEFAULT '',
  title             TEXT DEFAULT '',
  failure           TEXT DEFAULT '',
  material          TEXT DEFAULT '',
  batch             TEXT DEFAULT '',
  owner             TEXT DEFAULT '',
  created_on        TEXT DEFAULT '',
  nrb_disposition   TEXT DEFAULT '',
  disposition_date  TEXT DEFAULT '',
  classification    TEXT DEFAULT '',
  root_cause        TEXT DEFAULT '',
  responsible_area  TEXT DEFAULT '',
  psp_ref           TEXT DEFAULT '',
  nc_wbs            TEXT DEFAULT '',
  status            TEXT DEFAULT '',
  closure_date      TEXT DEFAULT '',
  supplier          TEXT DEFAULT '',
  match_key         TEXT DEFAULT '',
  updated_at        TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS nc_system_idx    ON nc (system);
CREATE INDEX IF NOT EXISTS nc_match_key_idx ON nc (match_key);

-- the Set-up lists. Adding a value here makes it available to everyone.
CREATE TABLE IF NOT EXISTS setup_value (
  id         SERIAL PRIMARY KEY,
  list_name  TEXT NOT NULL,
  value      TEXT NOT NULL,
  sort_order INT  DEFAULT 0,
  UNIQUE (list_name, value)
);

-- every uploaded file, kept byte-for-byte. This is what the download buttons serve.
CREATE TABLE IF NOT EXISTS upload (
  id          SERIAL PRIMARY KEY,
  kind        TEXT NOT NULL,          -- 'tc_report' | 'nc_sap_overview' | 'old_excel'
  filename    TEXT NOT NULL,
  uploaded_at TIMESTAMPTZ DEFAULT now(),
  content     BYTEA NOT NULL
);
CREATE INDEX IF NOT EXISTS upload_kind_idx ON upload (kind, uploaded_at DESC);

-- who changed what. An edit is never silent.
CREATE TABLE IF NOT EXISTS audit (
  id        SERIAL PRIMARY KEY,
  nc_id     INT,
  field     TEXT,
  old_value TEXT,
  new_value TEXT,
  source    TEXT DEFAULT 'edit',      -- 'edit' | 'import:<kind>' | 'seed'
  at        TIMESTAMPTZ DEFAULT now()
);


-- CAPA — the Capa Board. Standalone; no foreign key to nc, on purpose.
CREATE TABLE IF NOT EXISTS capa (
  id                SERIAL PRIMARY KEY,
  requestor         TEXT DEFAULT '',
  responsible       TEXT DEFAULT '',
  dept_responsible  TEXT DEFAULT '',
  creation_date     TEXT DEFAULT '',
  origin            TEXT DEFAULT '',
  nc_number         TEXT DEFAULT '',
  nc_type           TEXT DEFAULT '',
  psp_element       TEXT DEFAULT '',
  project           TEXT DEFAULT '',
  project_manager   TEXT DEFAULT '',
  supplier          TEXT DEFAULT '',
  capa_type         TEXT DEFAULT '',
  id_number         TEXT DEFAULT '',
  origin_l1         TEXT DEFAULT '',
  origin_l2         TEXT DEFAULT '',
  rc_l1             TEXT DEFAULT '',
  rc_l2             TEXT DEFAULT '',
  problem           TEXT DEFAULT '',
  classification    TEXT DEFAULT '',
  mait_flow         TEXT DEFAULT '',
  drb_planned       TEXT DEFAULT '',
  priority_sum      TEXT DEFAULT '',
  priority          TEXT DEFAULT '',
  open_date         TEXT DEFAULT '',
  due_date          TEXT DEFAULT '',
  capa_id           TEXT DEFAULT '',
  status            TEXT DEFAULT '',
  close_date        TEXT DEFAULT '',
  days_open         TEXT DEFAULT '',
  comments          TEXT DEFAULT '',
  implemented       TEXT DEFAULT '',
  dept_accountable  TEXT DEFAULT '',
  dept_assigned     TEXT DEFAULT '',
  change_needed     TEXT DEFAULT '',
  training          TEXT DEFAULT '',
  updated_at        TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS capa_status_idx ON capa (status);
CREATE INDEX IF NOT EXISTS capa_id_idx     ON capa (capa_id);


-- ------------------------------------------------------------------
-- Additive only. Safe to re-run; safe with the previous image running.
-- ------------------------------------------------------------------
ALTER TABLE nc ADD COLUMN IF NOT EXISTS root_cause       TEXT DEFAULT '';
ALTER TABLE nc ADD COLUMN IF NOT EXISTS responsible_area TEXT DEFAULT '';
ALTER TABLE nc ADD COLUMN IF NOT EXISTS problem_description TEXT DEFAULT '';
ALTER TABLE nc ADD COLUMN IF NOT EXISTS notes               TEXT DEFAULT '';

ALTER TABLE nc ADD COLUMN IF NOT EXISTS origin_l1 TEXT DEFAULT '';
ALTER TABLE nc ADD COLUMN IF NOT EXISTS origin_l2 TEXT DEFAULT '';
ALTER TABLE nc ADD COLUMN IF NOT EXISTS rc_l2     TEXT DEFAULT '';