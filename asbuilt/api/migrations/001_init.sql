-- 001 · initial schema
--
-- Two ideas hold this together:
--   1. A checkpoint is a ROW in `station`, not a table and not code. Adding a
--      checkpoint is an INSERT; removing one is `active=false`. No migration.
--   2. Every station's rows live in ONE `entry` table, keyed by station_id.
--      A table per station would mean a migration per checkpoint — the thing
--      we are deliberately avoiding.

CREATE TABLE IF NOT EXISTS schema_migrations (
  version    TEXT PRIMARY KEY,
  applied_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE IF NOT EXISTS unit (
  id           SERIAL PRIMARY KEY,
  equipment_no TEXT NOT NULL,
  serial       TEXT DEFAULT '',
  order_no     TEXT DEFAULT '',
  sach_nr      TEXT DEFAULT '',
  product      TEXT DEFAULT '',
  created_at   TIMESTAMPTZ DEFAULT now(),
  UNIQUE (equipment_no)
);

-- a checkpoint on the chain
CREATE TABLE IF NOT EXISTS station (
  id        SERIAL PRIMARY KEY,
  key       TEXT NOT NULL UNIQUE,      -- stable id used by uploads and gates
  name      TEXT NOT NULL,
  source    TEXT DEFAULT '',           -- shown under the name: SAP, LABELS, …
  icon      TEXT DEFAULT '•',
  position  INT  NOT NULL,             -- order along the chain
  active    BOOLEAN DEFAULT TRUE
);

-- a comparison between two sides. Each side is 'station:<key>' or 'bom:<kind>'
CREATE TABLE IF NOT EXISTS gate (
  id           SERIAL PRIMARY KEY,
  key          TEXT NOT NULL UNIQUE,
  expected_ref TEXT NOT NULL,
  present_ref  TEXT NOT NULL,
  label_a      TEXT DEFAULT '',
  label_b      TEXT DEFAULT '',
  position     INT  NOT NULL,
  active       BOOLEAN DEFAULT TRUE
);

-- EBOM and MBOM. kind = 'ebom' | 'mbom'
CREATE TABLE IF NOT EXISTS bom (
  id          SERIAL PRIMARY KEY,
  unit_id     INT REFERENCES unit(id) ON DELETE CASCADE,
  kind        TEXT NOT NULL,
  material    TEXT NOT NULL,
  revision    TEXT DEFAULT '',
  description TEXT DEFAULT '',
  qty         NUMERIC DEFAULT 0,
  traceable   BOOLEAN DEFAULT FALSE,
  parent      TEXT DEFAULT '',
  position    INT DEFAULT 1,
  source_file TEXT DEFAULT '',
  imported_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS bom_lookup ON bom (unit_id, kind, material);

-- every station's rows, one table
CREATE TABLE IF NOT EXISTS entry (
  id          SERIAL PRIMARY KEY,
  unit_id     INT REFERENCES unit(id)    ON DELETE CASCADE,
  station_id  INT REFERENCES station(id) ON DELETE CASCADE,
  material    TEXT NOT NULL,
  revision    TEXT DEFAULT '',
  description TEXT DEFAULT '',
  batch       TEXT DEFAULT '',
  serial      TEXT DEFAULT '',
  qty         NUMERIC DEFAULT 0,
  work_order  TEXT DEFAULT '',
  position    INT DEFAULT 1,
  source_file TEXT DEFAULT '',
  imported_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS entry_lookup ON entry (unit_id, station_id, material);

-- every uploaded file, byte for byte, so any import can be replayed
CREATE TABLE IF NOT EXISTS upload (
  id          SERIAL PRIMARY KEY,
  unit_id     INT REFERENCES unit(id)    ON DELETE CASCADE,
  station_id  INT REFERENCES station(id) ON DELETE CASCADE,
  filename    TEXT NOT NULL,
  rows_loaded INT DEFAULT 0,
  content     BYTEA NOT NULL,
  uploaded_at TIMESTAMPTZ DEFAULT now()
);
CREATE INDEX IF NOT EXISTS upload_lookup ON upload (unit_id, station_id, uploaded_at DESC);