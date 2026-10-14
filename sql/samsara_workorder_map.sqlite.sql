-- SQLite flavour (used by the unit tests and for local experiments).
CREATE TABLE IF NOT EXISTS samsara_workorder_map (
    ts_wo_id          INTEGER PRIMARY KEY,
    samsara_wo_id     TEXT NOT NULL UNIQUE,
    samsara_asset_id  TEXT,
    ts_vehicle_id     INTEGER,
    vehicle_name      TEXT,
    samsara_status    TEXT NOT NULL,
    fingerprint       TEXT,
    last_action       TEXT,
    first_synced_at   TEXT NOT NULL,
    last_synced_at    TEXT NOT NULL
);
