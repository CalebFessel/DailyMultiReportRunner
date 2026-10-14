-- TraumaSoft -> Samsara work order mapping ("endpoint table"), MySQL / MariaDB.
-- One row per TraumaSoft work order (or vehicle out-of-service episode) pushed to Samsara.
CREATE TABLE IF NOT EXISTS samsara_workorder_map (
    ts_wo_id          VARCHAR(64)  NOT NULL,
    samsara_wo_id     VARCHAR(64)  NOT NULL,
    samsara_asset_id  VARCHAR(64)  NULL,
    ts_vehicle_id     INT          NULL,
    vehicle_name      VARCHAR(100) NULL,
    samsara_status    VARCHAR(32)  NOT NULL,
    fingerprint       CHAR(64)     NULL,
    last_action       VARCHAR(16)  NULL,
    first_synced_at   DATETIME     NOT NULL,
    last_synced_at    DATETIME     NOT NULL,
    PRIMARY KEY (ts_wo_id),
    UNIQUE KEY ux_samsara_wo_id (samsara_wo_id),
    KEY ix_status (samsara_status)
);
