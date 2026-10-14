-- TraumaSoft -> Samsara work order mapping ("endpoint table"), SQL Server.
-- One row per TraumaSoft work order that has been pushed to Samsara.
IF OBJECT_ID('dbo.samsara_workorder_map', 'U') IS NULL
CREATE TABLE dbo.samsara_workorder_map (
    ts_wo_id          VARCHAR(64)   NOT NULL PRIMARY KEY,
    samsara_wo_id     VARCHAR(64)   NOT NULL,
    samsara_asset_id  VARCHAR(64)   NULL,
    ts_vehicle_id     INT           NULL,
    vehicle_name      NVARCHAR(100) NULL,
    samsara_status    VARCHAR(32)   NOT NULL,
    fingerprint       CHAR(64)      NULL,
    last_action       VARCHAR(16)   NULL,
    first_synced_at   DATETIME2(0)  NOT NULL,
    last_synced_at    DATETIME2(0)  NOT NULL,
    CONSTRAINT ux_samsara_workorder_map_samsara_wo_id UNIQUE (samsara_wo_id)
);
