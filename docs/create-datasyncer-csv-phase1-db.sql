IF DB_ID(N'DataSyncer_Test') IS NULL
BEGIN
    CREATE DATABASE [DataSyncer_Test];
END;
GO

USE [DataSyncer_Test];
GO

IF OBJECT_ID(N'dbo.csv_scenario_happy_path', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.csv_scenario_happy_path
    (
        id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        serial NVARCHAR(100) NOT NULL,
        work_date DATE NOT NULL,
        measured_at DATETIME2(0) NOT NULL,
        machine_code NVARCHAR(50) NOT NULL,
        ok_ng NVARCHAR(10) NOT NULL,
        measure1_p1 FLOAT NULL,
        measure2_p2 FLOAT NULL,
        measure3_p3 FLOAT NULL,
        measure4_p4 FLOAT NULL,
        beta1_p5 FLOAT NULL,
        beta2_p6 FLOAT NULL,
        beta3_p7 FLOAT NULL,
        beta4_p8 FLOAT NULL,
        beta5_p9 FLOAT NULL,
        beta6_p10 FLOAT NULL,
        beta_range_p5_p10 FLOAT NULL,
        beta_average_p5_p10 FLOAT NULL,
        comment NVARCHAR(400) NULL,
        inserted_at DATETIME2(0) NOT NULL CONSTRAINT DF_csv_scenario_happy_path_inserted_at DEFAULT SYSDATETIME()
    );
END;
GO

IF OBJECT_ID(N'dbo.csv_scenario_optional_missing', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.csv_scenario_optional_missing
    (
        id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        serial NVARCHAR(100) NOT NULL,
        work_date DATE NOT NULL,
        measured_at DATETIME2(0) NOT NULL,
        machine_code NVARCHAR(50) NOT NULL,
        ok_ng NVARCHAR(10) NOT NULL,
        measure1_p1 FLOAT NULL,
        measure2_p2 FLOAT NULL,
        measure3_p3 FLOAT NULL,
        measure4_p4 FLOAT NULL,
        beta1_p5 FLOAT NULL,
        beta2_p6 FLOAT NULL,
        beta3_p7 FLOAT NULL,
        beta4_p8 FLOAT NULL,
        beta5_p9 FLOAT NULL,
        beta6_p10 FLOAT NULL,
        beta_range_p5_p10 FLOAT NULL,
        beta_average_p5_p10 FLOAT NULL,
        comment NVARCHAR(400) NULL,
        inserted_at DATETIME2(0) NOT NULL CONSTRAINT DF_csv_scenario_optional_missing_inserted_at DEFAULT SYSDATETIME()
    );
END;
GO

IF OBJECT_ID(N'dbo.csv_scenario_duplicate_prevention', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.csv_scenario_duplicate_prevention
    (
        id BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        serial NVARCHAR(100) NOT NULL,
        work_date DATE NOT NULL,
        measured_at DATETIME2(0) NOT NULL,
        machine_code NVARCHAR(50) NOT NULL,
        ok_ng NVARCHAR(10) NOT NULL,
        measure1_p1 FLOAT NULL,
        measure2_p2 FLOAT NULL,
        measure3_p3 FLOAT NULL,
        measure4_p4 FLOAT NULL,
        beta1_p5 FLOAT NULL,
        beta2_p6 FLOAT NULL,
        beta3_p7 FLOAT NULL,
        beta4_p8 FLOAT NULL,
        beta5_p9 FLOAT NULL,
        beta6_p10 FLOAT NULL,
        beta_range_p5_p10 FLOAT NULL,
        beta_average_p5_p10 FLOAT NULL,
        comment NVARCHAR(400) NULL,
        inserted_at DATETIME2(0) NOT NULL CONSTRAINT DF_csv_scenario_duplicate_prevention_inserted_at DEFAULT SYSDATETIME()
    );
END;
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_csv_scenario_happy_path_serial'
      AND object_id = OBJECT_ID(N'dbo.csv_scenario_happy_path')
)
BEGIN
    CREATE UNIQUE INDEX UX_csv_scenario_happy_path_serial
        ON dbo.csv_scenario_happy_path(serial);
END;
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_csv_scenario_optional_missing_serial'
      AND object_id = OBJECT_ID(N'dbo.csv_scenario_optional_missing')
)
BEGIN
    CREATE UNIQUE INDEX UX_csv_scenario_optional_missing_serial
        ON dbo.csv_scenario_optional_missing(serial);
END;
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_csv_scenario_duplicate_prevention_serial'
      AND object_id = OBJECT_ID(N'dbo.csv_scenario_duplicate_prevention')
)
BEGIN
    CREATE UNIQUE INDEX UX_csv_scenario_duplicate_prevention_serial
        ON dbo.csv_scenario_duplicate_prevention(serial);
END;
GO
