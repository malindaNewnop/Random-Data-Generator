# DataSyncer Test Data Generator

## WinForms Application Requirements

| Field | Value |
| --- | --- |
| Document Version | 1.0 |
| Date | April 2026 |
| Status | Draft for Review |
| Prepared For | DataSyncer QA Team |
| Source Reference | `DataSyncer_Test_Data_Requirements_.pdf` |

## 1. Purpose and Scope

This document specifies the requirements for a .NET WinForms desktop application that automates the creation of all test data defined in the DataSyncer Scenario Test Data Requirements.

The application eliminates manual data preparation effort by generating files, database rows, and folder structures ready for immediate QA execution.

The application must cover all six DataSyncer scenario types:

- `CSVtoDB` - generate deterministic CSV files per scenario
- `DBtoDB` - insert rows into source and destination database tables
- `DBtoJSON` - insert exportable rows into source tables
- `SQLQuery` - populate test tables for update-query scenarios
- `ProgramExecution` - verify script/executable paths and output folders
- `FileSyncer` - create local files and optionally stage remote files

## 2. Target Users

The primary users are QA engineers executing DataSyncer functional tests.

The application must be operable without command-line access or manual SQL scripting. All data generation must be triggered from the UI with confirmation feedback.

## 3. Application Architecture

### 3.1 Technology Stack

| Component | Technology | Notes |
| --- | --- | --- |
| UI Framework | .NET WinForms (.NET 6 or later) | Target OS: Windows 10 / Windows Server 2019+ |
| Deployment | Single `.exe` deployment preferred | Same environment as DataSyncer host |
| Database Access | `System.Data.SqlClient` or `Microsoft.Data.SqlClient` | SQL Server support required; optional MySQL |
| File I/O | `System.IO` | CSV, JSON, TXT generation |
| Configuration | JSON config file (`appsettings.json`) | Connection strings, paths, flags |
| Logging | Serilog or built-in `ILogger` | Append to a rolling log file |

### 3.2 Application Layout

The main window uses a left-side navigation panel (`TreeView` or `TabControl`) with one section per scenario type. A status panel at the bottom shows execution progress and results. A settings screen allows configuration of connection strings, output folders, and flag values.

| Panel / Tab | Purpose |
| --- | --- |
| Home / Dashboard | Summary of configured connection strings and output paths; one-click `Generate All` button |
| CSVtoDB | CSV file generator for Scenarios 1-4 |
| DBtoDB | DB source/destination row inserter for Scenarios 1-3 |
| DBtoJSON | Export-source row inserter for Scenarios 1-3 |
| SQLQuery | Test table populator for Scenarios 1-3 |
| ProgramExecution | Path validator and output folder creator for Scenarios 1-3 |
| FileSyncer | Local file creator and remote folder checker for Scenarios 1-3 |
| Settings | Connection strings, output folder paths, flag values, ID range overrides |
| Log Viewer | Scrollable in-app log of all generate/insert operations |

## 4. Settings and Configuration

### 4.1 Configuration File

The application reads and writes `appsettings.json` on startup. All settings are editable from the Settings screen without restarting the application.

| Setting Key | Type | Default / Example | Purpose |
| --- | --- | --- | --- |
| `SqlConnectionString` | `string` | `Server=localhost;Database=DSTest;...` | Primary SQL Server connection |
| `OutputRootFolder` | `string` | `D:\DataSyncerTest\GeneratedData` | Root folder for all generated files |
| `SyncFlagNotProcessed` | `string` | `N` | Value meaning "not yet processed" |
| `SyncFlagProcessed` | `string` | `Y` | Value meaning "already processed" |
| `CsvIdPrefix` | `string` | `CSV-` | Prefix for `CSVtoDB` `RecordId` values |
| `ApiTestEndpoint` | `string` | `http://localhost:5000/api/test` | URL for `DBtoJSON` API export tests |
| `RemoteFileSharePath` | `string` | `\\server\share\DSTest` | UNC path for `FileSyncer` remote folder |
| `DatetimeBase` | `string` | `2026-04-06 09:00:00` | Base datetime for all generated timestamps |

### 4.2 ID Range Rules

The application must respect the ID series defined in the source requirements to prevent cross-test contamination.

| Scenario Type | ID Series | Override Allowed |
| --- | --- | --- |
| `CSVtoDB` | `100000` series (`CSV-100001` ...) | Yes - via Settings |
| `DBtoDB` | `200000` series | Yes - via Settings |
| `DBtoJSON` | `300000` series | Yes - via Settings |
| `SQLQuery` | `400000` series | Yes - via Settings |
| `FileSyncer` | Prefix by scenario name (`upload_`, `download_`, `twoway_`) | Yes - via Settings |

## 5. CSVtoDB Screen Requirements

### 5.1 Shared CSV Controls

- Output folder path, pre-filled from Settings and editable
- Date/time base input, default `2026-04-06 09:00:00`
- Buttons: `Generate Selected Scenario`, `Generate All CSV Scenarios`, `Clear Output Folder`
- Status label showing last generated file name and row count

### 5.2 Scenario 1 - Happy Path CSV Import

| Control | Type | Default | Validation |
| --- | --- | --- | --- |
| Output filename | `TextBox` | `csv_happy_20260406_090000.csv` | Must end in `.csv`; no path separator |
| Row count | `NumericUpDown` | `20` | `5-500` |
| Key start | `TextBox` | `CSV-100001` | Non-empty |
| Include `Comment` column | `CheckBox` | Checked | - |
| `MachineCode` values | `TextBox` | `MC-01, MC-02, MC-03` | Comma-separated |
| `Status` values | `TextBox` | `NEW, RUN, DONE, WAIT` | Comma-separated |

Pre-state note shown in UI:

> Destination must not already contain these `RecordId` values. Monitored input folder must contain only this file.

### 5.3 Scenario 2 - Optional Column Missing

| Control | Type | Default | Validation |
| --- | --- | --- | --- |
| Output filename | `TextBox` | `csv_optional_missing_20260406_091000.csv` | Must end in `.csv` |
| Row count | `NumericUpDown` | `5` | `2-100` |
| Key start | `TextBox` | `CSV-100021` | Non-empty |
| Omit `Comment` column | `CheckBox` | Checked (omit) | - |

Pre-state note:

> Mapping must be configured so the missing CSV header is optional, not required.

### 5.4 Scenario 3 - Duplicate Prevention

| Control | Type | Default | Notes |
| --- | --- | --- | --- |
| Seed filename | `TextBox` | `csv_duplicate_seed_20260406_092000.csv` | First run file |
| Row count | `NumericUpDown` | `10` | Key range `101001-101010` |
| Key start | `TextBox` | `CSV-101001` | - |
| Generate variant file | `CheckBox` | Unchecked | `8` duplicate + `2` new keys |
| Variant filename | `TextBox` | `csv_duplicate_variant.csv` | Enabled only when above checked |

### 5.5 Scenario 4 - Scheduled Recurring Run

| File Generated | Timestamp Suffix | Key Range | Notes |
| --- | --- | --- | --- |
| `file_A_20260406_100000.csv` | `100000` | `CSV-102001 - CSV-102010` | - |
| `file_B_20260406_100100.csv` | `100100` | `CSV-102011 - CSV-102020` | - |
| `file_C_20260406_100300.csv` | `100300` | `CSV-102021 - CSV-102030` | - |
| `file_A_REDROP` copy | `100300` | Same keys as `file_A` | Identical content to `file_A` |

- `Generate All 4 Files` button creates all three distinct files plus the re-drop copy in one action
- Output folder shown after generation with option to open in Explorer

## 6. DBtoDB Screen Requirements

### 6.1 Connection Controls

- Connection string selector, as a dropdown of named connections from Settings
- Source table name input, default `dbo.DS_Source_DBtoDB`
- Destination table name input, default `dbo.DS_Dest_DBtoDB`
- `Test Connection` button with inline result label
- `Create Tables if Missing` button; DDL for both source and destination based on the required section 2.2 shape

### 6.2 Scenario 1 - Immediate Flag-Based Copy

| Control | Type | Default |
| --- | --- | --- |
| Row count | `NumericUpDown` | `10` |
| `SourceId` start | `NumericUpDown` | `200001` |
| `MachineCode` pattern | `TextBox` | `MC-01, MC-02, MC-03, MC-04, MC-05` |
| `WorkCenter` | `TextBox` | `LINE-A` |
| `SyncFlag` value | `TextBox` | `N` from Settings |
| `EventTime` base | `TextBox` | `2026-04-06 11:00:00` |
| Clear destination before insert | `CheckBox` | Unchecked |

### 6.3 Scenario 2 - Duplicate-Key Recovery

- Row count: `5` (`SourceId 200011-200015`)
- `Pre-insert N rows to destination` option inserts `N` of the `5` keys with stale data to simulate a prior run
- Conflict behavior selector: `Skip`, `Overwrite`, `Raise Error`
- UI reminder label should tell the tester to verify job config matches the selected conflict behavior

### 6.4 Scenario 3 - Manual Copy by Date Range and Condition

- Row count: `15` total (`SourceId 200016-200030`)
- Date split controls:
  - Rows inside range: default `5`, `EventTime 2026-04-06`
  - Rows outside range: default `5`, `EventTime 2026-04-05`
  - Plus `5` mixed-condition rows
- WorkCenter filter values input: `LINE-A, LINE-B`
- `Pre-load destination with stale rows in selected range` checkbox for delete-and-reinsert validation

## 7. DBtoJSON Screen Requirements

### 7.1 Scenario 1 - File Export Happy Path

| Control | Default | Notes |
| --- | --- | --- |
| Source table | `dbo.DS_Source_DBtoJSON` | Editable |
| Row count | `5` | `ExportId 300001-300005` |
| `ExportFlag` value | `N` | From Settings |
| `DeviceCode` values | `DEV-01, DEV-02` | Comma-separated |
| `ResultCode` values | `READY, PASS` | Comma-separated |
| `CreatedAt` base | `2026-04-06 14:00:00` | Increments by 1 minute per row |
| Output folder (for pre-state) | `D:\DataSyncerTest\JsonOutput` | Shown read-only with `Open` button |

### 7.2 Scenario 2 - API Export with Response Mapping

- Row count: `3-5`, `ExportFlag = N`
- API field inputs:
  - `crtfcKey`
  - `useSe`
  - `sysUser`
  - `conectIp`
  - `dataUsgqty`
- Fields should be pre-filled with recommended values
- `Create response-writeback columns` button runs `ALTER TABLE` to add:
  - `ApiResultCode`
  - `ApiResultMessage`
  - `ApiReceivedAt`
- `Ping Endpoint` button verifies `DataSyncer.TestApiEndpoint` is reachable before insert

### 7.3 Scenario 3 - No New Data Re-run

- `Mark all rows as Processed` button updates `ExportFlag` to `Y` for all rows in the source table
- Display current row count with `ExportFlag = N` before and after the action

## 8. SQLQuery Screen Requirements

### 8.1 Scenario 1 - Scheduled Update Query

| Control | Default | Notes |
| --- | --- | --- |
| Test table name | `dbo.DS_SqlQuery_Test` | Editable |
| Total row count | `8` | `RowId 400001-400008` |
| `PENDING` row count | `5` | `Status = PENDING` |
| `DONE` row count | `3` | `Status = DONE` |
| `UpdatedByJob` | `(null before job)` | Confirmed null on insert |
| `UpdatedAt` | `(null before job)` | Confirmed null on insert |
| Schedule time label | `2026-04-06 16:00:00` | Informational only; not executed by this app |

### 8.2 Scenario 2 - Invalid SQL

No data generation is needed. The UI shows this reminder:

> Place intentionally malformed SQL in the job configuration. Optionally keep the Scenario 1 test table in place to confirm no valid change occurred.

### 8.3 Scenario 3 - Wrong DB Connection

No data generation is needed. The UI shows this reminder:

> Configure a broken connection string (bad host, bad credentials, or mismatched DB type). Scenario 1 test table may remain; no special data required.

## 9. ProgramExecution Screen Requirements

### 9.1 Scenario 1 - Successful Executable Run

- Script path input, default `TestResources/TwoDaySoak/program_execution_heartbeat.ps1`
- `OutputFile` argument input, default `D:\DataSyncerTest\program_execution_heartbeat.log`
- `MessagePrefix` argument input, default `ProgramExecution functional test`
- `Create Output Folder` button ensures the output folder exists
- `Validate Script Path` button checks the file exists and is `.ps1` or `.exe`
- `Archive Prior Output File` button renames an existing `.log` with a timestamp suffix

### 9.2 Scenario 2 - Invalid Executable Path

- Invalid path input, pre-filled with `D:\DoesNotExist\missing_tool.exe`
- `Confirm Path Does Not Exist` button verifies the path is absent so the negative test is valid

### 9.3 Scenario 3 - Executable with Arguments

- Script path: `program_execution_heartbeat.ps1` (same as Scenario 1)
- `OutputFile`: `D:\DataSyncerTest\arg_check.log`
- `MessagePrefix`: `Arg Test Value` (contains a space and confirms quoting is correct)
- `Preview Full Command` button shows the complete command line before execution

## 10. FileSyncer Screen Requirements

### 10.1 Scenario 1 - Client-to-Server Upload

| File to Generate | Content | Size Target |
| --- | --- | --- |
| `upload_20260406_170000.txt` | Short text content with timestamp | `< 1 KB` |
| `upload_20260406_170001.csv` | `5`-row CSV with header (`RecordId, Value, Timestamp`) | `< 5 KB` |
| `upload_20260406_170002.json` | `3`-object JSON array `[{id, value, ts}]` | `< 2 KB` |

- Local source folder input is editable and defaults to `OutputRootFolder\FileSyncer\Upload`
- `Generate Files` button creates all 3 files in the local folder
- Remote target folder path is shown read-only from Settings, with a reachability check

### 10.2 Scenario 2 - Server-to-Client Download

- Files to generate on remote side:
  - `download_20260406_171000.txt`
  - `download_20260406_171001.csv`
  - `download_20260406_171002.json`
- `Stage Remote Files` button copies generated files to the configured remote path and requires write access
- Local target folder input includes `Verify Empty / Clear`

### 10.3 Scenario 3 - Two-Way Sync Conflict Check

| File | Location | Content to Generate |
| --- | --- | --- |
| `twoway_local_only.txt` | Local only | `LOCAL ONLY file - 2026-04-06 17:20:00` |
| `twoway_remote_only.txt` | Remote only | `REMOTE ONLY file - 2026-04-06 17:20:00` |
| `twoway_shared_conflict.txt` (local) | Local | `LOCAL VERSION 2026-04-06 17:20:00` |
| `twoway_shared_conflict.txt` (remote) | Remote | `REMOTE VERSION 2026-04-06 17:21:00` |

- `Generate All Conflict Files` button creates all 4 files in their correct locations
- Checksum and size display for both local and remote conflict files for pre-state recording

## 11. Dashboard / Home Screen Requirements

- Connection string health shows green/red indicator per named connection
- Output folder status shows whether the root folder exists and is writable
- `Generate All Scenarios` button runs all enabled scenario generators in sequence with a progress bar
- Per-scenario toggle checkboxes include or exclude screens from `Generate All`
- Last run summary table shows scenario, rows/files created, timestamp, and status

## 12. Log Viewer Requirements

- Scrollable `RichTextBox` showing timestamped log entries
- Colour coding:
  - Green = success
  - Red = error
  - Yellow = warning
  - Grey = info
- `Clear Log` and `Save Log to File` buttons
- Log persisted to `DataSyncerTestGen_{date}.log` in the output folder

## 13. Error Handling Requirements

| Error Condition | Required Behaviour |
| --- | --- |
| DB connection failure | Show `MessageBox` with connection string, masked password, and inner exception; do not crash |
| Output folder not writable | Prompt user to create the folder or choose an alternative; do not silently fail |
| Row already exists (duplicate key) | Log warning with `SourceId` or `RecordId`; continue with remaining rows unless `Stop on Error` is checked |
| Script/exe path not found | Show inline validation error in red next to the path field |
| Remote path not reachable | Show reachability status with last-checked timestamp; allow the user to retry |
| Partial batch failure | Report rows succeeded / rows failed in the status label and write detail to the log |

## 14. Non-Functional Requirements

| Requirement | Specification |
| --- | --- |
| Performance | Full `Generate All` with all scenarios must complete within 60 seconds on local SQL Server |
| Deployment | Single EXE with embedded dependencies via .NET single-file publish; no installer required |
| Configuration | `appsettings.json` co-located with EXE and editable via Notepad as fallback |
| Logging | Rolling daily log file; maximum `10 MB` per file; `7`-day retention |
| Accessibility | All controls must have accessible names and tooltip text |
| Screen size | Minimum supported resolution: `1280 x 768` |
| Data safety | Never connects to production DB; Settings screen shows warning if the connection string contains `prod` or `prd` |

## 15. Out of Scope

- Actual execution of DataSyncer jobs; this app only prepares data
- Test result verification or assertions
- CI/CD pipeline integration; command-line interface is a future enhancement
- Support for database engines other than SQL Server in `v1.0`
- Automated cleanup of generated data after test runs

## 16. Open Questions / Decisions Required

| # | Question | Owner | Status |
| --- | --- | --- | --- |
| 1 | Should the app support MySQL in addition to SQL Server for `DBtoDB` tests? | QA Lead | Open |
| 2 | Is the remote `FileSyncer` path always a Windows UNC share, or also SFTP? | Infra Team | Open |
| 3 | Should `Generate All` be runnable from a command-line argument for CI use? | DevOps | Open |
| 4 | Which .NET version is already installed on QA machines (`6`, `8`, or `9`)? | QA Lead | Open |
| 5 | Should generated data be automatically rolled back after each test run? | QA Lead | Open |

## Appendix A - Recommended Project Structure

Suggested Visual Studio solution layout:

- `DataSyncerTestGen.sln`
- `DataSyncerTestGen/`
- `Program.cs` - entry point
- `Forms/MainForm.cs` - shell with nav panel
- `Forms/SettingsForm.cs`
- `Forms/CsvToDbPanel.cs`
- `Forms/DbToDbPanel.cs`
- `Forms/DbToJsonPanel.cs`
- `Forms/SqlQueryPanel.cs`
- `Forms/ProgramExecPanel.cs`
- `Forms/FileSyncerPanel.cs`
- `Forms/LogViewerPanel.cs`
- `Services/CsvGeneratorService.cs`
- `Services/DbInsertService.cs`
- `Services/FileGeneratorService.cs`
- `Services/ConnectionTestService.cs`
- `Models/AppSettings.cs`
- `Models/ScenarioResult.cs`
- `appsettings.json`

Each service class should be independently unit-testable and injected into the form panels via constructor.
