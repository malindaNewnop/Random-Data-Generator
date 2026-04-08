using System.Net.Http;
using System.Text;
using System.Globalization;
using Microsoft.Data.SqlClient;

namespace DataSyncer.TestDataGenerator.App;

public partial class Form1
{
    private ComboBox? _cmbDbToDbConnection;
    private TextBox? _txtDbToDbSourceTable;
    private TextBox? _txtDbToDbDestinationTable;
    private Label? _lblDbToDbConnectionStatus;
    private Label? _lblDbToDbStatus;
    private NumericUpDown? _numDbToDbS1RowCount;
    private NumericUpDown? _numDbToDbS1SourceIdStart;
    private TextBox? _txtDbToDbS1MachineCodes;
    private TextBox? _txtDbToDbS1WorkCenter;
    private TextBox? _txtDbToDbS1SyncFlag;
    private DateTimePicker? _dtDbToDbS1EventTimeBase;
    private CheckBox? _chkDbToDbS1ClearDestination;
    private NumericUpDown? _numDbToDbS2PreInsertCount;
    private ComboBox? _cmbDbToDbS2ConflictBehavior;
    private Label? _lblDbToDbS2ConflictNote;
    private NumericUpDown? _numDbToDbS3InsideRangeCount;
    private NumericUpDown? _numDbToDbS3OutsideRangeCount;
    private DateTimePicker? _dtDbToDbS3InsideDate;
    private DateTimePicker? _dtDbToDbS3OutsideDate;
    private TextBox? _txtDbToDbS3WorkCenters;
    private CheckBox? _chkDbToDbS3PreloadDestination;

    private ComboBox? _cmbDbToJsonConnection;
    private TextBox? _txtDbToJsonSourceTable;
    private Label? _lblDbToJsonConnectionStatus;
    private Label? _lblDbToJsonStatus;
    private NumericUpDown? _numDbToJsonS1RowCount;
    private TextBox? _txtDbToJsonS1ExportFlag;
    private TextBox? _txtDbToJsonS1DeviceCodes;
    private TextBox? _txtDbToJsonS1ResultCodes;
    private DateTimePicker? _dtDbToJsonS1CreatedAtBase;
    private TextBox? _txtDbToJsonOutputFolder;
    private NumericUpDown? _numDbToJsonS2RowCount;
    private TextBox? _txtDbToJsonS2CrtfcKey;
    private TextBox? _txtDbToJsonS2UseSe;
    private TextBox? _txtDbToJsonS2SysUser;
    private TextBox? _txtDbToJsonS2ConectIp;
    private TextBox? _txtDbToJsonS2DataUsgqty;
    private Label? _lblDbToJsonPingStatus;
    private Label? _lblDbToJsonReRunCounts;

    private ComboBox? _cmbSqlQueryConnection;
    private TextBox? _txtSqlQueryTableName;
    private Label? _lblSqlQueryConnectionStatus;
    private Label? _lblSqlQueryStatus;
    private NumericUpDown? _numSqlQueryTotalRows;
    private NumericUpDown? _numSqlQueryPendingRows;
    private NumericUpDown? _numSqlQueryDoneRows;

    private TextBox? _txtProgramS1ScriptPath;
    private TextBox? _txtProgramS1OutputFile;
    private TextBox? _txtProgramS1MessagePrefix;
    private Label? _lblProgramS1Status;
    private TextBox? _txtProgramS2InvalidPath;
    private Label? _lblProgramS2Status;
    private TextBox? _txtProgramS3ScriptPath;
    private TextBox? _txtProgramS3OutputFile;
    private TextBox? _txtProgramS3MessagePrefix;
    private TextBox? _txtProgramS3CommandPreview;

    private Control CreateDbToDbPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateDbToDbConnectionCard());
        stack.Controls.Add(CreateDbToDbScenario1Card());
        stack.Controls.Add(CreateDbToDbScenario2Card());
        stack.Controls.Add(CreateDbToDbScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateDbToJsonPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateDbToJsonConnectionCard());
        stack.Controls.Add(CreateDbToJsonScenario1Card());
        stack.Controls.Add(CreateDbToJsonScenario2Card());
        stack.Controls.Add(CreateDbToJsonScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateSqlQueryPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateSqlQueryScenario1Card());
        stack.Controls.Add(CreateSqlQueryReminderCard(
            "Scenario 2 - Invalid SQL",
            "Place intentionally malformed SQL in the DataSyncer job configuration. Optionally keep the Scenario 1 test table in place to confirm that no valid change occurred.",
            WarningSoft,
            WarningColor));
        stack.Controls.Add(CreateSqlQueryReminderCard(
            "Scenario 3 - Wrong DB Connection",
            "Configure a broken connection string with a bad host, credentials, or mismatched DB type. The Scenario 1 test table may remain; no special data is required.",
            DangerSoft,
            DangerColor));

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateProgramExecutionPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateProgramScenario1Card());
        stack.Controls.Add(CreateProgramScenario2Card());
        stack.Controls.Add(CreateProgramScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateDbToDbConnectionCard()
    {
        var card = CreateCard("DBtoDB", "Prepare source and destination tables for immediate-copy, duplicate-recovery, and date-range scenarios without manual SQL scripting.", out var content);
        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Connection", 0);
        _cmbDbToDbConnection = CreateConnectionSelector();
        grid.Controls.Add(_cmbDbToDbConnection, 1, 0);
        _lblDbToDbConnectionStatus = CreateInlineStatusLabel("Connection not tested yet.");
        grid.Controls.Add(_lblDbToDbConnectionStatus, 2, 0);

        AddLabelCell(grid, "Source table", 1);
        _txtDbToDbSourceTable = CreateCsvTextBox("dbo.DS_Source_DBtoDB", "Source table for DBtoDB test rows.");
        grid.Controls.Add(_txtDbToDbSourceTable, 1, 1);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 1);

        AddLabelCell(grid, "Destination table", 2);
        _txtDbToDbDestinationTable = CreateCsvTextBox("dbo.DS_Dest_DBtoDB", "Destination table used for stale-row and duplicate-key setup.");
        grid.Controls.Add(_txtDbToDbDestinationTable, 1, 2);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 2);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Test Connection", (_, _) => TestDbToDbConnection(), true));
        actions.Controls.Add(CreateActionButton("Create Tables if Missing", (_, _) => CreateDbToDbTablesIfMissing()));
        content.Controls.Add(actions);

        _lblDbToDbStatus = new Label
        {
            Text = "No DBtoDB action has run yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblDbToDbStatus);

        return card;
    }

    private Control CreateDbToDbScenario1Card()
    {
        var card = CreateCard("Scenario 1 - Immediate Flag-Based Copy", "Insert fresh source rows with the not-processed flag and optionally clear the destination table before the test run.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Row count", 0);
        _numDbToDbS1RowCount = CreateCsvNumeric(1, 500, 10, "Number of source rows to insert.");
        grid.Controls.Add(_numDbToDbS1RowCount, 1, 0);

        AddLabelCell(grid, "SourceId start", 1);
        _numDbToDbS1SourceIdStart = CreateCsvNumeric(200000, 299999, 200001, "Starting SourceId for Scenario 1.");
        grid.Controls.Add(_numDbToDbS1SourceIdStart, 1, 1);

        AddLabelCell(grid, "MachineCode pattern", 2);
        _txtDbToDbS1MachineCodes = CreateCsvTextBox("MC-01, MC-02, MC-03, MC-04, MC-05", "Comma-separated MachineCode values.");
        grid.Controls.Add(_txtDbToDbS1MachineCodes, 1, 2);

        AddLabelCell(grid, "WorkCenter", 3);
        _txtDbToDbS1WorkCenter = CreateCsvTextBox("LINE-A", "WorkCenter value applied to generated rows.");
        grid.Controls.Add(_txtDbToDbS1WorkCenter, 1, 3);

        AddLabelCell(grid, "SyncFlag value", 4);
        _txtDbToDbS1SyncFlag = CreateCsvTextBox(_settings.SyncFlagNotProcessed, "Sync flag value used for fresh rows.");
        grid.Controls.Add(_txtDbToDbS1SyncFlag, 1, 4);

        AddLabelCell(grid, "EventTime base", 5);
        _dtDbToDbS1EventTimeBase = CreateDateTimeInput(new DateTime(2026, 4, 6, 11, 0, 0), "Base EventTime for Scenario 1 rows.");
        grid.Controls.Add(_dtDbToDbS1EventTimeBase, 1, 5);

        AddLabelCell(grid, "Clear destination before insert", 6);
        _chkDbToDbS1ClearDestination = new CheckBox
        {
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 12)
        };
        _toolTip.SetToolTip(_chkDbToDbS1ClearDestination, "Delete all destination rows before seeding Scenario 1.");
        grid.Controls.Add(_chkDbToDbS1ClearDestination, 1, 6);

        content.Controls.Add(grid);
        content.Controls.Add(CreateActionButton("Generate Scenario 1 Data", (_, _) => GenerateDbToDbScenario1(), true));
        return card;
    }

    private Control CreateDbToDbScenario2Card()
    {
        var card = CreateCard("Scenario 2 - Duplicate-Key Recovery", "Insert five source rows and optionally pre-insert stale destination rows to simulate a previous run.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "SourceId range", 0);
        var sourceRangeLabel = new Label
        {
            Text = "200011 - 200015",
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 12)
        };
        grid.Controls.Add(sourceRangeLabel, 1, 0);

        AddLabelCell(grid, "Pre-insert N rows to destination", 1);
        _numDbToDbS2PreInsertCount = CreateCsvNumeric(0, 5, 2, "Number of stale destination rows to insert before the test.");
        grid.Controls.Add(_numDbToDbS2PreInsertCount, 1, 1);

        AddLabelCell(grid, "Conflict behavior", 2);
        _cmbDbToDbS2ConflictBehavior = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 220,
            Margin = new Padding(0, 0, 12, 12)
        };
        _cmbDbToDbS2ConflictBehavior.Items.AddRange(new object[] { "Skip", "Overwrite", "Raise Error" });
        _cmbDbToDbS2ConflictBehavior.SelectedIndex = 0;
        _cmbDbToDbS2ConflictBehavior.SelectedIndexChanged += (_, _) => UpdateDbToDbConflictNote();
        grid.Controls.Add(_cmbDbToDbS2ConflictBehavior, 1, 2);

        _lblDbToDbS2ConflictNote = CreateInlineStatusLabel("Verify the job configuration matches the selected conflict behavior: Skip.");
        grid.Controls.Add(_lblDbToDbS2ConflictNote, 1, 3);

        content.Controls.Add(grid);
        content.Controls.Add(CreateActionButton("Generate Scenario 2 Data", (_, _) => GenerateDbToDbScenario2(), true));
        return card;
    }

    private Control CreateDbToDbScenario3Card()
    {
        var card = CreateCard("Scenario 3 - Manual Copy by Date Range and Condition", "Create inside-range, outside-range, and mixed-condition rows to validate date filtering and delete-and-reinsert behavior.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Rows inside range", 0);
        _numDbToDbS3InsideRangeCount = CreateCsvNumeric(1, 20, 5, "Number of rows inside the selected date range.");
        grid.Controls.Add(_numDbToDbS3InsideRangeCount, 1, 0);

        AddLabelCell(grid, "Rows outside range", 1);
        _numDbToDbS3OutsideRangeCount = CreateCsvNumeric(1, 20, 5, "Number of rows outside the selected date range.");
        grid.Controls.Add(_numDbToDbS3OutsideRangeCount, 1, 1);

        AddLabelCell(grid, "Inside-range EventTime", 2);
        _dtDbToDbS3InsideDate = CreateDateTimeInput(new DateTime(2026, 4, 6, 9, 0, 0), "EventTime used for inside-range rows.");
        grid.Controls.Add(_dtDbToDbS3InsideDate, 1, 2);

        AddLabelCell(grid, "Outside-range EventTime", 3);
        _dtDbToDbS3OutsideDate = CreateDateTimeInput(new DateTime(2026, 4, 5, 9, 0, 0), "EventTime used for outside-range rows.");
        grid.Controls.Add(_dtDbToDbS3OutsideDate, 1, 3);

        AddLabelCell(grid, "WorkCenter filter values", 4);
        _txtDbToDbS3WorkCenters = CreateCsvTextBox("LINE-A, LINE-B", "Comma-separated WorkCenter values used to build mixed-condition rows.");
        grid.Controls.Add(_txtDbToDbS3WorkCenters, 1, 4);

        AddLabelCell(grid, "Pre-load destination with stale rows", 5);
        _chkDbToDbS3PreloadDestination = new CheckBox
        {
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 12)
        };
        _toolTip.SetToolTip(_chkDbToDbS3PreloadDestination, "Insert stale destination rows in the selected range for delete-and-reinsert validation.");
        grid.Controls.Add(_chkDbToDbS3PreloadDestination, 1, 5);

        content.Controls.Add(grid);
        content.Controls.Add(CreateInfoNote("Scenario layout", "This generator always adds 5 mixed-condition rows in addition to the inside-range and outside-range counts.", AccentSoft, AccentColor));
        content.Controls.Add(CreateActionButton("Generate Scenario 3 Data", (_, _) => GenerateDbToDbScenario3(), true));
        return card;
    }

    private Control CreateDbToJsonConnectionCard()
    {
        var card = CreateCard("DBtoJSON", "Prepare export-source rows for file export, API export, and no-new-data rerun scenarios.", out var content);
        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Connection", 0);
        _cmbDbToJsonConnection = CreateConnectionSelector();
        grid.Controls.Add(_cmbDbToJsonConnection, 1, 0);
        _lblDbToJsonConnectionStatus = CreateInlineStatusLabel("Connection not tested yet.");
        grid.Controls.Add(_lblDbToJsonConnectionStatus, 2, 0);

        AddLabelCell(grid, "Source table", 1);
        _txtDbToJsonSourceTable = CreateCsvTextBox("dbo.DS_Source_DBtoJSON", "Source table for DBtoJSON export rows.");
        grid.Controls.Add(_txtDbToJsonSourceTable, 1, 1);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 1);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Test Connection", (_, _) => TestDbToJsonConnection(), true));
        actions.Controls.Add(CreateActionButton("Create Source Table if Missing", (_, _) => CreateDbToJsonTableIfMissing()));
        content.Controls.Add(actions);

        _lblDbToJsonStatus = new Label
        {
            Text = "No DBtoJSON action has run yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblDbToJsonStatus);

        return card;
    }

    private Control CreateDbToJsonScenario1Card()
    {
        var card = CreateCard("Scenario 1 - File Export Happy Path", "Insert exportable rows flagged as new and stage the expected output folder for pre-state verification.", out var content);
        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Row count", 0);
        _numDbToJsonS1RowCount = CreateCsvNumeric(1, 100, 5, "Number of export rows to insert.");
        grid.Controls.Add(_numDbToJsonS1RowCount, 1, 0);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 0);

        AddLabelCell(grid, "ExportFlag value", 1);
        _txtDbToJsonS1ExportFlag = CreateCsvTextBox(_settings.SyncFlagNotProcessed, "Export flag value for new rows.");
        grid.Controls.Add(_txtDbToJsonS1ExportFlag, 1, 1);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 1);

        AddLabelCell(grid, "DeviceCode values", 2);
        _txtDbToJsonS1DeviceCodes = CreateCsvTextBox("DEV-01, DEV-02", "Comma-separated DeviceCode values.");
        grid.Controls.Add(_txtDbToJsonS1DeviceCodes, 1, 2);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 2);

        AddLabelCell(grid, "ResultCode values", 3);
        _txtDbToJsonS1ResultCodes = CreateCsvTextBox("READY, PASS", "Comma-separated ResultCode values.");
        grid.Controls.Add(_txtDbToJsonS1ResultCodes, 1, 3);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 3);

        AddLabelCell(grid, "CreatedAt base", 4);
        _dtDbToJsonS1CreatedAtBase = CreateDateTimeInput(new DateTime(2026, 4, 6, 14, 0, 0), "Base CreatedAt timestamp for DBtoJSON rows.");
        grid.Controls.Add(_dtDbToJsonS1CreatedAtBase, 1, 4);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 4);

        AddLabelCell(grid, "Output folder (pre-state)", 5);
        _txtDbToJsonOutputFolder = CreateCsvTextBox(@"D:\DataSyncerTest\JsonOutput", "Expected JSON output folder used for pre-state checks.");
        _txtDbToJsonOutputFolder.ReadOnly = true;
        grid.Controls.Add(_txtDbToJsonOutputFolder, 1, 5);
        grid.Controls.Add(CreateMiniButton("Open", (_, _) => OpenDbToJsonOutputFolder()), 2, 5);

        content.Controls.Add(grid);
        content.Controls.Add(CreateActionButton("Generate Scenario 1 Data", (_, _) => GenerateDbToJsonScenario1(), true));
        return card;
    }

    private Control CreateDbToJsonScenario2Card()
    {
        var card = CreateCard("Scenario 2 - API Export with Response Mapping", "Insert API export rows, verify the endpoint is reachable, and add response-writeback columns when needed.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Row count", 0);
        _numDbToJsonS2RowCount = CreateCsvNumeric(3, 5, 3, "Number of API export rows to insert.");
        grid.Controls.Add(_numDbToJsonS2RowCount, 1, 0);

        AddLabelCell(grid, "crtfcKey", 1);
        _txtDbToJsonS2CrtfcKey = CreateCsvTextBox("QA-CRTFC-KEY", "Recommended crtfcKey value.");
        grid.Controls.Add(_txtDbToJsonS2CrtfcKey, 1, 1);

        AddLabelCell(grid, "useSe", 2);
        _txtDbToJsonS2UseSe = CreateCsvTextBox("API", "Recommended useSe value.");
        grid.Controls.Add(_txtDbToJsonS2UseSe, 1, 2);

        AddLabelCell(grid, "sysUser", 3);
        _txtDbToJsonS2SysUser = CreateCsvTextBox("datasyncer.qa", "Recommended sysUser value.");
        grid.Controls.Add(_txtDbToJsonS2SysUser, 1, 3);

        AddLabelCell(grid, "conectIp", 4);
        _txtDbToJsonS2ConectIp = CreateCsvTextBox("127.0.0.1", "Recommended conectIp value.");
        grid.Controls.Add(_txtDbToJsonS2ConectIp, 1, 4);

        AddLabelCell(grid, "dataUsgqty", 5);
        _txtDbToJsonS2DataUsgqty = CreateCsvTextBox("2048", "Recommended dataUsgqty value.");
        grid.Controls.Add(_txtDbToJsonS2DataUsgqty, 1, 5);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Create Response-Writeback Columns", (_, _) => CreateDbToJsonResponseColumns(), true));
        actions.Controls.Add(CreateActionButton("Ping Endpoint", (_, _) => PingDbToJsonEndpoint()));
        actions.Controls.Add(CreateActionButton("Generate Scenario 2 Data", (_, _) => GenerateDbToJsonScenario2()));
        content.Controls.Add(actions);

        _lblDbToJsonPingStatus = CreateInlineStatusLabel("Endpoint has not been checked yet.");
        content.Controls.Add(_lblDbToJsonPingStatus);
        return card;
    }

    private Control CreateDbToJsonScenario3Card()
    {
        var card = CreateCard("Scenario 3 - No New Data Re-run", "Mark all export rows as processed and show the before/after count of rows still flagged as new.", out var content);

        _lblDbToJsonReRunCounts = new Label
        {
            Text = "Unprocessed row count: not checked yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 8)
        };
        content.Controls.Add(_lblDbToJsonReRunCounts);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Refresh Unprocessed Count", (_, _) => RefreshDbToJsonUnprocessedCount(), true));
        actions.Controls.Add(CreateActionButton("Mark All Rows as Processed", (_, _) => MarkDbToJsonRowsProcessed()));
        content.Controls.Add(actions);

        return card;
    }

    private Control CreateSqlQueryScenario1Card()
    {
        var card = CreateCard("SQLQuery Scenario 1 - Scheduled Update Query", "Create and populate the test table with PENDING and DONE rows while leaving UpdatedByJob and UpdatedAt null for verification.", out var content);
        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Connection", 0);
        _cmbSqlQueryConnection = CreateConnectionSelector();
        grid.Controls.Add(_cmbSqlQueryConnection, 1, 0);
        _lblSqlQueryConnectionStatus = CreateInlineStatusLabel("Connection not tested yet.");
        grid.Controls.Add(_lblSqlQueryConnectionStatus, 2, 0);

        AddLabelCell(grid, "Test table name", 1);
        _txtSqlQueryTableName = CreateCsvTextBox("dbo.DS_SqlQuery_Test", "SQLQuery test table name.");
        grid.Controls.Add(_txtSqlQueryTableName, 1, 1);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 1);

        AddLabelCell(grid, "Total row count", 2);
        _numSqlQueryTotalRows = CreateCsvNumeric(1, 100, 8, "Total number of rows to seed.");
        grid.Controls.Add(_numSqlQueryTotalRows, 1, 2);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 2);

        AddLabelCell(grid, "PENDING row count", 3);
        _numSqlQueryPendingRows = CreateCsvNumeric(0, 100, 5, "Number of PENDING rows.");
        grid.Controls.Add(_numSqlQueryPendingRows, 1, 3);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 3);

        AddLabelCell(grid, "DONE row count", 4);
        _numSqlQueryDoneRows = CreateCsvNumeric(0, 100, 3, "Number of DONE rows.");
        grid.Controls.Add(_numSqlQueryDoneRows, 1, 4);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 4);

        AddLabelCell(grid, "Schedule time label", 5);
        var scheduleLabel = new Label
        {
            Text = "2026-04-06 16:00:00",
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 12)
        };
        grid.Controls.Add(scheduleLabel, 1, 5);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 5);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Test Connection", (_, _) => TestSqlQueryConnection(), true));
        actions.Controls.Add(CreateActionButton("Create Table if Missing", (_, _) => CreateSqlQueryTableIfMissing()));
        actions.Controls.Add(CreateActionButton("Generate Scenario 1 Data", (_, _) => GenerateSqlQueryScenario1()));
        content.Controls.Add(actions);

        _lblSqlQueryStatus = new Label
        {
            Text = "No SQLQuery action has run yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblSqlQueryStatus);
        return card;
    }

    private Control CreateSqlQueryReminderCard(string title, string body, Color background, Color foreground)
    {
        var card = CreateCard(title, "Reminder-only scenario. No seed data is required here.", out var content);
        content.Controls.Add(CreateInfoNote("Reminder", body, background, foreground));
        return card;
    }

    private Control CreateProgramScenario1Card()
    {
        var card = CreateCard("ProgramExecution Scenario 1 - Successful Executable Run", "Validate the expected script, create the output folder, and archive any prior output file before the DataSyncer job runs.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Script path", 0);
        _txtProgramS1ScriptPath = CreateCsvTextBox(@"TestResources\TwoDaySoak\program_execution_heartbeat.ps1", "Relative or absolute path to the test script.");
        grid.Controls.Add(_txtProgramS1ScriptPath, 1, 0);

        AddLabelCell(grid, "OutputFile", 1);
        _txtProgramS1OutputFile = CreateCsvTextBox(@"D:\DataSyncerTest\program_execution_heartbeat.log", "Output log file path used by the test script.");
        grid.Controls.Add(_txtProgramS1OutputFile, 1, 1);

        AddLabelCell(grid, "MessagePrefix", 2);
        _txtProgramS1MessagePrefix = CreateCsvTextBox("ProgramExecution functional test", "MessagePrefix argument value.");
        grid.Controls.Add(_txtProgramS1MessagePrefix, 1, 2);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Create Output Folder", (_, _) => CreateProgramOutputFolder(_txtProgramS1OutputFile, UpdateProgramStatus), true));
        actions.Controls.Add(CreateActionButton("Validate Script Path", (_, _) => ValidateProgramScriptPath(_txtProgramS1ScriptPath, UpdateProgramStatus)));
        actions.Controls.Add(CreateActionButton("Archive Prior Output File", (_, _) => ArchiveProgramOutputFile(_txtProgramS1OutputFile, UpdateProgramStatus)));
        content.Controls.Add(actions);

        _lblProgramS1Status = new Label
        {
            Text = "No ProgramExecution validation has run yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblProgramS1Status);
        return card;
    }

    private Control CreateProgramScenario2Card()
    {
        var card = CreateCard("ProgramExecution Scenario 2 - Invalid Executable Path", "Confirm that the negative-test path is still invalid before using it in the job configuration.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Invalid path", 0);
        _txtProgramS2InvalidPath = CreateCsvTextBox(@"D:\DoesNotExist\missing_tool.exe", "Path that should remain missing for negative-path validation.");
        grid.Controls.Add(_txtProgramS2InvalidPath, 1, 0);
        content.Controls.Add(grid);

        content.Controls.Add(CreateActionButton("Confirm Path Does Not Exist", (_, _) => ConfirmProgramPathMissing(), true));

        _lblProgramS2Status = new Label
        {
            Text = "Negative path not checked yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblProgramS2Status);
        return card;
    }

    private Control CreateProgramScenario3Card()
    {
        var card = CreateCard("ProgramExecution Scenario 3 - Executable with Arguments", "Preview the full command line so QA can verify script path and quoted arguments before the job runs.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Script path", 0);
        _txtProgramS3ScriptPath = CreateCsvTextBox(@"TestResources\TwoDaySoak\program_execution_heartbeat.ps1", "Relative or absolute path to the test script.");
        grid.Controls.Add(_txtProgramS3ScriptPath, 1, 0);

        AddLabelCell(grid, "OutputFile", 1);
        _txtProgramS3OutputFile = CreateCsvTextBox(@"D:\DataSyncerTest\arg_check.log", "Output log file used for the argument-validation scenario.");
        grid.Controls.Add(_txtProgramS3OutputFile, 1, 1);

        AddLabelCell(grid, "MessagePrefix", 2);
        _txtProgramS3MessagePrefix = CreateCsvTextBox("Arg Test Value", "MessagePrefix argument containing spaces.");
        grid.Controls.Add(_txtProgramS3MessagePrefix, 1, 2);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Preview Full Command", (_, _) => PreviewProgramCommand(), true));
        actions.Controls.Add(CreateActionButton("Create Output Folder", (_, _) => CreateProgramOutputFolder(_txtProgramS3OutputFile, UpdateProgramStatus)));
        content.Controls.Add(actions);

        _txtProgramS3CommandPreview = new TextBox
        {
            ReadOnly = true,
            Multiline = true,
            Height = 70,
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 8, 0, 0),
            BackColor = Color.FromArgb(248, 250, 252)
        };
        _toolTip.SetToolTip(_txtProgramS3CommandPreview, "Preview of the full command line.");
        content.Controls.Add(_txtProgramS3CommandPreview);
        return card;
    }

    private ComboBox CreateConnectionSelector()
    {
        var comboBox = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Width = 260,
            Margin = new Padding(0, 0, 12, 12)
        };

        PopulateConnectionSelector(comboBox);
        return comboBox;
    }

    private void PopulateConnectionSelector(ComboBox comboBox)
    {
        comboBox.Items.Clear();
        comboBox.Items.Add(new ConnectionChoice("Primary SQL Connection", _settings.SqlConnectionString));
        comboBox.SelectedIndex = 0;
    }

    private DateTimePicker CreateDateTimeInput(DateTime value, string tooltip)
    {
        var picker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Value = value,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12)
        };
        _toolTip.SetToolTip(picker, tooltip);
        return picker;
    }

    private Label CreateInlineStatusLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            MaximumSize = new Size(280, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 12)
        };
    }

    private void TestDbToDbConnection()
    {
        TestConnectionForSelector(_cmbDbToDbConnection, _lblDbToDbConnectionStatus, "DBtoDB");
    }

    private void TestDbToJsonConnection()
    {
        TestConnectionForSelector(_cmbDbToJsonConnection, _lblDbToJsonConnectionStatus, "DBtoJSON");
    }

    private void TestSqlQueryConnection()
    {
        TestConnectionForSelector(_cmbSqlQueryConnection, _lblSqlQueryConnectionStatus, "SQLQuery");
    }

    private void TestConnectionForSelector(ComboBox? selector, Label? statusLabel, string area)
    {
        if (!TryGetSelectedConnectionString(selector, out var connectionString))
        {
            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();

            var message = area + " connection succeeded.";
            SetInlineStatus(statusLabel, message, SuccessColor);
            _logService.LogSuccess(message);
            WriteStatus(message);
            RefreshLogViewer();
        }
        catch (Exception ex)
        {
            var message = area + " connection failed: " + ex.Message;
            SetInlineStatus(statusLabel, message, DangerColor);
            _logService.LogError(message);
            WriteStatus(area + " connection failed");
            RefreshLogViewer();

            MessageBox.Show(
                "Connection test failed." + Environment.NewLine + Environment.NewLine +
                "Connection String:" + Environment.NewLine + MaskConnectionString(connectionString) + Environment.NewLine + Environment.NewLine +
                (ex.InnerException is null ? ex.Message : ex.Message + Environment.NewLine + ex.InnerException.Message),
                area + " Connection",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void CreateDbToDbTablesIfMissing()
    {
        if (_txtDbToDbSourceTable is null || _txtDbToDbDestinationTable is null)
        {
            return;
        }

        string? sourceError = null;
        string? destinationError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToDbConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToDbSourceTable.Text, out var sourceTable, out var sourceDisplay, out sourceError) ||
            !TryNormalizeSqlTableName(_txtDbToDbDestinationTable.Text, out var destinationTable, out var destinationDisplay, out destinationError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoDB", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (!string.IsNullOrWhiteSpace(destinationError))
            {
                MessageBox.Show(destinationError, "DBtoDB", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();

            ExecuteNonQuery(connection, BuildDbToDbCreateTableSql(sourceTable));
            ExecuteNonQuery(connection, BuildDbToDbCreateTableSql(destinationTable));

            SetDbToDbStatus("Created or verified " + sourceDisplay + " and " + destinationDisplay + ".", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoDB", ex);
        }
    }

    private void GenerateDbToDbScenario1()
    {
        if (_numDbToDbS1RowCount is null ||
            _numDbToDbS1SourceIdStart is null ||
            _txtDbToDbS1MachineCodes is null ||
            _txtDbToDbS1WorkCenter is null ||
            _txtDbToDbS1SyncFlag is null ||
            _dtDbToDbS1EventTimeBase is null ||
            _chkDbToDbS1ClearDestination is null ||
            _txtDbToDbSourceTable is null ||
            _txtDbToDbDestinationTable is null)
        {
            return;
        }

        string? sourceError = null;
        string? destinationError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToDbConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToDbSourceTable.Text, out var sourceTable, out _, out sourceError) ||
            !TryNormalizeSqlTableName(_txtDbToDbDestinationTable.Text, out var destinationTable, out _, out destinationError))
        {
            ShowDbValidationError(sourceError, destinationError);
            return;
        }

        var machineCodes = ParseCommaSeparatedValues(_txtDbToDbS1MachineCodes.Text);
        if (machineCodes.Count == 0 || string.IsNullOrWhiteSpace(_txtDbToDbS1WorkCenter.Text) || string.IsNullOrWhiteSpace(_txtDbToDbS1SyncFlag.Text))
        {
            MessageBox.Show("Scenario 1 requires MachineCode values, WorkCenter, and SyncFlag.", "DBtoDB", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rows = new List<DbToDbSeedRow>();
        var startId = Decimal.ToInt32(_numDbToDbS1SourceIdStart.Value);
        var count = Decimal.ToInt32(_numDbToDbS1RowCount.Value);
        for (var i = 0; i < count; i++)
        {
            rows.Add(new DbToDbSeedRow(
                startId + i,
                machineCodes[i % machineCodes.Count],
                _txtDbToDbS1WorkCenter.Text.Trim(),
                _txtDbToDbS1SyncFlag.Text.Trim(),
                _dtDbToDbS1EventTimeBase.Value.AddMinutes(i),
                "Scenario1 Row " + (i + 1).ToString(CultureInfo.InvariantCulture)));
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToDbTablesExist(connection, sourceTable, destinationTable);

            using var transaction = connection.BeginTransaction();
            if (_chkDbToDbS1ClearDestination.Checked)
            {
                ExecuteNonQuery(connection, transaction, "DELETE FROM " + destinationTable);
            }

            InsertDbToDbRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            transaction.Commit();

            SetDbToDbStatus("Scenario 1 prepared " + rows.Count + " source rows.", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoDB", ex);
        }
    }

    private void GenerateDbToDbScenario2()
    {
        if (_numDbToDbS2PreInsertCount is null ||
            _cmbDbToDbS2ConflictBehavior is null ||
            _txtDbToDbSourceTable is null ||
            _txtDbToDbDestinationTable is null)
        {
            return;
        }

        string? sourceError = null;
        string? destinationError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToDbConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToDbSourceTable.Text, out var sourceTable, out _, out sourceError) ||
            !TryNormalizeSqlTableName(_txtDbToDbDestinationTable.Text, out var destinationTable, out _, out destinationError))
        {
            ShowDbValidationError(sourceError, destinationError);
            return;
        }

        var sourceRows = new List<DbToDbSeedRow>();
        for (var i = 0; i < 5; i++)
        {
            sourceRows.Add(new DbToDbSeedRow(
                200011 + i,
                "MC-0" + ((i % 5) + 1).ToString(CultureInfo.InvariantCulture),
                "LINE-A",
                _settings.SyncFlagNotProcessed,
                new DateTime(2026, 4, 6, 11, 30, 0).AddMinutes(i),
                "Scenario2 Fresh Row " + (i + 1).ToString(CultureInfo.InvariantCulture)));
        }

        var staleCount = Decimal.ToInt32(_numDbToDbS2PreInsertCount.Value);
        var staleRows = sourceRows.Take(staleCount)
            .Select((row, index) => row with
            {
                SyncFlag = _settings.SyncFlagProcessed,
                Notes = "Scenario2 Stale Row " + (index + 1).ToString(CultureInfo.InvariantCulture)
            })
            .ToList();

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToDbTablesExist(connection, sourceTable, destinationTable);

            using var transaction = connection.BeginTransaction();
            InsertDbToDbRows(connection, transaction, sourceTable, sourceRows, skipDuplicates: true);
            if (staleRows.Count > 0)
            {
                InsertDbToDbRows(connection, transaction, destinationTable, staleRows, skipDuplicates: true);
            }

            transaction.Commit();

            UpdateDbToDbConflictNote();
            SetDbToDbStatus(
                "Scenario 2 prepared 5 source rows and " + staleRows.Count + " stale destination rows. Conflict mode reminder: " + _cmbDbToDbS2ConflictBehavior.SelectedItem,
                SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoDB", ex);
        }
    }

    private void GenerateDbToDbScenario3()
    {
        if (_numDbToDbS3InsideRangeCount is null ||
            _numDbToDbS3OutsideRangeCount is null ||
            _dtDbToDbS3InsideDate is null ||
            _dtDbToDbS3OutsideDate is null ||
            _txtDbToDbS3WorkCenters is null ||
            _chkDbToDbS3PreloadDestination is null ||
            _txtDbToDbSourceTable is null ||
            _txtDbToDbDestinationTable is null)
        {
            return;
        }

        string? sourceError = null;
        string? destinationError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToDbConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToDbSourceTable.Text, out var sourceTable, out _, out sourceError) ||
            !TryNormalizeSqlTableName(_txtDbToDbDestinationTable.Text, out var destinationTable, out _, out destinationError))
        {
            ShowDbValidationError(sourceError, destinationError);
            return;
        }

        var workCenters = ParseCommaSeparatedValues(_txtDbToDbS3WorkCenters.Text);
        if (workCenters.Count == 0)
        {
            MessageBox.Show("Scenario 3 requires at least one WorkCenter value.", "DBtoDB", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rows = new List<DbToDbSeedRow>();
        var currentId = 200016;
        for (var i = 0; i < Decimal.ToInt32(_numDbToDbS3InsideRangeCount.Value); i++)
        {
            rows.Add(new DbToDbSeedRow(currentId++, "MC-0" + ((i % 5) + 1), workCenters[0], _settings.SyncFlagNotProcessed, _dtDbToDbS3InsideDate.Value.AddMinutes(i), "Inside range"));
        }

        for (var i = 0; i < Decimal.ToInt32(_numDbToDbS3OutsideRangeCount.Value); i++)
        {
            rows.Add(new DbToDbSeedRow(currentId++, "MC-0" + ((i % 5) + 1), workCenters[0], _settings.SyncFlagNotProcessed, _dtDbToDbS3OutsideDate.Value.AddMinutes(i), "Outside range"));
        }

        for (var i = 0; i < 5; i++)
        {
            var workCenter = workCenters[Math.Min(i % workCenters.Count, workCenters.Count - 1)];
            var syncFlag = i == 4 ? _settings.SyncFlagProcessed : _settings.SyncFlagNotProcessed;
            var eventTime = i < 3 ? _dtDbToDbS3InsideDate.Value.AddHours(1).AddMinutes(i) : _dtDbToDbS3OutsideDate.Value.AddHours(1).AddMinutes(i);
            rows.Add(new DbToDbSeedRow(currentId++, "MC-MIX-" + (i + 1).ToString(CultureInfo.InvariantCulture), workCenter, syncFlag, eventTime, "Mixed condition"));
        }

        var staleRows = rows.Take(Math.Min(5, rows.Count))
            .Select((row, index) => row with
            {
                SyncFlag = _settings.SyncFlagProcessed,
                Notes = "Scenario3 Stale Destination " + (index + 1).ToString(CultureInfo.InvariantCulture)
            })
            .ToList();

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToDbTablesExist(connection, sourceTable, destinationTable);

            using var transaction = connection.BeginTransaction();
            InsertDbToDbRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            if (_chkDbToDbS3PreloadDestination.Checked)
            {
                InsertDbToDbRows(connection, transaction, destinationTable, staleRows, skipDuplicates: true);
            }

            transaction.Commit();

            SetDbToDbStatus(
                "Scenario 3 prepared " + rows.Count + " source rows" +
                (_chkDbToDbS3PreloadDestination.Checked ? " and " + staleRows.Count + " stale destination rows." : "."),
                SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoDB", ex);
        }
    }

    private void GenerateDbToJsonScenario1()
    {
        if (_numDbToJsonS1RowCount is null ||
            _txtDbToJsonS1ExportFlag is null ||
            _txtDbToJsonS1DeviceCodes is null ||
            _txtDbToJsonS1ResultCodes is null ||
            _dtDbToJsonS1CreatedAtBase is null ||
            _txtDbToJsonSourceTable is null)
        {
            return;
        }

        string? sourceError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToJsonConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToJsonSourceTable.Text, out var sourceTable, out _, out sourceError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        var deviceCodes = ParseCommaSeparatedValues(_txtDbToJsonS1DeviceCodes.Text);
        var resultCodes = ParseCommaSeparatedValues(_txtDbToJsonS1ResultCodes.Text);
        if (deviceCodes.Count == 0 || resultCodes.Count == 0 || string.IsNullOrWhiteSpace(_txtDbToJsonS1ExportFlag.Text))
        {
            MessageBox.Show("Scenario 1 requires ExportFlag, DeviceCode values, and ResultCode values.", "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rows = new List<DbToJsonSeedRow>();
        var startId = _settings.DbToJsonIdStart;
        for (var i = 0; i < Decimal.ToInt32(_numDbToJsonS1RowCount.Value); i++)
        {
            rows.Add(new DbToJsonSeedRow(
                startId + i,
                _txtDbToJsonS1ExportFlag.Text.Trim(),
                deviceCodes[i % deviceCodes.Count],
                resultCodes[i % resultCodes.Count],
                _dtDbToJsonS1CreatedAtBase.Value.AddMinutes(i),
                "{\"id\":" + (startId + i).ToString(CultureInfo.InvariantCulture) + "}",
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty));
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToJsonTableExists(connection, sourceTable);

            using var transaction = connection.BeginTransaction();
            InsertDbToJsonRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            transaction.Commit();

            SetDbToJsonStatus("Scenario 1 prepared " + rows.Count + " export rows.", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void GenerateDbToJsonScenario2()
    {
        if (_numDbToJsonS2RowCount is null ||
            _txtDbToJsonSourceTable is null ||
            _txtDbToJsonS2CrtfcKey is null ||
            _txtDbToJsonS2UseSe is null ||
            _txtDbToJsonS2SysUser is null ||
            _txtDbToJsonS2ConectIp is null ||
            _txtDbToJsonS2DataUsgqty is null)
        {
            return;
        }

        string? sourceError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToJsonConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToJsonSourceTable.Text, out var sourceTable, out _, out sourceError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        var rows = new List<DbToJsonSeedRow>();
        var startId = _settings.DbToJsonIdStart + 100;
        for (var i = 0; i < Decimal.ToInt32(_numDbToJsonS2RowCount.Value); i++)
        {
            rows.Add(new DbToJsonSeedRow(
                startId + i,
                _settings.SyncFlagNotProcessed,
                "DEV-0" + ((i % 2) + 1).ToString(CultureInfo.InvariantCulture),
                i % 2 == 0 ? "READY" : "PASS",
                new DateTime(2026, 4, 6, 14, 30, 0).AddMinutes(i),
                "{\"api\":true}",
                _txtDbToJsonS2CrtfcKey.Text.Trim(),
                _txtDbToJsonS2UseSe.Text.Trim(),
                _txtDbToJsonS2SysUser.Text.Trim(),
                _txtDbToJsonS2ConectIp.Text.Trim(),
                _txtDbToJsonS2DataUsgqty.Text.Trim()));
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToJsonTableExists(connection, sourceTable);

            using var transaction = connection.BeginTransaction();
            InsertDbToJsonRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            transaction.Commit();

            SetDbToJsonStatus("Scenario 2 prepared " + rows.Count + " API export rows.", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void CreateDbToJsonTableIfMissing()
    {
        if (_txtDbToJsonSourceTable is null)
        {
            return;
        }

        string? sourceError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToJsonConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToJsonSourceTable.Text, out var sourceTable, out var sourceDisplay, out sourceError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToJsonTableExists(connection, sourceTable);

            SetDbToJsonStatus("Created or verified " + sourceDisplay + ".", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void CreateDbToJsonResponseColumns()
    {
        if (_txtDbToJsonSourceTable is null)
        {
            return;
        }

        string? sourceError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToJsonConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToJsonSourceTable.Text, out var sourceTable, out _, out sourceError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToJsonTableExists(connection, sourceTable);
            ExecuteNonQuery(connection, BuildDbToJsonResponseColumnsSql(sourceTable));

            SetDbToJsonStatus("Response-writeback columns verified.", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void PingDbToJsonEndpoint()
    {
        var endpoint = _settings.ApiTestEndpoint;
        if (string.IsNullOrWhiteSpace(endpoint))
        {
            SetInlineStatus(_lblDbToJsonPingStatus, "API endpoint is not configured.", WarningColor);
            return;
        }

        try
        {
            using var client = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(5)
            };
            using var response = client.GetAsync(endpoint).GetAwaiter().GetResult();

            var message = "Endpoint reachable: HTTP " + (int)response.StatusCode;
            SetInlineStatus(_lblDbToJsonPingStatus, message, SuccessColor);
            SetDbToJsonStatus(message, SuccessColor);
        }
        catch (Exception ex)
        {
            var message = "Endpoint ping failed: " + ex.Message;
            SetInlineStatus(_lblDbToJsonPingStatus, message, DangerColor);
            SetDbToJsonStatus(message, DangerColor);
        }
    }

    private void RefreshDbToJsonUnprocessedCount()
    {
        if (_txtDbToJsonSourceTable is null)
        {
            return;
        }

        string? sourceError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToJsonConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToJsonSourceTable.Text, out var sourceTable, out _, out sourceError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToJsonTableExists(connection, sourceTable);
            var count = GetDbToJsonUnprocessedCount(connection, sourceTable);
            if (_lblDbToJsonReRunCounts is not null)
            {
                _lblDbToJsonReRunCounts.Text = "Rows with ExportFlag = " + _settings.SyncFlagNotProcessed + ": " + count.ToString(CultureInfo.InvariantCulture);
                _lblDbToJsonReRunCounts.ForeColor = SuccessColor;
            }
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void MarkDbToJsonRowsProcessed()
    {
        if (_txtDbToJsonSourceTable is null)
        {
            return;
        }

        string? sourceError = null;
        if (!TryGetSelectedConnectionString(_cmbDbToJsonConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtDbToJsonSourceTable.Text, out var sourceTable, out _, out sourceError))
        {
            if (!string.IsNullOrWhiteSpace(sourceError))
            {
                MessageBox.Show(sourceError, "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToJsonTableExists(connection, sourceTable);
            var before = GetDbToJsonUnprocessedCount(connection, sourceTable);

            using var command = new SqlCommand("UPDATE " + sourceTable + " SET ExportFlag = @processed WHERE ExportFlag <> @processed OR ExportFlag IS NULL", connection);
            command.Parameters.AddWithValue("@processed", _settings.SyncFlagProcessed);
            command.ExecuteNonQuery();

            var after = GetDbToJsonUnprocessedCount(connection, sourceTable);
            if (_lblDbToJsonReRunCounts is not null)
            {
                _lblDbToJsonReRunCounts.Text = "Rows with ExportFlag = " + _settings.SyncFlagNotProcessed + ": before=" + before + ", after=" + after;
                _lblDbToJsonReRunCounts.ForeColor = SuccessColor;
            }

            SetDbToJsonStatus("Marked all DBtoJSON rows as processed.", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void GenerateSqlQueryScenario1()
    {
        if (_numSqlQueryTotalRows is null ||
            _numSqlQueryPendingRows is null ||
            _numSqlQueryDoneRows is null ||
            _txtSqlQueryTableName is null)
        {
            return;
        }

        string? tableError = null;
        if (!TryGetSelectedConnectionString(_cmbSqlQueryConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtSqlQueryTableName.Text, out var tableName, out _, out tableError))
        {
            if (!string.IsNullOrWhiteSpace(tableError))
            {
                MessageBox.Show(tableError, "SQLQuery", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        var total = Decimal.ToInt32(_numSqlQueryTotalRows.Value);
        var pending = Decimal.ToInt32(_numSqlQueryPendingRows.Value);
        var done = Decimal.ToInt32(_numSqlQueryDoneRows.Value);
        if (pending + done != total)
        {
            MessageBox.Show("PENDING row count plus DONE row count must equal the total row count.", "SQLQuery", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureSqlQueryTableExists(connection, tableName);

            using var transaction = connection.BeginTransaction();
            ExecuteNonQuery(connection, transaction, "DELETE FROM " + tableName);

            var baseTime = new DateTime(2026, 4, 6, 16, 0, 0);
            for (var i = 0; i < total; i++)
            {
                using var command = new SqlCommand(
                    "INSERT INTO " + tableName + " (RowId, Status, WorkCenter, EventTime, UpdatedByJob, UpdatedAt, Notes) VALUES (@id, @status, @workCenter, @eventTime, NULL, NULL, @notes)",
                    connection,
                    transaction);
                command.Parameters.AddWithValue("@id", _settings.SqlQueryIdStart + i);
                command.Parameters.AddWithValue("@status", i < pending ? "PENDING" : "DONE");
                command.Parameters.AddWithValue("@workCenter", i < pending ? "LINE-A" : "LINE-B");
                command.Parameters.AddWithValue("@eventTime", baseTime.AddMinutes(i));
                command.Parameters.AddWithValue("@notes", "SQLQuery Seed Row " + (i + 1).ToString(CultureInfo.InvariantCulture));
                command.ExecuteNonQuery();
            }

            transaction.Commit();
            SetSqlQueryStatus("Scenario 1 prepared " + total + " rows with UpdatedByJob and UpdatedAt left null.", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("SQLQuery", ex);
        }
    }

    private void CreateSqlQueryTableIfMissing()
    {
        if (_txtSqlQueryTableName is null)
        {
            return;
        }

        string? tableError = null;
        if (!TryGetSelectedConnectionString(_cmbSqlQueryConnection, out var connectionString) ||
            !TryNormalizeSqlTableName(_txtSqlQueryTableName.Text, out var tableName, out var displayName, out tableError))
        {
            if (!string.IsNullOrWhiteSpace(tableError))
            {
                MessageBox.Show(tableError, "SQLQuery", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureSqlQueryTableExists(connection, tableName);
            SetSqlQueryStatus("Created or verified " + displayName + ".", SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("SQLQuery", ex);
        }
    }

    private void CreateProgramOutputFolder(TextBox? outputFileTextBox, Action<string, Color> statusUpdater)
    {
        var outputFile = outputFileTextBox?.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFile))
        {
            statusUpdater("OutputFile path is required.", WarningColor);
            return;
        }

        var directory = Path.GetDirectoryName(outputFile);
        if (string.IsNullOrWhiteSpace(directory))
        {
            statusUpdater("OutputFile must include a folder path.", WarningColor);
            return;
        }

        Directory.CreateDirectory(directory);
        statusUpdater("Created or verified output folder: " + directory, SuccessColor);
    }

    private void ValidateProgramScriptPath(TextBox? scriptPathTextBox, Action<string, Color> statusUpdater)
    {
        var path = ResolveProgramPath(scriptPathTextBox?.Text);
        if (string.IsNullOrWhiteSpace(path))
        {
            statusUpdater("Script path is required.", WarningColor);
            return;
        }

        if (!File.Exists(path))
        {
            statusUpdater("Script path not found: " + path, DangerColor);
            return;
        }

        var extension = Path.GetExtension(path);
        if (!string.Equals(extension, ".ps1", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(extension, ".exe", StringComparison.OrdinalIgnoreCase))
        {
            statusUpdater("Script path must point to a .ps1 or .exe file.", DangerColor);
            return;
        }

        statusUpdater("Validated script path: " + path, SuccessColor);
    }

    private void ArchiveProgramOutputFile(TextBox? outputFileTextBox, Action<string, Color> statusUpdater)
    {
        var outputFile = outputFileTextBox?.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFile))
        {
            statusUpdater("OutputFile path is required.", WarningColor);
            return;
        }

        var directory = Path.GetDirectoryName(outputFile);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (!File.Exists(outputFile))
        {
            statusUpdater("No existing output file to archive.", WarningColor);
            return;
        }

        var archivePath =
            Path.Combine(
                Path.GetDirectoryName(outputFile) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputFile) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + Path.GetExtension(outputFile));

        File.Move(outputFile, archivePath, overwrite: false);
        statusUpdater("Archived prior output to " + archivePath, SuccessColor);
    }

    private void ConfirmProgramPathMissing()
    {
        var invalidPath = _txtProgramS2InvalidPath?.Text.Trim();
        if (string.IsNullOrWhiteSpace(invalidPath))
        {
            UpdateProgramScenario2Status("Invalid path is required.", WarningColor);
            return;
        }

        if (File.Exists(invalidPath) || Directory.Exists(invalidPath))
        {
            UpdateProgramScenario2Status("The path exists. Choose a path that remains invalid for the negative test.", DangerColor);
            return;
        }

        UpdateProgramScenario2Status("Confirmed path does not exist: " + invalidPath, SuccessColor);
    }

    private void PreviewProgramCommand()
    {
        if (_txtProgramS3ScriptPath is null ||
            _txtProgramS3OutputFile is null ||
            _txtProgramS3MessagePrefix is null ||
            _txtProgramS3CommandPreview is null)
        {
            return;
        }

        var resolvedPath = ResolveProgramPath(_txtProgramS3ScriptPath.Text.Trim()) ?? _txtProgramS3ScriptPath.Text.Trim();
        var command =
            "powershell.exe -ExecutionPolicy Bypass -File \"" + resolvedPath + "\" " +
            "-OutputFile \"" + _txtProgramS3OutputFile.Text.Trim() + "\" " +
            "-MessagePrefix \"" + _txtProgramS3MessagePrefix.Text.Trim().Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";

        _txtProgramS3CommandPreview.Text = command;
        UpdateProgramStatus("Previewed full command line.", SuccessColor);
    }

    private void OpenDbToJsonOutputFolder()
    {
        var folder = _txtDbToJsonOutputFolder?.Text.Trim();
        if (string.IsNullOrWhiteSpace(folder))
        {
            return;
        }

        Directory.CreateDirectory(folder);
        OpenPath(folder);
    }

    private void UpdateProgramStatus(string message, Color color)
    {
        if (_lblProgramS1Status is not null)
        {
            _lblProgramS1Status.Text = message;
            _lblProgramS1Status.ForeColor = color;
        }

        RecordRunSummary("ProgramExecution", message, color == SuccessColor ? "Success" : color == DangerColor ? "Error" : "Info");
        _logService.LogInfo("ProgramExecution: " + message);
        WriteStatus(message);
        RefreshLogViewer();
    }

    private void UpdateProgramScenario2Status(string message, Color color)
    {
        if (_lblProgramS2Status is not null)
        {
            _lblProgramS2Status.Text = message;
            _lblProgramS2Status.ForeColor = color;
        }

        RecordRunSummary("ProgramExecution", message, color == SuccessColor ? "Success" : color == DangerColor ? "Error" : "Info");
        _logService.LogInfo("ProgramExecution: " + message);
        WriteStatus(message);
        RefreshLogViewer();
    }

    private void RefreshAdvancedModuleDefaultsFromSettings()
    {
        PopulateConnectionSelectorIfReady(_cmbDbToDbConnection);
        PopulateConnectionSelectorIfReady(_cmbDbToJsonConnection);
        PopulateConnectionSelectorIfReady(_cmbSqlQueryConnection);

        if (_txtDbToDbS1SyncFlag is not null)
        {
            _txtDbToDbS1SyncFlag.Text = _settings.SyncFlagNotProcessed;
        }

        if (_txtDbToJsonS1ExportFlag is not null)
        {
            _txtDbToJsonS1ExportFlag.Text = _settings.SyncFlagNotProcessed;
        }

        RefreshFileSyncerDefaultsFromSettings();
    }

    private void PopulateConnectionSelectorIfReady(ComboBox? comboBox)
    {
        if (comboBox is not null)
        {
            PopulateConnectionSelector(comboBox);
        }
    }

    private static string? ResolveProgramPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return Path.IsPathRooted(value)
            ? value
            : Path.Combine(AppContext.BaseDirectory, value);
    }

    private bool TryGetSelectedConnectionString(ComboBox? selector, out string connectionString)
    {
        connectionString = string.Empty;
        if (selector?.SelectedItem is not ConnectionChoice choice || string.IsNullOrWhiteSpace(choice.ConnectionString))
        {
            MessageBox.Show("A SQL connection must be configured first.", "Connection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        connectionString = choice.ConnectionString;
        return true;
    }

    private static void SetInlineStatus(Label? label, string message, Color color)
    {
        if (label is not null)
        {
            label.Text = message;
            label.ForeColor = color;
        }
    }

    private void SetDbToDbStatus(string message, Color color)
    {
        if (_lblDbToDbStatus is not null)
        {
            _lblDbToDbStatus.Text = message;
            _lblDbToDbStatus.ForeColor = color;
        }

        RecordRunSummary("DBtoDB", message, color == SuccessColor ? "Success" : color == DangerColor ? "Error" : "Info");
        _logService.LogInfo("DBtoDB: " + message);
        WriteStatus(message);
        RefreshLogViewer();
    }

    private void SetDbToJsonStatus(string message, Color color)
    {
        if (_lblDbToJsonStatus is not null)
        {
            _lblDbToJsonStatus.Text = message;
            _lblDbToJsonStatus.ForeColor = color;
        }

        RecordRunSummary("DBtoJSON", message, color == SuccessColor ? "Success" : color == DangerColor ? "Error" : "Info");
        _logService.LogInfo("DBtoJSON: " + message);
        WriteStatus(message);
        RefreshLogViewer();
    }

    private void SetSqlQueryStatus(string message, Color color)
    {
        if (_lblSqlQueryStatus is not null)
        {
            _lblSqlQueryStatus.Text = message;
            _lblSqlQueryStatus.ForeColor = color;
        }

        RecordRunSummary("SQLQuery", message, color == SuccessColor ? "Success" : color == DangerColor ? "Error" : "Info");
        _logService.LogInfo("SQLQuery: " + message);
        WriteStatus(message);
        RefreshLogViewer();
    }

    private void UpdateDbToDbConflictNote()
    {
        if (_cmbDbToDbS2ConflictBehavior?.SelectedItem is null)
        {
            return;
        }

        SetInlineStatus(
            _lblDbToDbS2ConflictNote,
            "Verify the job configuration matches the selected conflict behavior: " + _cmbDbToDbS2ConflictBehavior.SelectedItem + ".",
            AccentColor);
    }

    private void ShowDbValidationError(string? first, string? second)
    {
        var message = !string.IsNullOrWhiteSpace(first) ? first : second;
        if (!string.IsNullOrWhiteSpace(message))
        {
            MessageBox.Show(message, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void HandleDbActionError(string area, Exception ex)
    {
        RecordRunSummary(area, ex.Message, "Error");
        _logService.LogError(area + ": " + ex.Message);
        RefreshLogViewer();
        WriteStatus(area + " action failed");

        MessageBox.Show(
            area + " action failed." + Environment.NewLine + Environment.NewLine + ex.Message,
            area,
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }

    private static void ExecuteNonQuery(SqlConnection connection, string sql)
    {
        using var command = new SqlCommand(sql, connection);
        command.ExecuteNonQuery();
    }

    private static void ExecuteNonQuery(SqlConnection connection, SqlTransaction transaction, string sql)
    {
        using var command = new SqlCommand(sql, connection, transaction);
        command.ExecuteNonQuery();
    }

    private void EnsureDbToDbTablesExist(SqlConnection connection, string sourceTable, string destinationTable)
    {
        ExecuteNonQuery(connection, BuildDbToDbCreateTableSql(sourceTable));
        ExecuteNonQuery(connection, BuildDbToDbCreateTableSql(destinationTable));
    }

    private void EnsureDbToJsonTableExists(SqlConnection connection, string sourceTable)
    {
        ExecuteNonQuery(connection, BuildDbToJsonCreateTableSql(sourceTable));
    }

    private void EnsureSqlQueryTableExists(SqlConnection connection, string tableName)
    {
        ExecuteNonQuery(connection, BuildSqlQueryCreateTableSql(tableName));
    }

    private static string BuildDbToDbCreateTableSql(string tableName)
    {
        var schemaName = tableName.Split('.')[0].Trim('[', ']');
        var simpleName = tableName.Split('.')[1].Trim('[', ']');

        return $@"
IF SCHEMA_ID('{schemaName}') IS NULL
    EXEC('CREATE SCHEMA [{schemaName}]');

IF OBJECT_ID(N'[{schemaName}].[{simpleName}]', N'U') IS NULL
BEGIN
    CREATE TABLE {tableName} (
        SourceId INT NOT NULL PRIMARY KEY,
        MachineCode NVARCHAR(50) NOT NULL,
        WorkCenter NVARCHAR(50) NOT NULL,
        SyncFlag NVARCHAR(20) NOT NULL,
        EventTime DATETIME2 NOT NULL,
        Notes NVARCHAR(200) NULL,
        CreatedAt DATETIME2 NOT NULL CONSTRAINT DF_{simpleName}_CreatedAt DEFAULT SYSUTCDATETIME()
    );
END";
    }

    private static string BuildDbToJsonCreateTableSql(string tableName)
    {
        var schemaName = tableName.Split('.')[0].Trim('[', ']');
        var simpleName = tableName.Split('.')[1].Trim('[', ']');

        return $@"
IF SCHEMA_ID('{schemaName}') IS NULL
    EXEC('CREATE SCHEMA [{schemaName}]');

IF OBJECT_ID(N'[{schemaName}].[{simpleName}]', N'U') IS NULL
BEGIN
    CREATE TABLE {tableName} (
        ExportId INT NOT NULL PRIMARY KEY,
        ExportFlag NVARCHAR(20) NOT NULL,
        DeviceCode NVARCHAR(50) NOT NULL,
        ResultCode NVARCHAR(50) NOT NULL,
        CreatedAt DATETIME2 NOT NULL,
        JsonPayload NVARCHAR(MAX) NULL,
        CrtfcKey NVARCHAR(100) NULL,
        UseSe NVARCHAR(50) NULL,
        SysUser NVARCHAR(100) NULL,
        ConectIp NVARCHAR(50) NULL,
        DataUsgqty NVARCHAR(50) NULL
    );
END";
    }

    private static string BuildDbToJsonResponseColumnsSql(string tableName)
    {
        return $@"
IF COL_LENGTH('{tableName.Replace("[", string.Empty).Replace("]", string.Empty)}', 'ApiResultCode') IS NULL
    ALTER TABLE {tableName} ADD ApiResultCode NVARCHAR(50) NULL;
IF COL_LENGTH('{tableName.Replace("[", string.Empty).Replace("]", string.Empty)}', 'ApiResultMessage') IS NULL
    ALTER TABLE {tableName} ADD ApiResultMessage NVARCHAR(400) NULL;
IF COL_LENGTH('{tableName.Replace("[", string.Empty).Replace("]", string.Empty)}', 'ApiReceivedAt') IS NULL
    ALTER TABLE {tableName} ADD ApiReceivedAt DATETIME2 NULL;";
    }

    private static string BuildSqlQueryCreateTableSql(string tableName)
    {
        var schemaName = tableName.Split('.')[0].Trim('[', ']');
        var simpleName = tableName.Split('.')[1].Trim('[', ']');

        return $@"
IF SCHEMA_ID('{schemaName}') IS NULL
    EXEC('CREATE SCHEMA [{schemaName}]');

IF OBJECT_ID(N'[{schemaName}].[{simpleName}]', N'U') IS NULL
BEGIN
    CREATE TABLE {tableName} (
        RowId INT NOT NULL PRIMARY KEY,
        Status NVARCHAR(20) NOT NULL,
        WorkCenter NVARCHAR(50) NOT NULL,
        EventTime DATETIME2 NOT NULL,
        UpdatedByJob NVARCHAR(100) NULL,
        UpdatedAt DATETIME2 NULL,
        Notes NVARCHAR(200) NULL
    );
END";
    }

    private void InsertDbToDbRows(SqlConnection connection, SqlTransaction transaction, string tableName, IReadOnlyList<DbToDbSeedRow> rows, bool skipDuplicates)
    {
        foreach (var row in rows)
        {
            try
            {
                using var command = new SqlCommand(
                    "INSERT INTO " + tableName + " (SourceId, MachineCode, WorkCenter, SyncFlag, EventTime, Notes) VALUES (@id, @machineCode, @workCenter, @syncFlag, @eventTime, @notes)",
                    connection,
                    transaction);
                command.Parameters.AddWithValue("@id", row.SourceId);
                command.Parameters.AddWithValue("@machineCode", row.MachineCode);
                command.Parameters.AddWithValue("@workCenter", row.WorkCenter);
                command.Parameters.AddWithValue("@syncFlag", row.SyncFlag);
                command.Parameters.AddWithValue("@eventTime", row.EventTime);
                command.Parameters.AddWithValue("@notes", row.Notes);
                command.ExecuteNonQuery();
            }
            catch (SqlException ex) when (ex.Number == 2627 || ex.Number == 2601)
            {
                _logService.LogWarning("Duplicate SourceId skipped for " + tableName + ": " + row.SourceId);
                if (!skipDuplicates)
                {
                    throw;
                }
            }
        }
    }

    private void InsertDbToJsonRows(SqlConnection connection, SqlTransaction transaction, string tableName, IReadOnlyList<DbToJsonSeedRow> rows, bool skipDuplicates)
    {
        foreach (var row in rows)
        {
            try
            {
                using var command = new SqlCommand(
                    "INSERT INTO " + tableName + " (ExportId, ExportFlag, DeviceCode, ResultCode, CreatedAt, JsonPayload, CrtfcKey, UseSe, SysUser, ConectIp, DataUsgqty) VALUES (@id, @flag, @device, @result, @createdAt, @payload, @crtfcKey, @useSe, @sysUser, @conectIp, @dataUsgqty)",
                    connection,
                    transaction);
                command.Parameters.AddWithValue("@id", row.ExportId);
                command.Parameters.AddWithValue("@flag", row.ExportFlag);
                command.Parameters.AddWithValue("@device", row.DeviceCode);
                command.Parameters.AddWithValue("@result", row.ResultCode);
                command.Parameters.AddWithValue("@createdAt", row.CreatedAt);
                command.Parameters.AddWithValue("@payload", row.JsonPayload);
                command.Parameters.AddWithValue("@crtfcKey", row.CrtfcKey);
                command.Parameters.AddWithValue("@useSe", row.UseSe);
                command.Parameters.AddWithValue("@sysUser", row.SysUser);
                command.Parameters.AddWithValue("@conectIp", row.ConectIp);
                command.Parameters.AddWithValue("@dataUsgqty", row.DataUsgqty);
                command.ExecuteNonQuery();
            }
            catch (SqlException ex) when (ex.Number == 2627 || ex.Number == 2601)
            {
                _logService.LogWarning("Duplicate ExportId skipped for " + tableName + ": " + row.ExportId);
                if (!skipDuplicates)
                {
                    throw;
                }
            }
        }
    }

    private int GetDbToJsonUnprocessedCount(SqlConnection connection, string tableName)
    {
        using var command = new SqlCommand("SELECT COUNT(*) FROM " + tableName + " WHERE ExportFlag = @flag", connection);
        command.Parameters.AddWithValue("@flag", _settings.SyncFlagNotProcessed);
        return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private static bool TryNormalizeSqlTableName(string input, out string normalizedName, out string displayName, out string? error)
    {
        normalizedName = string.Empty;
        displayName = string.Empty;
        error = null;

        var parts = input.Trim().Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 1)
        {
            parts = new[] { "dbo", parts[0] };
        }

        if (parts.Length != 2)
        {
            error = "Table name must be in the format schema.table or table.";
            return false;
        }

        if (!parts.All(IsSqlIdentifier))
        {
            error = "Table names may contain only letters, digits, and underscores, and must begin with a letter or underscore.";
            return false;
        }

        normalizedName = "[" + parts[0] + "].[" + parts[1] + "]";
        displayName = parts[0] + "." + parts[1];
        return true;
    }

    private static bool IsSqlIdentifier(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (!(char.IsLetter(value[0]) || value[0] == '_'))
        {
            return false;
        }

        for (var i = 1; i < value.Length; i++)
        {
            if (!(char.IsLetterOrDigit(value[i]) || value[i] == '_'))
            {
                return false;
            }
        }

        return true;
    }

    private sealed record ConnectionChoice(string Name, string ConnectionString)
    {
        public override string ToString() => Name;
    }

    private sealed record DbToDbSeedRow(int SourceId, string MachineCode, string WorkCenter, string SyncFlag, DateTime EventTime, string Notes);

    private sealed record DbToJsonSeedRow(
        int ExportId,
        string ExportFlag,
        string DeviceCode,
        string ResultCode,
        DateTime CreatedAt,
        string JsonPayload,
        string CrtfcKey,
        string UseSe,
        string SysUser,
        string ConectIp,
        string DataUsgqty);
}
