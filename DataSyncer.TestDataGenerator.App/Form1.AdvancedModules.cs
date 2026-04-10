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
    private ComboBox? _cmbDbToDbScheduleMode;
    private DateTimePicker? _dtDbToDbScheduleAt;
    private NumericUpDown? _numDbToDbScheduleIntervalMinutes;
    private ComboBox? _cmbDbToDbScheduleTarget;
    private Label? _lblDbToDbScheduleStatus;
    private ProgressBar? _progressDbToDbScenario3;
    private Label? _lblDbToDbScenario3Progress;
    private readonly System.Windows.Forms.Timer _dbToDbScheduleTimer = new() { Interval = 1000 };
    private DateTime? _dbToDbScheduledRunAt;
    private DbToDbScheduleMode _armedDbToDbScheduleMode = DbToDbScheduleMode.OneTime;
    private TimeSpan? _dbToDbScheduleInterval;
    private DbToDbAutomationExecutionContext? _dbToDbAutomationExecutionContext;
    private int _dbToDbAutomationRunSequence;

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
    private ComboBox? _cmbDbToJsonScheduleMode;
    private DateTimePicker? _dtDbToJsonScheduleAt;
    private NumericUpDown? _numDbToJsonScheduleIntervalMinutes;
    private ComboBox? _cmbDbToJsonScheduleTarget;
    private Label? _lblDbToJsonScheduleStatus;
    private readonly System.Windows.Forms.Timer _dbToJsonScheduleTimer = new() { Interval = 1000 };
    private DateTime? _dbToJsonScheduledRunAt;
    private DbToJsonScheduleMode _armedDbToJsonScheduleMode = DbToJsonScheduleMode.OneTime;
    private TimeSpan? _dbToJsonScheduleInterval;
    private DbToJsonAutomationExecutionContext? _dbToJsonAutomationExecutionContext;
    private int _dbToJsonAutomationRunSequence;

    private ComboBox? _cmbSqlQueryConnection;
    private TextBox? _txtSqlQueryTableName;
    private Label? _lblSqlQueryConnectionStatus;
    private Label? _lblSqlQueryStatus;
    private NumericUpDown? _numSqlQueryTotalRows;
    private NumericUpDown? _numSqlQueryPendingRows;
    private NumericUpDown? _numSqlQueryDoneRows;
    private DateTimePicker? _dtSqlQueryEventTimeBase;
    private TextBox? _txtSqlQueryScenario1UpdateSql;
    private TextBox? _txtSqlQueryScenario2InvalidSql;
    private TextBox? _txtSqlQueryScenario3BrokenConnection;
    private ComboBox? _cmbSqlQueryScheduleMode;
    private DateTimePicker? _dtSqlQueryScheduleAt;
    private NumericUpDown? _numSqlQueryScheduleIntervalMinutes;
    private ComboBox? _cmbSqlQueryScheduleTarget;
    private Label? _lblSqlQueryScheduleStatus;
    private readonly System.Windows.Forms.Timer _sqlQueryScheduleTimer = new() { Interval = 1000 };
    private DateTime? _sqlQueryScheduledRunAt;
    private SqlQueryScheduleMode _armedSqlQueryScheduleMode = SqlQueryScheduleMode.OneTime;
    private TimeSpan? _sqlQueryScheduleInterval;
    private SqlQueryAutomationExecutionContext? _sqlQueryAutomationExecutionContext;
    private int _sqlQueryAutomationRunSequence;

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
    private Label? _lblProgramS3Status;
    private ComboBox? _cmbProgramScheduleMode;
    private DateTimePicker? _dtProgramScheduleAt;
    private NumericUpDown? _numProgramScheduleIntervalMinutes;
    private ComboBox? _cmbProgramScheduleTarget;
    private Label? _lblProgramScheduleStatus;
    private readonly System.Windows.Forms.Timer _programScheduleTimer = new() { Interval = 1000 };
    private DateTime? _programScheduledRunAt;
    private ProgramScheduleMode _armedProgramScheduleMode = ProgramScheduleMode.OneTime;
    private TimeSpan? _programScheduleInterval;
    private ProgramAutomationExecutionContext? _programAutomationExecutionContext;
    private int _programAutomationRunSequence;

    private Control CreateDbToDbPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateDbToDbConnectionCard());
        stack.Controls.Add(CreateDbToDbAutomationCard());
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
        stack.Controls.Add(CreateDbToJsonAutomationCard());
        stack.Controls.Add(CreateDbToJsonScenario1Card());
        stack.Controls.Add(CreateDbToJsonScenario2Card());
        stack.Controls.Add(CreateDbToJsonScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateSqlQueryPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateSqlQueryAutomationCard());
        stack.Controls.Add(CreateSqlQueryScenario1Card());
        stack.Controls.Add(CreateSqlQueryScenario2Card());
        stack.Controls.Add(CreateSqlQueryScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateProgramExecutionPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateProgramAutomationCard());
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

    private Control CreateDbToDbAutomationCard()
    {
        var card = CreateCard("DBtoDB Automation Timer", "Keep this application open and arm recurring DBtoDB scenario preparation. For example, choose Every N minutes and set the value to 1 to refresh DB test data every minute.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Automation mode", 0);
        _cmbDbToDbScheduleMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "DBtoDB automation mode"
        };
        _cmbDbToDbScheduleMode.Items.AddRange(new object[]
        {
            "Run Once At Selected Time",
            "Repeat Daily At Selected Time",
            "Repeat Every N Minutes"
        });
        _cmbDbToDbScheduleMode.SelectedIndex = 0;
        _cmbDbToDbScheduleMode.SelectedIndexChanged += (_, _) => UpdateDbToDbAutomationControlState();
        _toolTip.SetToolTip(_cmbDbToDbScheduleMode, "Choose whether the DBtoDB scheduler runs once, repeats daily, or repeats every N minutes.");
        grid.Controls.Add(_cmbDbToDbScheduleMode, 1, 0);

        AddLabelCell(grid, "Run at", 1);
        _dtDbToDbScheduleAt = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Value = DateTime.Now.AddMinutes(5),
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "DBtoDB scheduled run time"
        };
        _toolTip.SetToolTip(_dtDbToDbScheduleAt, "Date and time used for one-time and daily DBtoDB automation.");
        grid.Controls.Add(_dtDbToDbScheduleAt, 1, 1);

        AddLabelCell(grid, "Every N minutes", 2);
        _numDbToDbScheduleIntervalMinutes = CreateCsvNumeric(1, 1440, 1, "Interval in minutes for recurring DBtoDB preparation.");
        _numDbToDbScheduleIntervalMinutes.AccessibleName = "DBtoDB automation interval minutes";
        grid.Controls.Add(_numDbToDbScheduleIntervalMinutes, 1, 2);

        AddLabelCell(grid, "Timer target", 3);
        _cmbDbToDbScheduleTarget = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "DBtoDB automation target"
        };
        _cmbDbToDbScheduleTarget.Items.AddRange(new object[]
        {
            "Generate Scenario 1",
            "Generate Scenario 2",
            "Generate Scenario 3",
            "Generate All DBtoDB Scenarios"
        });
        _cmbDbToDbScheduleTarget.SelectedIndex = 0;
        _toolTip.SetToolTip(_cmbDbToDbScheduleTarget, "Choose which DBtoDB scenario setup the timer should generate.");
        grid.Controls.Add(_cmbDbToDbScheduleTarget, 1, 3);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Arm Timer", (_, _) => ArmDbToDbSchedule(), true));
        actions.Controls.Add(CreateActionButton("Cancel Timer", (_, _) => CancelDbToDbSchedule()));
        actions.Controls.Add(CreateActionButton("Run Target Now", (_, _) => RunDbToDbScheduledTarget()));
        content.Controls.Add(actions);

        _lblDbToDbScheduleStatus = new Label
        {
            Text = "DBtoDB scheduler idle. Choose a mode, then arm the timer.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblDbToDbScheduleStatus);

        content.Controls.Add(CreateInfoNote(
            "Automation note",
            "Scheduled DBtoDB runs use a new SourceId batch offset each time so repeated timer runs do not keep colliding with the same IDs.",
            AccentSoft,
            AccentColor));

        _dbToDbScheduleTimer.Tick -= OnDbToDbScheduleTimerTick;
        _dbToDbScheduleTimer.Tick += OnDbToDbScheduleTimerTick;

        UpdateDbToDbAutomationControlState();
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
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use this setup when the DBtoDB job should fetch fresh source rows, copy them into the destination, and then update the source flags from the not-processed value to the processed value. The destination table is created automatically if needed.",
            AccentSoft,
            AccentColor));
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
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
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
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "This setup leaves the same records in source and destination so the DBtoDB job can exercise duplicate-safe handling. Use the conflict reminder to align your DataSyncer job configuration with the scenario you want to validate.",
            WarningSoft,
            WarningColor));
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

        _progressDbToDbScenario3 = new ProgressBar
        {
            Style = ProgressBarStyle.Continuous,
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            Height = 22,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 10, 0, 0)
        };
        content.Controls.Add(_progressDbToDbScenario3);

        _lblDbToDbScenario3Progress = new Label
        {
            Text = "Scenario 3 progress is idle.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblDbToDbScenario3Progress);

        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use this setup to validate delete-and-reinsert behavior for a selected date range and filter condition. When preloading destination rows is enabled, the destination is primed with matching stale rows so the manual copy path can remove them before reinsert.",
            AccentSoft,
            AccentColor));
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

    private Control CreateDbToJsonAutomationCard()
    {
        var card = CreateCard("DBtoJSON Automation Timer", "Keep this application open and arm recurring DBtoJSON preparation so DataSyncer can keep finding fresh export rows at fixed time gaps.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Automation mode", 0);
        _cmbDbToJsonScheduleMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "DBtoJSON automation mode"
        };
        _cmbDbToJsonScheduleMode.Items.AddRange(new object[]
        {
            "Run Once At Selected Time",
            "Repeat Daily At Selected Time",
            "Repeat Every N Minutes"
        });
        _cmbDbToJsonScheduleMode.SelectedIndex = 0;
        _cmbDbToJsonScheduleMode.SelectedIndexChanged += (_, _) => UpdateDbToJsonAutomationControlState();
        _toolTip.SetToolTip(_cmbDbToJsonScheduleMode, "Choose whether the DBtoJSON scheduler runs once, repeats daily, or repeats every N minutes.");
        grid.Controls.Add(_cmbDbToJsonScheduleMode, 1, 0);

        AddLabelCell(grid, "Run at", 1);
        _dtDbToJsonScheduleAt = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Value = DateTime.Now.AddMinutes(5),
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "DBtoJSON scheduled run time"
        };
        _toolTip.SetToolTip(_dtDbToJsonScheduleAt, "Date and time used for one-time and daily DBtoJSON automation.");
        grid.Controls.Add(_dtDbToJsonScheduleAt, 1, 1);

        AddLabelCell(grid, "Every N minutes", 2);
        _numDbToJsonScheduleIntervalMinutes = CreateCsvNumeric(1, 1440, 1, "Interval in minutes for recurring DBtoJSON preparation.");
        _numDbToJsonScheduleIntervalMinutes.AccessibleName = "DBtoJSON automation interval minutes";
        grid.Controls.Add(_numDbToJsonScheduleIntervalMinutes, 1, 2);

        AddLabelCell(grid, "Timer target", 3);
        _cmbDbToJsonScheduleTarget = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "DBtoJSON automation target"
        };
        _cmbDbToJsonScheduleTarget.Items.AddRange(new object[]
        {
            "Generate Scenario 1",
            "Generate Scenario 2",
            "Prepare Scenario 3 Re-run State",
            "Generate Scenarios 1 + 2"
        });
        _cmbDbToJsonScheduleTarget.SelectedIndex = 0;
        _toolTip.SetToolTip(_cmbDbToJsonScheduleTarget, "Choose which DBtoJSON scenario setup the timer should generate.");
        grid.Controls.Add(_cmbDbToJsonScheduleTarget, 1, 3);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Arm Timer", (_, _) => ArmDbToJsonSchedule(), true));
        actions.Controls.Add(CreateActionButton("Cancel Timer", (_, _) => CancelDbToJsonSchedule()));
        actions.Controls.Add(CreateActionButton("Run Target Now", (_, _) => RunDbToJsonScheduledTarget()));
        content.Controls.Add(actions);

        _lblDbToJsonScheduleStatus = new Label
        {
            Text = "DBtoJSON scheduler idle. Choose a mode, then arm the timer.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblDbToJsonScheduleStatus);

        content.Controls.Add(CreateInfoNote(
            "Parallel run note",
            "This timer prepares fresh DB export rows on a cadence while DataSyncer runs separately. The generator app seeds the source table; the actual JSON file creation or API send is still performed by DataSyncer.",
            AccentSoft,
            AccentColor));

        _dbToJsonScheduleTimer.Tick -= OnDbToJsonScheduleTimerTick;
        _dbToJsonScheduleTimer.Tick += OnDbToJsonScheduleTimerTick;

        UpdateDbToJsonAutomationControlState();
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
        grid.Controls.Add(_txtDbToJsonOutputFolder, 1, 5);
        var outputActions = CreateInlineActionPanel();
        outputActions.Controls.Add(CreateMiniButton("Browse", (_, _) => BrowseForFolder(_txtDbToJsonOutputFolder, "Select the DBtoJSON output folder")));
        outputActions.Controls.Add(CreateMiniButton("Open", (_, _) => OpenDbToJsonOutputFolder()));
        outputActions.Controls.Add(CreateMiniButton("Prepare", (_, _) => PrepareDbToJsonOutputFolder()));
        grid.Controls.Add(outputActions, 2, 5);

        content.Controls.Add(grid);
        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate Scenario 1 Data", (_, _) => GenerateDbToJsonScenario1(), true));
        actions.Controls.Add(CreateActionButton("Prepare Output Folder", (_, _) => PrepareDbToJsonOutputFolder()));
        actions.Controls.Add(CreateActionButton("Open Output Folder", (_, _) => OpenDbToJsonOutputFolder()));
        content.Controls.Add(actions);
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use this setup when DataSyncer is configured for folder output with JSONArray format. After the job runs, expect one JSON file in the configured folder, 5 exported records, processed source flags, and an updated state file.",
            AccentSoft,
            AccentColor));
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
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use this setup when DataSyncer is configured for API export. The API should receive the payload, success status should appear in logs, response-mapped columns should be written back to the source table, and flags should update only after the API call succeeds.",
            AccentSoft,
            AccentColor));
        return card;
    }

    private Control CreateDbToJsonScenario3Card()
    {
        var card = CreateCard("Scenario 3 - No New Data Re-run", "Prepare the source table so a second DBtoJSON run sees zero new rows and does not create duplicate exports.", out var content);

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
        actions.Controls.Add(CreateActionButton("Prepare Re-run State", (_, _) => PrepareDbToJsonScenario3()));
        actions.Controls.Add(CreateActionButton("Mark All Rows as Processed", (_, _) => MarkDbToJsonRowsProcessed()));
        content.Controls.Add(actions);
        content.Controls.Add(CreateInfoNote(
            "Expected after rerun",
            "After DataSyncer has already exported the rows once, this setup helps validate that a second run finds zero new records, produces no extra JSON file or API send, and reports zero processed records in the latest state.",
            WarningSoft,
            WarningColor));

        return card;
    }

    private Control CreateSqlQueryAutomationCard()
    {
        var card = CreateCard("SQLQuery Automation Timer", "Keep this application open and let it refresh SQLQuery prep on a schedule. Use Every N minutes when you want DataSyncer to keep finding fresh rows at fixed gaps.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Automation mode", 0);
        _cmbSqlQueryScheduleMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "SQLQuery automation mode"
        };
        _cmbSqlQueryScheduleMode.Items.AddRange(new object[]
        {
            "Run Once At Selected Time",
            "Repeat Daily At Selected Time",
            "Repeat Every N Minutes"
        });
        _cmbSqlQueryScheduleMode.SelectedIndex = 0;
        _cmbSqlQueryScheduleMode.SelectedIndexChanged += (_, _) => UpdateSqlQueryAutomationControlState();
        grid.Controls.Add(_cmbSqlQueryScheduleMode, 1, 0);

        AddLabelCell(grid, "Run at", 1);
        _dtSqlQueryScheduleAt = CreateDateTimeInput(DateTime.Now.AddMinutes(5), "Date and time used for one-time and daily SQLQuery preparation.");
        grid.Controls.Add(_dtSqlQueryScheduleAt, 1, 1);

        AddLabelCell(grid, "Every N minutes", 2);
        _numSqlQueryScheduleIntervalMinutes = CreateCsvNumeric(1, 1440, 1, "Interval in minutes for recurring SQLQuery preparation.");
        _numSqlQueryScheduleIntervalMinutes.AccessibleName = "SQLQuery automation interval minutes";
        grid.Controls.Add(_numSqlQueryScheduleIntervalMinutes, 1, 2);

        AddLabelCell(grid, "Timer target", 3);
        _cmbSqlQueryScheduleTarget = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "SQLQuery automation target"
        };
        _cmbSqlQueryScheduleTarget.Items.AddRange(new object[]
        {
            "Scenario 1 - Scheduled Update Data",
            "Scenario 2 - Invalid SQL Sample",
            "Scenario 3 - Broken Connection Sample",
            "All SQLQuery Scenarios"
        });
        _cmbSqlQueryScheduleTarget.SelectedIndex = 0;
        grid.Controls.Add(_cmbSqlQueryScheduleTarget, 1, 3);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Arm Timer", (_, _) => ArmSqlQuerySchedule(), true));
        actions.Controls.Add(CreateActionButton("Cancel Timer", (_, _) => CancelSqlQuerySchedule()));
        actions.Controls.Add(CreateActionButton("Run Target Now", (_, _) => RunSqlQueryScheduledTarget()));
        content.Controls.Add(actions);

        _lblSqlQueryScheduleStatus = new Label
        {
            Text = "SQLQuery scheduler idle. Choose a mode, then arm the timer.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblSqlQueryScheduleStatus);

        _sqlQueryScheduleTimer.Tick -= OnSqlQueryScheduleTimerTick;
        _sqlQueryScheduleTimer.Tick += OnSqlQueryScheduleTimerTick;
        UpdateSqlQueryAutomationControlState();
        return card;
    }

    private Control CreateSqlQueryScenario1Card()
    {
        var card = CreateCard("SQLQuery Scenario 1 - Scheduled Update Query", "Create the test rows, keep the processed columns null, and generate a copy-ready UPDATE statement for DataSyncer.", out var content);
        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Connection", 0);
        _cmbSqlQueryConnection = CreateConnectionSelector();
        grid.Controls.Add(_cmbSqlQueryConnection, 1, 0);
        _lblSqlQueryConnectionStatus = CreateInlineStatusLabel("Connection not tested yet.");
        grid.Controls.Add(_lblSqlQueryConnectionStatus, 2, 0);

        AddLabelCell(grid, "Test table name", 1);
        _txtSqlQueryTableName = CreateCsvTextBox("dbo.DS_SqlQuery_Test", "SQLQuery test table name.");
        _txtSqlQueryTableName.TextChanged += (_, _) => RefreshSqlQueryScenarioTemplates();
        grid.Controls.Add(_txtSqlQueryTableName, 1, 1);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 1);

        AddLabelCell(grid, "Total row count", 2);
        _numSqlQueryTotalRows = CreateCsvNumeric(1, 100, 8, "Total number of rows to seed.");
        _numSqlQueryTotalRows.ValueChanged += (_, _) => RefreshSqlQueryScenarioTemplates();
        grid.Controls.Add(_numSqlQueryTotalRows, 1, 2);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 2);

        AddLabelCell(grid, "PENDING row count", 3);
        _numSqlQueryPendingRows = CreateCsvNumeric(0, 100, 5, "Number of PENDING rows.");
        _numSqlQueryPendingRows.ValueChanged += (_, _) => RefreshSqlQueryScenarioTemplates();
        grid.Controls.Add(_numSqlQueryPendingRows, 1, 3);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 3);

        AddLabelCell(grid, "DONE row count", 4);
        _numSqlQueryDoneRows = CreateCsvNumeric(0, 100, 3, "Number of DONE rows.");
        _numSqlQueryDoneRows.ValueChanged += (_, _) => RefreshSqlQueryScenarioTemplates();
        grid.Controls.Add(_numSqlQueryDoneRows, 1, 4);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 4);

        AddLabelCell(grid, "Base event time", 5);
        _dtSqlQueryEventTimeBase = CreateDateTimeInput(new DateTime(2026, 4, 6, 16, 0, 0), "Base EventTime used when seeding SQLQuery scenario rows.");
        _dtSqlQueryEventTimeBase.ValueChanged += (_, _) => RefreshSqlQueryScenarioTemplates();
        grid.Controls.Add(_dtSqlQueryEventTimeBase, 1, 5);
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
        actions.Controls.Add(CreateActionButton("Refresh SQL Samples", (_, _) => RefreshSqlQueryScenarioTemplates()));
        actions.Controls.Add(CreateActionButton("Generate Scenario 1 Data", (_, _) => GenerateSqlQueryScenario1()));
        content.Controls.Add(actions);

        _txtSqlQueryScenario1UpdateSql = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            Height = 92,
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 8, 0, 0),
            BackColor = Color.FromArgb(248, 250, 252)
        };
        _toolTip.SetToolTip(_txtSqlQueryScenario1UpdateSql, "Suggested SQLQuery UPDATE statement for Scenario 1.");
        content.Controls.Add(_txtSqlQueryScenario1UpdateSql);

        _lblSqlQueryStatus = new Label
        {
            Text = "No SQLQuery action has run yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblSqlQueryStatus);
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use the generated UPDATE sample with a near-future SQLQuery schedule. After DataSyncer runs, only the PENDING LINE-A rows in the prepared RowId range should update, and the service log row count should match the affected-row count in the table.",
            AccentSoft,
            AccentColor));
        RefreshSqlQueryScenarioTemplates();
        return card;
    }

    private Control CreateSqlQueryScenario2Card()
    {
        var card = CreateCard("SQLQuery Scenario 2 - Invalid SQL", "Generate a malformed statement you can paste into the DataSyncer job to verify clear error logging without silent success.", out var content);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Load Invalid SQL Sample", (_, _) => PrepareSqlQueryScenario2(), true));
        content.Controls.Add(actions);

        _txtSqlQueryScenario2InvalidSql = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            Height = 76,
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 8, 0, 0),
            BackColor = Color.FromArgb(248, 250, 252)
        };
        _toolTip.SetToolTip(_txtSqlQueryScenario2InvalidSql, "Malformed SQL sample for the invalid SQL scenario.");
        content.Controls.Add(_txtSqlQueryScenario2InvalidSql);
        content.Controls.Add(CreateInfoNote(
            "Expected result",
            "DataSyncer should log the SQL parser or execution error clearly, avoid any silent success path, and keep the service running.",
            WarningSoft,
            WarningColor));
        RefreshSqlQueryScenarioTemplates();
        return card;
    }

    private Control CreateSqlQueryScenario3Card()
    {
        var card = CreateCard("SQLQuery Scenario 3 - Wrong Database Type or Connection", "Generate a broken SQL Server connection sample and a reminder to try an unsupported DB type path if the job supports provider selection.", out var content);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate Broken Connection Sample", (_, _) => PrepareSqlQueryScenario3(), true));
        content.Controls.Add(actions);

        _txtSqlQueryScenario3BrokenConnection = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            Height = 94,
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 8, 0, 0),
            BackColor = Color.FromArgb(248, 250, 252)
        };
        _toolTip.SetToolTip(_txtSqlQueryScenario3BrokenConnection, "Broken connection example for the SQLQuery wrong-connection scenario.");
        content.Controls.Add(_txtSqlQueryScenario3BrokenConnection);
        content.Controls.Add(CreateInfoNote(
            "Expected result",
            "The SQLQuery job should fail gracefully, write a clear connection or provider error, and leave any unrelated DataSyncer jobs unaffected.",
            DangerSoft,
            DangerColor));
        RefreshSqlQueryScenarioTemplates();
        return card;
    }

    private Control CreateProgramAutomationCard()
    {
        var card = CreateCard("ProgramExecution Automation Timer", "Keep the script validation and output folders ready on a schedule. This is useful when DataSyncer will trigger program jobs repeatedly at fixed time gaps.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Automation mode", 0);
        _cmbProgramScheduleMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "ProgramExecution automation mode"
        };
        _cmbProgramScheduleMode.Items.AddRange(new object[]
        {
            "Run Once At Selected Time",
            "Repeat Daily At Selected Time",
            "Repeat Every N Minutes"
        });
        _cmbProgramScheduleMode.SelectedIndex = 0;
        _cmbProgramScheduleMode.SelectedIndexChanged += (_, _) => UpdateProgramAutomationControlState();
        grid.Controls.Add(_cmbProgramScheduleMode, 1, 0);

        AddLabelCell(grid, "Run at", 1);
        _dtProgramScheduleAt = CreateDateTimeInput(DateTime.Now.AddMinutes(5), "Date and time used for one-time and daily ProgramExecution preparation.");
        grid.Controls.Add(_dtProgramScheduleAt, 1, 1);

        AddLabelCell(grid, "Every N minutes", 2);
        _numProgramScheduleIntervalMinutes = CreateCsvNumeric(1, 1440, 1, "Interval in minutes for recurring ProgramExecution preparation.");
        _numProgramScheduleIntervalMinutes.AccessibleName = "ProgramExecution automation interval minutes";
        grid.Controls.Add(_numProgramScheduleIntervalMinutes, 1, 2);

        AddLabelCell(grid, "Timer target", 3);
        _cmbProgramScheduleTarget = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "ProgramExecution automation target"
        };
        _cmbProgramScheduleTarget.Items.AddRange(new object[]
        {
            "Scenario 1 - Successful Run Prep",
            "Scenario 2 - Invalid Path Check",
            "Scenario 3 - Argument Echo Prep",
            "All ProgramExecution Scenarios"
        });
        _cmbProgramScheduleTarget.SelectedIndex = 0;
        grid.Controls.Add(_cmbProgramScheduleTarget, 1, 3);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Arm Timer", (_, _) => ArmProgramSchedule(), true));
        actions.Controls.Add(CreateActionButton("Cancel Timer", (_, _) => CancelProgramSchedule()));
        actions.Controls.Add(CreateActionButton("Run Target Now", (_, _) => RunProgramScheduledTarget()));
        content.Controls.Add(actions);

        _lblProgramScheduleStatus = new Label
        {
            Text = "ProgramExecution scheduler idle. Choose a mode, then arm the timer.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblProgramScheduleStatus);

        _programScheduleTimer.Tick -= OnProgramScheduleTimerTick;
        _programScheduleTimer.Tick += OnProgramScheduleTimerTick;
        UpdateProgramAutomationControlState();
        return card;
    }

    private Control CreateProgramScenario1Card()
    {
        var card = CreateCard("ProgramExecution Scenario 1 - Successful Executable Run", "Validate the expected script, create the output folder, archive any prior output file, and use a helper script that writes to stdout for log capture checks.", out var content);
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
        actions.Controls.Add(CreateActionButton("Prepare Scenario 1", (_, _) => PrepareProgramScenario1(), true));
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
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "The heartbeat helper writes the same line to stdout and the optional output file, so DataSyncer can validate process start, captured standard output, and clean completion logging in a single run.",
            AccentSoft,
            AccentColor));
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
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use this invalid path to confirm that DataSyncer logs the launch failure clearly, keeps the service stable, and does not report a silent success.",
            WarningSoft,
            WarningColor));
        return card;
    }

    private Control CreateProgramScenario3Card()
    {
        var card = CreateCard("ProgramExecution Scenario 3 - Executable with Arguments", "Use a script that echoes the received arguments so DataSyncer logs can be compared directly with the configured command line.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Script path", 0);
        _txtProgramS3ScriptPath = CreateCsvTextBox(@"TestResources\TwoDaySoak\program_execution_echo_args.ps1", "Relative or absolute path to the argument-echo test script.");
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
        actions.Controls.Add(CreateActionButton("Prepare Scenario 3", (_, _) => PrepareProgramScenario3(), true));
        actions.Controls.Add(CreateActionButton("Preview Full Command", (_, _) => PreviewProgramCommand(), true));
        actions.Controls.Add(CreateActionButton("Validate Script Path", (_, _) => ValidateProgramScriptPath(_txtProgramS3ScriptPath, UpdateProgramScenario3Status)));
        actions.Controls.Add(CreateActionButton("Create Output Folder", (_, _) => CreateProgramOutputFolder(_txtProgramS3OutputFile, UpdateProgramScenario3Status)));
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
        _lblProgramS3Status = new Label
        {
            Text = "Argument-echo scenario not prepared yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblProgramS3Status);
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "The argument-echo helper prints each received value, so DataSyncer logs can be checked directly against the configured command line and quoted arguments.",
            AccentSoft,
            AccentColor));
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
            MaximumSize = new Size(720, 0),
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
        var startId = Decimal.ToInt32(_numDbToDbS1SourceIdStart.Value) + GetDbToDbAutomationIdOffset();
        var count = Decimal.ToInt32(_numDbToDbS1RowCount.Value);
        for (var i = 0; i < count; i++)
        {
            rows.Add(new DbToDbSeedRow(
                startId + i,
                machineCodes[i % machineCodes.Count],
                _txtDbToDbS1WorkCenter.Text.Trim(),
                _txtDbToDbS1SyncFlag.Text.Trim(),
                _dtDbToDbS1EventTimeBase.Value.AddMinutes(i),
                "Scenario1 Row " + (i + 1).ToString(CultureInfo.InvariantCulture) + GetDbToDbAutomationSuffix()));
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToDbTablesExist(connection, sourceTable, destinationTable);

            using var transaction = connection.BeginTransaction();
            var minId = rows.Min(row => row.SourceId);
            var maxId = rows.Max(row => row.SourceId);
            var deletedSourceRows = DeleteDbToDbRowsByIdRange(connection, transaction, sourceTable, minId, maxId);
            var deletedDestinationRows = _chkDbToDbS1ClearDestination.Checked
                ? ExecuteNonQueryCount(connection, transaction, "DELETE FROM " + destinationTable)
                : DeleteDbToDbRowsByIdRange(connection, transaction, destinationTable, minId, maxId);
            if (_chkDbToDbS1ClearDestination.Checked)
            {
                _logService.LogInfo("DBtoDB Scenario 1 cleared all destination rows before seeding.");
            }

            InsertDbToDbRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            transaction.Commit();

            SetDbToDbStatus(
                "Scenario 1 ready: source rows seeded=" + rows.Count +
                ", replaced in source=" + deletedSourceRows +
                ", destination rows cleared=" + deletedDestinationRows +
                ". After the DBtoDB immediate job, expect " + rows.Count + " destination inserts and " + rows.Count + " source flag updates." +
                GetDbToDbAutomationStatusSuffix(),
                SuccessColor);
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
        var startId = 200011 + GetDbToDbAutomationIdOffset();
        for (var i = 0; i < 5; i++)
        {
            sourceRows.Add(new DbToDbSeedRow(
                startId + i,
                "MC-0" + ((i % 5) + 1).ToString(CultureInfo.InvariantCulture),
                "LINE-A",
                _settings.SyncFlagNotProcessed,
                new DateTime(2026, 4, 6, 11, 30, 0).AddMinutes(i),
                "Scenario2 Fresh Row " + (i + 1).ToString(CultureInfo.InvariantCulture) + GetDbToDbAutomationSuffix()));
        }

        var staleCount = Decimal.ToInt32(_numDbToDbS2PreInsertCount.Value);
        var staleRows = sourceRows.Take(staleCount)
            .Select((row, index) => row with
            {
                SyncFlag = _settings.SyncFlagProcessed,
                Notes = "Scenario2 Stale Row " + (index + 1).ToString(CultureInfo.InvariantCulture) + GetDbToDbAutomationSuffix()
            })
            .ToList();

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToDbTablesExist(connection, sourceTable, destinationTable);

            using var transaction = connection.BeginTransaction();
            var minId = sourceRows.Min(row => row.SourceId);
            var maxId = sourceRows.Max(row => row.SourceId);
            var deletedSourceRows = DeleteDbToDbRowsByIdRange(connection, transaction, sourceTable, minId, maxId);
            var deletedDestinationRows = DeleteDbToDbRowsByIdRange(connection, transaction, destinationTable, minId, maxId);
            InsertDbToDbRows(connection, transaction, sourceTable, sourceRows, skipDuplicates: true);
            if (staleRows.Count > 0)
            {
                InsertDbToDbRows(connection, transaction, destinationTable, staleRows, skipDuplicates: true);
            }

            transaction.Commit();

            UpdateDbToDbConflictNote();
            SetDbToDbStatus(
                "Scenario 2 ready: source rows seeded=" + sourceRows.Count +
                ", duplicate destination rows seeded=" + staleRows.Count +
                ", replaced in source=" + deletedSourceRows +
                ", cleared from destination=" + deletedDestinationRows +
                ". Conflict mode reminder: " + _cmbDbToDbS2ConflictBehavior.SelectedItem +
                ". After the job, duplicates should be handled safely and source flags should follow the configured path." +
                GetDbToDbAutomationStatusSuffix(),
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

        UpdateDbToDbScenario3Progress(5, "Scenario 3: preparing source and destination setup...", AccentColor);

        var rows = new List<DbToDbSeedRow>();
        var currentId = 200016 + GetDbToDbAutomationIdOffset();
        for (var i = 0; i < Decimal.ToInt32(_numDbToDbS3InsideRangeCount.Value); i++)
        {
            rows.Add(new DbToDbSeedRow(currentId++, "MC-0" + ((i % 5) + 1), workCenters[0], _settings.SyncFlagNotProcessed, _dtDbToDbS3InsideDate.Value.AddMinutes(i), "Inside range" + GetDbToDbAutomationSuffix()));
        }

        for (var i = 0; i < Decimal.ToInt32(_numDbToDbS3OutsideRangeCount.Value); i++)
        {
            rows.Add(new DbToDbSeedRow(currentId++, "MC-0" + ((i % 5) + 1), workCenters[0], _settings.SyncFlagNotProcessed, _dtDbToDbS3OutsideDate.Value.AddMinutes(i), "Outside range" + GetDbToDbAutomationSuffix()));
        }

        for (var i = 0; i < 5; i++)
        {
            var workCenter = workCenters[Math.Min(i % workCenters.Count, workCenters.Count - 1)];
            var syncFlag = i == 4 ? _settings.SyncFlagProcessed : _settings.SyncFlagNotProcessed;
            var eventTime = i < 3 ? _dtDbToDbS3InsideDate.Value.AddHours(1).AddMinutes(i) : _dtDbToDbS3OutsideDate.Value.AddHours(1).AddMinutes(i);
            rows.Add(new DbToDbSeedRow(currentId++, "MC-MIX-" + (i + 1).ToString(CultureInfo.InvariantCulture), workCenter, syncFlag, eventTime, "Mixed condition" + GetDbToDbAutomationSuffix()));
        }

        UpdateDbToDbScenario3Progress(20, "Scenario 3: built inside-range, outside-range, and mixed-condition rows.", AccentColor);

        var insideRangeAnchor = _dtDbToDbS3InsideDate.Value;
        var staleRows = rows
            .Where(row =>
                row.EventTime >= insideRangeAnchor &&
                row.EventTime < insideRangeAnchor.AddHours(12) &&
                string.Equals(row.WorkCenter, workCenters[0], StringComparison.OrdinalIgnoreCase))
            .Take(Math.Min(5, rows.Count))
            .Select((row, index) => row with
            {
                SyncFlag = _settings.SyncFlagProcessed,
                Notes = "Scenario3 Stale Destination " + (index + 1).ToString(CultureInfo.InvariantCulture) + GetDbToDbAutomationSuffix()
            })
            .ToList();

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();
            EnsureDbToDbTablesExist(connection, sourceTable, destinationTable);
            UpdateDbToDbScenario3Progress(35, "Scenario 3: verified source and destination tables.", AccentColor);

            using var transaction = connection.BeginTransaction();
            var minId = rows.Min(row => row.SourceId);
            var maxId = rows.Max(row => row.SourceId);
            var deletedSourceRows = DeleteDbToDbRowsByIdRange(connection, transaction, sourceTable, minId, maxId);
            var deletedDestinationRows = DeleteDbToDbRowsByIdRange(connection, transaction, destinationTable, minId, maxId);
            UpdateDbToDbScenario3Progress(60, "Scenario 3: cleared prior rows from the same ID window.", AccentColor);

            InsertDbToDbRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            UpdateDbToDbScenario3Progress(80, "Scenario 3: inserted filtered source data.", AccentColor);
            if (_chkDbToDbS3PreloadDestination.Checked)
            {
                InsertDbToDbRows(connection, transaction, destinationTable, staleRows, skipDuplicates: true);
                UpdateDbToDbScenario3Progress(95, "Scenario 3: preloaded matching stale destination rows for delete-and-reinsert validation.", AccentColor);
            }
            else
            {
                UpdateDbToDbScenario3Progress(95, "Scenario 3: destination preload skipped; source-only setup is ready.", AccentColor);
            }

            transaction.Commit();
            UpdateDbToDbScenario3Progress(100, "Scenario 3 setup completed.", SuccessColor);

            SetDbToDbStatus(
                "Scenario 3 ready: source rows seeded=" + rows.Count +
                ", replaced in source=" + deletedSourceRows +
                ", cleared from destination=" + deletedDestinationRows +
                (_chkDbToDbS3PreloadDestination.Checked
                    ? ", stale destination rows seeded=" + staleRows.Count + "."
                    : ", destination preload skipped.") +
                " Use the selected inside-range date and first WorkCenter filter value as the manual copy slice. Progress updates are shown in this card while setup runs." +
                GetDbToDbAutomationStatusSuffix(),
                SuccessColor);
        }
        catch (Exception ex)
        {
            UpdateDbToDbScenario3Progress(0, "Scenario 3 progress reset after an error.", DangerColor);
            HandleDbActionError("DBtoDB", ex);
        }
    }

    private void OnDbToDbScheduleTimerTick(object? sender, EventArgs e)
    {
        HandleDbToDbScheduleTick();
    }

    private void ArmDbToDbSchedule()
    {
        if (_dtDbToDbScheduleAt is null ||
            _cmbDbToDbScheduleMode is null ||
            _numDbToDbScheduleIntervalMinutes is null ||
            _cmbDbToDbScheduleTarget is null)
        {
            return;
        }

        var mode = GetSelectedDbToDbScheduleMode();
        var scheduledAt = mode switch
        {
            DbToDbScheduleMode.Daily => ResolveNextDbToDbScheduledRun(_dtDbToDbScheduleAt.Value),
            DbToDbScheduleMode.EveryNMinutes => DateTime.Now.AddMinutes(decimal.ToDouble(_numDbToDbScheduleIntervalMinutes.Value)),
            _ => _dtDbToDbScheduleAt.Value
        };

        if (mode == DbToDbScheduleMode.OneTime && scheduledAt <= DateTime.Now)
        {
            MessageBox.Show(
                "Pick a future time for the DBtoDB scheduler, or switch to a recurring mode to let the app calculate the next run automatically.",
                "DBtoDB Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _armedDbToDbScheduleMode = mode;
        _dbToDbScheduleInterval = mode == DbToDbScheduleMode.EveryNMinutes
            ? TimeSpan.FromMinutes(decimal.ToDouble(_numDbToDbScheduleIntervalMinutes.Value))
            : null;
        _dbToDbScheduledRunAt = scheduledAt;
        _dbToDbScheduleTimer.Start();
        UpdateDbToDbScheduleStatus();

        _logService.LogInfo(
            "DBtoDB scheduler armed for " +
            scheduledAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetDbToDbScheduleTargetLabel() +
            " | " + GetArmedDbToDbScheduleModeLabel());
        RefreshLogViewer();
        WriteStatus("DBtoDB scheduler armed");
    }

    private void CancelDbToDbSchedule()
    {
        _dbToDbScheduleTimer.Stop();
        _dbToDbScheduledRunAt = null;
        _dbToDbScheduleInterval = null;
        _armedDbToDbScheduleMode = DbToDbScheduleMode.OneTime;
        UpdateDbToDbScheduleStatus();
        _logService.LogInfo("DBtoDB scheduler canceled");
        RefreshLogViewer();
        WriteStatus("DBtoDB scheduler canceled");
    }

    private void HandleDbToDbScheduleTick()
    {
        if (!_dbToDbScheduledRunAt.HasValue)
        {
            _dbToDbScheduleTimer.Stop();
            UpdateDbToDbScheduleStatus();
            return;
        }

        if (DateTime.Now < _dbToDbScheduledRunAt.Value)
        {
            UpdateDbToDbScheduleStatus();
            return;
        }

        RunDbToDbScheduledTarget();

        if (_armedDbToDbScheduleMode == DbToDbScheduleMode.Daily)
        {
            _dbToDbScheduledRunAt = ResolveNextDbToDbScheduledRun(_dbToDbScheduledRunAt.Value.AddDays(1));
            _dbToDbScheduleTimer.Start();
        }
        else if (_armedDbToDbScheduleMode == DbToDbScheduleMode.EveryNMinutes && _dbToDbScheduleInterval.HasValue)
        {
            _dbToDbScheduledRunAt = ResolveNextDbToDbRecurringRun(_dbToDbScheduledRunAt.Value, _dbToDbScheduleInterval.Value);
            _dbToDbScheduleTimer.Start();
        }
        else
        {
            _dbToDbScheduleTimer.Stop();
            _dbToDbScheduledRunAt = null;
            _dbToDbScheduleInterval = null;
        }

        UpdateDbToDbScheduleStatus();
    }

    private void RunDbToDbScheduledTarget()
    {
        _dbToDbAutomationExecutionContext = CreateDbToDbAutomationExecutionContext();
        try
        {
            var targetLabel = GetDbToDbScheduleTargetLabel();
            switch (_cmbDbToDbScheduleTarget?.SelectedIndex)
            {
                case 1:
                    GenerateDbToDbScenario2();
                    break;
                case 2:
                    GenerateDbToDbScenario3();
                    break;
                case 3:
                    GenerateDbToDbScenario1();
                    GenerateDbToDbScenario2();
                    GenerateDbToDbScenario3();
                    break;
                default:
                    GenerateDbToDbScenario1();
                    break;
            }

            _logService.LogSuccess("DBtoDB scheduler executed: " + targetLabel);
            RefreshLogViewer();
            WriteStatus("DBtoDB scheduler executed");
        }
        catch (Exception ex)
        {
            _logService.LogError("DBtoDB scheduler failed: " + ex.Message);
            RefreshLogViewer();
            WriteStatus("DBtoDB scheduler failed");

            MessageBox.Show(
                "Scheduled DBtoDB generation failed." + Environment.NewLine + Environment.NewLine + ex.Message,
                "DBtoDB Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            _dbToDbAutomationExecutionContext = null;
        }
    }

    private void UpdateDbToDbScheduleStatus()
    {
        if (_lblDbToDbScheduleStatus is null)
        {
            return;
        }

        if (!_dbToDbScheduledRunAt.HasValue)
        {
            _lblDbToDbScheduleStatus.Text = "DBtoDB scheduler idle. Choose a mode, then arm the timer.";
            _lblDbToDbScheduleStatus.ForeColor = MutedText;
            return;
        }

        var remaining = _dbToDbScheduledRunAt.Value - DateTime.Now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        _lblDbToDbScheduleStatus.Text =
            "Armed for " + _dbToDbScheduledRunAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetDbToDbScheduleTargetLabel() +
            " | " + GetArmedDbToDbScheduleModeLabel() +
            Environment.NewLine +
            "Time remaining: " + remaining.ToString(@"dd\.hh\:mm\:ss", CultureInfo.InvariantCulture);
        _lblDbToDbScheduleStatus.ForeColor = AccentColor;
    }

    private void UpdateDbToDbAutomationControlState()
    {
        var mode = GetSelectedDbToDbScheduleMode();
        if (_dtDbToDbScheduleAt is not null)
        {
            _dtDbToDbScheduleAt.Enabled = mode != DbToDbScheduleMode.EveryNMinutes;
        }

        if (_numDbToDbScheduleIntervalMinutes is not null)
        {
            _numDbToDbScheduleIntervalMinutes.Enabled = mode == DbToDbScheduleMode.EveryNMinutes;
        }
    }

    private DbToDbScheduleMode GetSelectedDbToDbScheduleMode()
    {
        return _cmbDbToDbScheduleMode?.SelectedIndex switch
        {
            1 => DbToDbScheduleMode.Daily,
            2 => DbToDbScheduleMode.EveryNMinutes,
            _ => DbToDbScheduleMode.OneTime
        };
    }

    private string GetDbToDbScheduleTargetLabel()
    {
        return _cmbDbToDbScheduleTarget?.SelectedIndex switch
        {
            1 => "Generate Scenario 2",
            2 => "Generate Scenario 3",
            3 => "Generate All DBtoDB Scenarios",
            _ => "Generate Scenario 1"
        };
    }

    private string GetArmedDbToDbScheduleModeLabel()
    {
        return _armedDbToDbScheduleMode switch
        {
            DbToDbScheduleMode.Daily => "repeats daily",
            DbToDbScheduleMode.EveryNMinutes when _dbToDbScheduleInterval.HasValue =>
                "every " + FormatDbToDbIntervalLabel(_dbToDbScheduleInterval.Value),
            _ => "one-time"
        };
    }

    private static DateTime ResolveNextDbToDbScheduledRun(DateTime requestedTime)
    {
        while (requestedTime <= DateTime.Now)
        {
            requestedTime = requestedTime.AddDays(1);
        }

        return requestedTime;
    }

    private static DateTime ResolveNextDbToDbRecurringRun(DateTime previousScheduledAt, TimeSpan interval)
    {
        if (interval <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(interval), "Recurring DBtoDB interval must be greater than zero.");
        }

        var nextRun = previousScheduledAt.Add(interval);
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.Add(interval);
        }

        return nextRun;
    }

    private DbToDbAutomationExecutionContext CreateDbToDbAutomationExecutionContext()
    {
        _dbToDbAutomationRunSequence++;
        if (_dbToDbAutomationRunSequence <= 0)
        {
            _dbToDbAutomationRunSequence = 1;
        }

        return new DbToDbAutomationExecutionContext(DateTime.Now, _dbToDbAutomationRunSequence);
    }

    private static string FormatDbToDbIntervalLabel(TimeSpan interval)
    {
        var totalMinutes = Math.Max(1, (int)Math.Round(interval.TotalMinutes, MidpointRounding.AwayFromZero));
        return totalMinutes.ToString(CultureInfo.InvariantCulture) + (totalMinutes == 1 ? " minute" : " minutes");
    }

    private void UpdateDbToDbScenario3Progress(int value, string message, Color color)
    {
        if (_progressDbToDbScenario3 is not null)
        {
            _progressDbToDbScenario3.Value = Math.Max(_progressDbToDbScenario3.Minimum, Math.Min(_progressDbToDbScenario3.Maximum, value));
        }

        if (_lblDbToDbScenario3Progress is not null)
        {
            _lblDbToDbScenario3Progress.Text = message;
            _lblDbToDbScenario3Progress.ForeColor = color;
        }

        Application.DoEvents();
    }

    private int GetDbToDbAutomationIdOffset()
    {
        return _dbToDbAutomationExecutionContext is null
            ? 0
            : _dbToDbAutomationExecutionContext.RunSequence * 1000;
    }

    private string GetDbToDbAutomationSuffix()
    {
        if (_dbToDbAutomationExecutionContext is null)
        {
            return string.Empty;
        }

        return " | batch " +
               _dbToDbAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
               "_r" +
               _dbToDbAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture);
    }

    private string GetDbToDbAutomationStatusSuffix()
    {
        return _dbToDbAutomationExecutionContext is null
            ? string.Empty
            : " Generated by scheduler batch " +
              _dbToDbAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
              "_r" +
              _dbToDbAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture) +
              ".";
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

        if (EnsureDbToJsonOutputFolder() is null)
        {
            return;
        }

        var rows = new List<DbToJsonSeedRow>();
        var startId = _settings.DbToJsonIdStart + GetDbToJsonAutomationIdOffset();
        var baseTime = GetDbToJsonScenario1BaseTime();
        var rowCount = Decimal.ToInt32(_numDbToJsonS1RowCount.Value);
        for (var i = 0; i < rowCount; i++)
        {
            var exportId = startId + i;
            rows.Add(new DbToJsonSeedRow(
                exportId,
                _txtDbToJsonS1ExportFlag.Text.Trim(),
                deviceCodes[i % deviceCodes.Count],
                resultCodes[i % resultCodes.Count],
                baseTime.AddMinutes(i),
                BuildDbToJsonPayload(exportId, false),
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
            var deletedRows = DeleteDbToJsonRowsByIdRange(connection, transaction, sourceTable, startId, startId + rowCount - 1);
            InsertDbToJsonRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            transaction.Commit();

            SetDbToJsonStatus(
                "Scenario 1 ready: export rows seeded=" + rows.Count +
                ", replaced in source=" + deletedRows +
                ". After the DBtoJSON folder export job runs, expect one JSON file in " + GetDbToJsonOutputFolderDisplay() +
                ", " + rows.Count + " exported rows, processed flags, and an updated state file." +
                GetDbToJsonAutomationStatusSuffix(),
                SuccessColor);
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

        PingDbToJsonEndpoint();

        var rows = new List<DbToJsonSeedRow>();
        var startId = _settings.DbToJsonIdStart + 100 + GetDbToJsonAutomationIdOffset();
        var baseTime = GetDbToJsonScenario2BaseTime();
        var rowCount = Decimal.ToInt32(_numDbToJsonS2RowCount.Value);
        for (var i = 0; i < rowCount; i++)
        {
            var exportId = startId + i;
            rows.Add(new DbToJsonSeedRow(
                exportId,
                _settings.SyncFlagNotProcessed,
                "DEV-0" + ((i % 2) + 1).ToString(CultureInfo.InvariantCulture),
                i % 2 == 0 ? "READY" : "PASS",
                baseTime.AddMinutes(i),
                BuildDbToJsonPayload(exportId, true),
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
            ExecuteNonQuery(connection, BuildDbToJsonResponseColumnsSql(sourceTable));

            using var transaction = connection.BeginTransaction();
            var deletedRows = DeleteDbToJsonRowsByIdRange(connection, transaction, sourceTable, startId, startId + rowCount - 1);
            InsertDbToJsonRows(connection, transaction, sourceTable, rows, skipDuplicates: true);
            transaction.Commit();

            SetDbToJsonStatus(
                "Scenario 2 ready: API export rows seeded=" + rows.Count +
                ", replaced in source=" + deletedRows +
                ". Response-writeback columns were verified automatically. After the DBtoJSON API job runs, expect API success logs, mapped response values written to the source table, and processed flags updated only after API completion." +
                GetDbToJsonAutomationStatusSuffix(),
                SuccessColor);
        }
        catch (Exception ex)
        {
            HandleDbActionError("DBtoJSON", ex);
        }
    }

    private void PrepareDbToJsonScenario3()
    {
        MarkDbToJsonRowsProcessed();
        RefreshDbToJsonUnprocessedCount();

        SetDbToJsonStatus(
            "Scenario 3 ready: all DBtoJSON rows are marked processed so the next DataSyncer run should report no new records and produce no extra JSON file or API send." +
            GetDbToJsonAutomationStatusSuffix(),
            SuccessColor);
    }

    private void OnDbToJsonScheduleTimerTick(object? sender, EventArgs e)
    {
        HandleDbToJsonScheduleTick();
    }

    private void ArmDbToJsonSchedule()
    {
        if (_dtDbToJsonScheduleAt is null ||
            _cmbDbToJsonScheduleMode is null ||
            _numDbToJsonScheduleIntervalMinutes is null ||
            _cmbDbToJsonScheduleTarget is null)
        {
            return;
        }

        var mode = GetSelectedDbToJsonScheduleMode();
        var scheduledAt = mode switch
        {
            DbToJsonScheduleMode.Daily => ResolveNextDbToJsonScheduledRun(_dtDbToJsonScheduleAt.Value),
            DbToJsonScheduleMode.EveryNMinutes => DateTime.Now.AddMinutes(decimal.ToDouble(_numDbToJsonScheduleIntervalMinutes.Value)),
            _ => _dtDbToJsonScheduleAt.Value
        };

        if (mode == DbToJsonScheduleMode.OneTime && scheduledAt <= DateTime.Now)
        {
            MessageBox.Show(
                "Pick a future time for the DBtoJSON scheduler, or switch to a recurring mode to let the app calculate the next run automatically.",
                "DBtoJSON Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _armedDbToJsonScheduleMode = mode;
        _dbToJsonScheduleInterval = mode == DbToJsonScheduleMode.EveryNMinutes
            ? TimeSpan.FromMinutes(decimal.ToDouble(_numDbToJsonScheduleIntervalMinutes.Value))
            : null;
        _dbToJsonScheduledRunAt = scheduledAt;
        _dbToJsonScheduleTimer.Start();
        UpdateDbToJsonScheduleStatus();

        _logService.LogInfo(
            "DBtoJSON scheduler armed for " +
            scheduledAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetDbToJsonScheduleTargetLabel() +
            " | " + GetArmedDbToJsonScheduleModeLabel());
        RefreshLogViewer();
        WriteStatus("DBtoJSON scheduler armed");
    }

    private void CancelDbToJsonSchedule()
    {
        _dbToJsonScheduleTimer.Stop();
        _dbToJsonScheduledRunAt = null;
        _dbToJsonScheduleInterval = null;
        _armedDbToJsonScheduleMode = DbToJsonScheduleMode.OneTime;
        UpdateDbToJsonScheduleStatus();
        _logService.LogInfo("DBtoJSON scheduler canceled");
        RefreshLogViewer();
        WriteStatus("DBtoJSON scheduler canceled");
    }

    private void HandleDbToJsonScheduleTick()
    {
        if (!_dbToJsonScheduledRunAt.HasValue)
        {
            _dbToJsonScheduleTimer.Stop();
            UpdateDbToJsonScheduleStatus();
            return;
        }

        if (DateTime.Now < _dbToJsonScheduledRunAt.Value)
        {
            UpdateDbToJsonScheduleStatus();
            return;
        }

        RunDbToJsonScheduledTarget();

        if (_armedDbToJsonScheduleMode == DbToJsonScheduleMode.Daily)
        {
            _dbToJsonScheduledRunAt = ResolveNextDbToJsonScheduledRun(_dbToJsonScheduledRunAt.Value.AddDays(1));
            _dbToJsonScheduleTimer.Start();
        }
        else if (_armedDbToJsonScheduleMode == DbToJsonScheduleMode.EveryNMinutes && _dbToJsonScheduleInterval.HasValue)
        {
            _dbToJsonScheduledRunAt = ResolveNextDbToJsonRecurringRun(_dbToJsonScheduledRunAt.Value, _dbToJsonScheduleInterval.Value);
            _dbToJsonScheduleTimer.Start();
        }
        else
        {
            _dbToJsonScheduleTimer.Stop();
            _dbToJsonScheduledRunAt = null;
            _dbToJsonScheduleInterval = null;
        }

        UpdateDbToJsonScheduleStatus();
    }

    private void RunDbToJsonScheduledTarget()
    {
        _dbToJsonAutomationExecutionContext = CreateDbToJsonAutomationExecutionContext();
        try
        {
            var targetLabel = GetDbToJsonScheduleTargetLabel();
            switch (_cmbDbToJsonScheduleTarget?.SelectedIndex)
            {
                case 1:
                    GenerateDbToJsonScenario2();
                    break;
                case 2:
                    PrepareDbToJsonScenario3();
                    break;
                case 3:
                    GenerateDbToJsonScenario1();
                    GenerateDbToJsonScenario2();
                    break;
                default:
                    GenerateDbToJsonScenario1();
                    break;
            }

            _logService.LogSuccess("DBtoJSON scheduler executed: " + targetLabel);
            RefreshLogViewer();
            WriteStatus("DBtoJSON scheduler executed");
        }
        catch (Exception ex)
        {
            _logService.LogError("DBtoJSON scheduler failed: " + ex.Message);
            RefreshLogViewer();
            WriteStatus("DBtoJSON scheduler failed");

            MessageBox.Show(
                "Scheduled DBtoJSON generation failed." + Environment.NewLine + Environment.NewLine + ex.Message,
                "DBtoJSON Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            _dbToJsonAutomationExecutionContext = null;
        }
    }

    private void UpdateDbToJsonScheduleStatus()
    {
        if (_lblDbToJsonScheduleStatus is null)
        {
            return;
        }

        if (!_dbToJsonScheduledRunAt.HasValue)
        {
            _lblDbToJsonScheduleStatus.Text = "DBtoJSON scheduler idle. Choose a mode, then arm the timer.";
            _lblDbToJsonScheduleStatus.ForeColor = MutedText;
            return;
        }

        var remaining = _dbToJsonScheduledRunAt.Value - DateTime.Now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        _lblDbToJsonScheduleStatus.Text =
            "Armed for " + _dbToJsonScheduledRunAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetDbToJsonScheduleTargetLabel() +
            " | " + GetArmedDbToJsonScheduleModeLabel() +
            Environment.NewLine +
            "Time remaining: " + remaining.ToString(@"dd\.hh\:mm\:ss", CultureInfo.InvariantCulture);
        _lblDbToJsonScheduleStatus.ForeColor = AccentColor;
    }

    private void UpdateDbToJsonAutomationControlState()
    {
        var mode = GetSelectedDbToJsonScheduleMode();
        if (_dtDbToJsonScheduleAt is not null)
        {
            _dtDbToJsonScheduleAt.Enabled = mode != DbToJsonScheduleMode.EveryNMinutes;
        }

        if (_numDbToJsonScheduleIntervalMinutes is not null)
        {
            _numDbToJsonScheduleIntervalMinutes.Enabled = mode == DbToJsonScheduleMode.EveryNMinutes;
        }
    }

    private DbToJsonScheduleMode GetSelectedDbToJsonScheduleMode()
    {
        return _cmbDbToJsonScheduleMode?.SelectedIndex switch
        {
            1 => DbToJsonScheduleMode.Daily,
            2 => DbToJsonScheduleMode.EveryNMinutes,
            _ => DbToJsonScheduleMode.OneTime
        };
    }

    private string GetDbToJsonScheduleTargetLabel()
    {
        return _cmbDbToJsonScheduleTarget?.SelectedIndex switch
        {
            1 => "Generate Scenario 2",
            2 => "Prepare Scenario 3 Re-run State",
            3 => "Generate Scenarios 1 + 2",
            _ => "Generate Scenario 1"
        };
    }

    private string GetArmedDbToJsonScheduleModeLabel()
    {
        return _armedDbToJsonScheduleMode switch
        {
            DbToJsonScheduleMode.Daily => "repeats daily",
            DbToJsonScheduleMode.EveryNMinutes when _dbToJsonScheduleInterval.HasValue =>
                "every " + FormatDbToJsonIntervalLabel(_dbToJsonScheduleInterval.Value),
            _ => "one-time"
        };
    }

    private static DateTime ResolveNextDbToJsonScheduledRun(DateTime requestedTime)
    {
        while (requestedTime <= DateTime.Now)
        {
            requestedTime = requestedTime.AddDays(1);
        }

        return requestedTime;
    }

    private static DateTime ResolveNextDbToJsonRecurringRun(DateTime previousScheduledAt, TimeSpan interval)
    {
        if (interval <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(interval), "Recurring DBtoJSON interval must be greater than zero.");
        }

        var nextRun = previousScheduledAt.Add(interval);
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.Add(interval);
        }

        return nextRun;
    }

    private DbToJsonAutomationExecutionContext CreateDbToJsonAutomationExecutionContext()
    {
        _dbToJsonAutomationRunSequence++;
        if (_dbToJsonAutomationRunSequence <= 0)
        {
            _dbToJsonAutomationRunSequence = 1;
        }

        return new DbToJsonAutomationExecutionContext(DateTime.Now, _dbToJsonAutomationRunSequence);
    }

    private static string FormatDbToJsonIntervalLabel(TimeSpan interval)
    {
        var totalMinutes = Math.Max(1, (int)Math.Round(interval.TotalMinutes, MidpointRounding.AwayFromZero));
        return totalMinutes.ToString(CultureInfo.InvariantCulture) + (totalMinutes == 1 ? " minute" : " minutes");
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
            var baseRowId = _settings.SqlQueryIdStart + GetSqlQueryAutomationIdOffset();
            var maximumRowId = baseRowId + total - 1;
            var deleted = DeleteSqlQueryRowsByIdRange(connection, transaction, tableName, baseRowId, maximumRowId);
            var baseTime = GetSqlQueryScenario1BaseTime();
            for (var i = 0; i < total; i++)
            {
                using var command = new SqlCommand(
                    "INSERT INTO " + tableName + " (RowId, Status, WorkCenter, EventTime, UpdatedByJob, UpdatedAt, Notes) VALUES (@id, @status, @workCenter, @eventTime, NULL, NULL, @notes)",
                    connection,
                    transaction);
                command.Parameters.AddWithValue("@id", baseRowId + i);
                command.Parameters.AddWithValue("@status", i < pending ? "PENDING" : "DONE");
                command.Parameters.AddWithValue("@workCenter", i < pending ? "LINE-A" : "LINE-B");
                command.Parameters.AddWithValue("@eventTime", baseTime.AddMinutes(i));
                command.Parameters.AddWithValue("@notes", "SQLQuery Seed Row " + (i + 1).ToString(CultureInfo.InvariantCulture) + GetSqlQueryAutomationStatusSuffix());
                command.ExecuteNonQuery();
            }

            transaction.Commit();
            RefreshSqlQueryScenarioTemplates();
            SetSqlQueryStatus(
                "Scenario 1 prepared " + total + " rows (" + pending + " PENDING, " + done + " DONE), cleared " + deleted +
                " prior row(s) in RowId range " + baseRowId + "-" + maximumRowId + ".",
                SuccessColor);
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

    private void RefreshSqlQueryScenarioTemplates()
    {
        if (_txtSqlQueryScenario1UpdateSql is not null)
        {
            _txtSqlQueryScenario1UpdateSql.Text = BuildSqlQueryScenario1UpdateSql();
        }

        if (_txtSqlQueryScenario2InvalidSql is not null)
        {
            _txtSqlQueryScenario2InvalidSql.Text = BuildSqlQueryScenario2InvalidSql();
        }

        if (_txtSqlQueryScenario3BrokenConnection is not null)
        {
            _txtSqlQueryScenario3BrokenConnection.Text = BuildSqlQueryBrokenConnectionSample();
        }
    }

    private void PrepareSqlQueryScenario2()
    {
        RefreshSqlQueryScenarioTemplates();
        SetSqlQueryStatus("Prepared malformed SQL sample for Scenario 2.", WarningColor);
    }

    private void PrepareSqlQueryScenario3()
    {
        RefreshSqlQueryScenarioTemplates();
        SetSqlQueryStatus("Prepared broken-connection guidance for Scenario 3.", WarningColor);
    }

    private string BuildSqlQueryScenario1UpdateSql()
    {
        var tableName = ResolveSqlQueryDisplayTableName();
        var pending = Decimal.ToInt32(_numSqlQueryPendingRows?.Value ?? 0);
        var total = Decimal.ToInt32(_numSqlQueryTotalRows?.Value ?? 0);
        var baseRowId = _settings.SqlQueryIdStart + GetSqlQueryAutomationIdOffset();
        var maxPendingRowId = pending > 0 ? baseRowId + pending - 1 : baseRowId;
        var expectedRows = pending.ToString(CultureInfo.InvariantCulture);
        var baseTime = GetSqlQueryScenario1BaseTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        var finalRowId = total > 0 ? baseRowId + total - 1 : baseRowId;

        return
            "-- Expected affected rows: " + expectedRows + Environment.NewLine +
            "-- Seeded RowId range: " + baseRowId + " to " + finalRowId + Environment.NewLine +
            "-- Seeded EventTime base: " + baseTime + Environment.NewLine +
            "UPDATE " + tableName + Environment.NewLine +
            "SET Status = 'DONE'," + Environment.NewLine +
            "    UpdatedByJob = 'DataSyncer SQLQuery'," + Environment.NewLine +
            "    UpdatedAt = SYSUTCDATETIME()" + Environment.NewLine +
            "WHERE Status = 'PENDING'" + Environment.NewLine +
            "  AND WorkCenter = 'LINE-A'" + Environment.NewLine +
            "  AND RowId BETWEEN " + baseRowId + " AND " + maxPendingRowId + ";";
    }

    private string BuildSqlQueryScenario2InvalidSql()
    {
        var tableName = ResolveSqlQueryDisplayTableName();
        return
            "-- Intentionally malformed SQL for negative-path validation" + Environment.NewLine +
            "UPDTE " + tableName + Environment.NewLine +
            "SET Status = 'DONE'" + Environment.NewLine +
            "WHERE Status = 'PENDING';";
    }

    private string BuildSqlQueryBrokenConnectionSample()
    {
        var builder = new SqlConnectionStringBuilder();
        if (_cmbSqlQueryConnection?.SelectedItem is ConnectionChoice choice && !string.IsNullOrWhiteSpace(choice.ConnectionString))
        {
            try
            {
                builder.ConnectionString = choice.ConnectionString;
            }
            catch
            {
                builder.DataSource = "invalid-host";
                builder.InitialCatalog = "MissingCatalog_QA";
            }
        }

        if (string.IsNullOrWhiteSpace(builder.DataSource))
        {
            builder.DataSource = "invalid-host";
        }

        if (string.IsNullOrWhiteSpace(builder.InitialCatalog))
        {
            builder.InitialCatalog = "MissingCatalog_QA";
        }

        builder.DataSource = "invalid-host\\BROKEN";
        builder.InitialCatalog = "MissingCatalog_QA";
        builder.UserID = "broken_user";
        builder.Password = "broken_password";
        builder.IntegratedSecurity = false;
        builder.TrustServerCertificate = true;
        builder.ConnectTimeout = 3;

        return
            "-- Broken SQL Server connection sample" + Environment.NewLine +
            builder.ConnectionString + Environment.NewLine + Environment.NewLine +
            "-- If DataSyncer exposes a DB-type selector, also switch it away from SQL Server" + Environment.NewLine +
            "-- to verify unsupported-provider handling remains isolated to this job.";
    }

    private string ResolveSqlQueryDisplayTableName()
    {
        if (_txtSqlQueryTableName is null)
        {
            return "dbo.DS_SqlQuery_Test";
        }

        return TryNormalizeSqlTableName(_txtSqlQueryTableName.Text, out var tableName, out _, out _)
            ? tableName
            : _txtSqlQueryTableName.Text.Trim();
    }

    private DateTime GetSqlQueryScenario1BaseTime()
    {
        return _sqlQueryAutomationExecutionContext?.RunStamp ?? _dtSqlQueryEventTimeBase?.Value ?? new DateTime(2026, 4, 6, 16, 0, 0);
    }

    private int GetSqlQueryAutomationIdOffset()
    {
        return _sqlQueryAutomationExecutionContext is null
            ? 0
            : _sqlQueryAutomationExecutionContext.RunSequence * 1000;
    }

    private string GetSqlQueryAutomationStatusSuffix()
    {
        return _sqlQueryAutomationExecutionContext is null
            ? string.Empty
            : " Scheduler batch " +
              _sqlQueryAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
              "_r" +
              _sqlQueryAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture) +
              ".";
    }

    private void OnSqlQueryScheduleTimerTick(object? sender, EventArgs e)
    {
        HandleSqlQueryScheduleTick();
    }

    private void ArmSqlQuerySchedule()
    {
        if (_dtSqlQueryScheduleAt is null ||
            _cmbSqlQueryScheduleMode is null ||
            _numSqlQueryScheduleIntervalMinutes is null)
        {
            return;
        }

        var mode = GetSelectedSqlQueryScheduleMode();
        var scheduledAt = mode switch
        {
            SqlQueryScheduleMode.Daily => ResolveNextSqlQueryScheduledRun(_dtSqlQueryScheduleAt.Value),
            SqlQueryScheduleMode.EveryNMinutes => DateTime.Now.AddMinutes(decimal.ToDouble(_numSqlQueryScheduleIntervalMinutes.Value)),
            _ => _dtSqlQueryScheduleAt.Value
        };

        if (mode == SqlQueryScheduleMode.OneTime && scheduledAt <= DateTime.Now)
        {
            MessageBox.Show(
                "Choose a time in the future for a one-time SQLQuery automation run.",
                "SQLQuery Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _armedSqlQueryScheduleMode = mode;
        _sqlQueryScheduleInterval = mode == SqlQueryScheduleMode.EveryNMinutes
            ? TimeSpan.FromMinutes(decimal.ToDouble(_numSqlQueryScheduleIntervalMinutes.Value))
            : null;
        _sqlQueryScheduledRunAt = scheduledAt;
        _sqlQueryScheduleTimer.Start();
        UpdateSqlQueryScheduleStatus();
        SetSqlQueryStatus(
            "Armed SQLQuery automation for " + scheduledAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetSqlQueryScheduleTargetLabel() +
            " | " + GetArmedSqlQueryScheduleModeLabel(),
            SuccessColor);
    }

    private void CancelSqlQuerySchedule()
    {
        _sqlQueryScheduleTimer.Stop();
        _sqlQueryScheduledRunAt = null;
        _sqlQueryScheduleInterval = null;
        _armedSqlQueryScheduleMode = SqlQueryScheduleMode.OneTime;
        UpdateSqlQueryScheduleStatus();
        SetSqlQueryStatus("Cancelled SQLQuery automation timer.", WarningColor);
    }

    private void HandleSqlQueryScheduleTick()
    {
        if (!_sqlQueryScheduledRunAt.HasValue)
        {
            _sqlQueryScheduleTimer.Stop();
            UpdateSqlQueryScheduleStatus();
            return;
        }

        if (DateTime.Now < _sqlQueryScheduledRunAt.Value)
        {
            UpdateSqlQueryScheduleStatus();
            return;
        }

        RunSqlQueryScheduledTarget();

        if (_armedSqlQueryScheduleMode == SqlQueryScheduleMode.Daily)
        {
            _sqlQueryScheduledRunAt = ResolveNextSqlQueryScheduledRun(_sqlQueryScheduledRunAt.Value.AddDays(1));
            _sqlQueryScheduleTimer.Start();
        }
        else if (_armedSqlQueryScheduleMode == SqlQueryScheduleMode.EveryNMinutes && _sqlQueryScheduleInterval.HasValue)
        {
            _sqlQueryScheduledRunAt = ResolveNextSqlQueryRecurringRun(_sqlQueryScheduledRunAt.Value, _sqlQueryScheduleInterval.Value);
            _sqlQueryScheduleTimer.Start();
        }
        else
        {
            _sqlQueryScheduleTimer.Stop();
            _sqlQueryScheduledRunAt = null;
            _sqlQueryScheduleInterval = null;
        }

        UpdateSqlQueryScheduleStatus();
    }

    private void RunSqlQueryScheduledTarget()
    {
        try
        {
            _sqlQueryAutomationExecutionContext = new SqlQueryAutomationExecutionContext(DateTime.Now, ++_sqlQueryAutomationRunSequence);
            switch (_cmbSqlQueryScheduleTarget?.SelectedIndex)
            {
                case 1:
                    PrepareSqlQueryScenario2();
                    break;
                case 2:
                    PrepareSqlQueryScenario3();
                    break;
                case 3:
                    GenerateSqlQueryScenario1();
                    PrepareSqlQueryScenario2();
                    PrepareSqlQueryScenario3();
                    break;
                default:
                    GenerateSqlQueryScenario1();
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Scheduled SQLQuery preparation failed." + Environment.NewLine + Environment.NewLine + ex.Message,
                "SQLQuery Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            SetSqlQueryStatus("Scheduled SQLQuery preparation failed: " + ex.Message, DangerColor);
        }
        finally
        {
            _sqlQueryAutomationExecutionContext = null;
            RefreshSqlQueryScenarioTemplates();
        }
    }

    private void UpdateSqlQueryScheduleStatus()
    {
        if (_lblSqlQueryScheduleStatus is null)
        {
            return;
        }

        if (!_sqlQueryScheduledRunAt.HasValue)
        {
            _lblSqlQueryScheduleStatus.Text = "SQLQuery scheduler idle. Choose a mode, then arm the timer.";
            _lblSqlQueryScheduleStatus.ForeColor = MutedText;
            return;
        }

        var remaining = _sqlQueryScheduledRunAt.Value - DateTime.Now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        _lblSqlQueryScheduleStatus.Text =
            "Armed for " + _sqlQueryScheduledRunAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetSqlQueryScheduleTargetLabel() +
            " | " + GetArmedSqlQueryScheduleModeLabel() +
            " | remaining " + remaining.ToString(@"dd\.hh\:mm\:ss", CultureInfo.InvariantCulture);
        _lblSqlQueryScheduleStatus.ForeColor = AccentColor;
    }

    private void UpdateSqlQueryAutomationControlState()
    {
        var mode = GetSelectedSqlQueryScheduleMode();
        if (_dtSqlQueryScheduleAt is not null)
        {
            _dtSqlQueryScheduleAt.Enabled = mode != SqlQueryScheduleMode.EveryNMinutes;
        }

        if (_numSqlQueryScheduleIntervalMinutes is not null)
        {
            _numSqlQueryScheduleIntervalMinutes.Enabled = mode == SqlQueryScheduleMode.EveryNMinutes;
        }
    }

    private SqlQueryScheduleMode GetSelectedSqlQueryScheduleMode()
    {
        return _cmbSqlQueryScheduleMode?.SelectedIndex switch
        {
            1 => SqlQueryScheduleMode.Daily,
            2 => SqlQueryScheduleMode.EveryNMinutes,
            _ => SqlQueryScheduleMode.OneTime
        };
    }

    private string GetSqlQueryScheduleTargetLabel()
    {
        return _cmbSqlQueryScheduleTarget?.SelectedIndex switch
        {
            1 => "Scenario 2 - Invalid SQL Sample",
            2 => "Scenario 3 - Broken Connection Sample",
            3 => "All SQLQuery Scenarios",
            _ => "Scenario 1 - Scheduled Update Data"
        };
    }

    private string GetArmedSqlQueryScheduleModeLabel()
    {
        return _armedSqlQueryScheduleMode switch
        {
            SqlQueryScheduleMode.Daily => "repeats daily",
            SqlQueryScheduleMode.EveryNMinutes when _sqlQueryScheduleInterval.HasValue =>
                "every " + FormatSqlQueryIntervalLabel(_sqlQueryScheduleInterval.Value),
            _ => "one-time run"
        };
    }

    private static DateTime ResolveNextSqlQueryScheduledRun(DateTime requestedTime)
    {
        var nextRun = requestedTime;
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.AddDays(1);
        }

        return nextRun;
    }

    private static DateTime ResolveNextSqlQueryRecurringRun(DateTime previousScheduledAt, TimeSpan interval)
    {
        var nextRun = previousScheduledAt.Add(interval);
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.Add(interval);
        }

        return nextRun;
    }

    private static string FormatSqlQueryIntervalLabel(TimeSpan interval)
    {
        var totalMinutes = Math.Max(1, Convert.ToInt32(interval.TotalMinutes));
        return totalMinutes.ToString(CultureInfo.InvariantCulture) + " minute" + (totalMinutes == 1 ? string.Empty : "s");
    }

    private void PrepareProgramScenario1()
    {
        var resolvedPath = ResolveProgramPath(_txtProgramS1ScriptPath?.Text);
        if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
        {
            UpdateProgramStatus("Script path not found: " + (resolvedPath ?? "(empty)"), DangerColor);
            return;
        }

        if (!TryPrepareProgramOutput(_txtProgramS1OutputFile?.Text, out var outputDirectory, out var archivePath, out var errorMessage))
        {
            UpdateProgramStatus(errorMessage, DangerColor);
            return;
        }

        var archiveMessage = string.IsNullOrWhiteSpace(archivePath) ? "No prior output file was present." : "Archived prior output to " + archivePath + ".";
        UpdateProgramStatus("Scenario 1 ready. Script validated at " + resolvedPath + ", output folder prepared at " + outputDirectory + ". " + archiveMessage + GetProgramAutomationStatusSuffix(), SuccessColor);
    }

    private void PrepareProgramScenario3()
    {
        var resolvedPath = ResolveProgramPath(_txtProgramS3ScriptPath?.Text);
        if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath))
        {
            UpdateProgramScenario3Status("Script path not found: " + (resolvedPath ?? "(empty)"), DangerColor);
            return;
        }

        if (!TryPrepareProgramOutput(_txtProgramS3OutputFile?.Text, out var outputDirectory, out var archivePath, out var errorMessage))
        {
            UpdateProgramScenario3Status(errorMessage, DangerColor);
            return;
        }

        PreviewProgramCommand();
        var archiveMessage = string.IsNullOrWhiteSpace(archivePath) ? "No prior output file was present." : "Archived prior output to " + archivePath + ".";
        UpdateProgramScenario3Status("Scenario 3 ready. Argument-echo script validated at " + resolvedPath + ", output folder prepared at " + outputDirectory + ". " + archiveMessage + GetProgramAutomationStatusSuffix(), SuccessColor);
    }

    private static bool TryPrepareProgramOutput(string? outputFile, out string directory, out string? archivePath, out string errorMessage)
    {
        directory = string.Empty;
        archivePath = null;
        errorMessage = string.Empty;

        if (string.IsNullOrWhiteSpace(outputFile))
        {
            errorMessage = "OutputFile path is required.";
            return false;
        }

        directory = Path.GetDirectoryName(outputFile) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(directory))
        {
            errorMessage = "OutputFile must include a folder path.";
            return false;
        }

        Directory.CreateDirectory(directory);
        if (!File.Exists(outputFile))
        {
            return true;
        }

        archivePath =
            Path.Combine(
                directory,
                Path.GetFileNameWithoutExtension(outputFile) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + Path.GetExtension(outputFile));
        File.Move(outputFile, archivePath, overwrite: false);
        return true;
    }

    private string GetProgramAutomationStatusSuffix()
    {
        return _programAutomationExecutionContext is null
            ? string.Empty
            : " Scheduler batch " +
              _programAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
              "_r" +
              _programAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture) +
              ".";
    }

    private void OnProgramScheduleTimerTick(object? sender, EventArgs e)
    {
        HandleProgramScheduleTick();
    }

    private void ArmProgramSchedule()
    {
        if (_dtProgramScheduleAt is null ||
            _cmbProgramScheduleMode is null ||
            _numProgramScheduleIntervalMinutes is null)
        {
            return;
        }

        var mode = GetSelectedProgramScheduleMode();
        var scheduledAt = mode switch
        {
            ProgramScheduleMode.Daily => ResolveNextProgramScheduledRun(_dtProgramScheduleAt.Value),
            ProgramScheduleMode.EveryNMinutes => DateTime.Now.AddMinutes(decimal.ToDouble(_numProgramScheduleIntervalMinutes.Value)),
            _ => _dtProgramScheduleAt.Value
        };

        if (mode == ProgramScheduleMode.OneTime && scheduledAt <= DateTime.Now)
        {
            MessageBox.Show(
                "Choose a time in the future for a one-time ProgramExecution automation run.",
                "ProgramExecution Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _armedProgramScheduleMode = mode;
        _programScheduleInterval = mode == ProgramScheduleMode.EveryNMinutes
            ? TimeSpan.FromMinutes(decimal.ToDouble(_numProgramScheduleIntervalMinutes.Value))
            : null;
        _programScheduledRunAt = scheduledAt;
        _programScheduleTimer.Start();
        UpdateProgramScheduleStatus();
        UpdateProgramStatus(
            "Armed ProgramExecution automation for " + scheduledAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetProgramScheduleTargetLabel() +
            " | " + GetArmedProgramScheduleModeLabel(),
            SuccessColor);
    }

    private void CancelProgramSchedule()
    {
        _programScheduleTimer.Stop();
        _programScheduledRunAt = null;
        _programScheduleInterval = null;
        _armedProgramScheduleMode = ProgramScheduleMode.OneTime;
        UpdateProgramScheduleStatus();
        UpdateProgramStatus("Cancelled ProgramExecution automation timer.", WarningColor);
    }

    private void HandleProgramScheduleTick()
    {
        if (!_programScheduledRunAt.HasValue)
        {
            _programScheduleTimer.Stop();
            UpdateProgramScheduleStatus();
            return;
        }

        if (DateTime.Now < _programScheduledRunAt.Value)
        {
            UpdateProgramScheduleStatus();
            return;
        }

        RunProgramScheduledTarget();

        if (_armedProgramScheduleMode == ProgramScheduleMode.Daily)
        {
            _programScheduledRunAt = ResolveNextProgramScheduledRun(_programScheduledRunAt.Value.AddDays(1));
            _programScheduleTimer.Start();
        }
        else if (_armedProgramScheduleMode == ProgramScheduleMode.EveryNMinutes && _programScheduleInterval.HasValue)
        {
            _programScheduledRunAt = ResolveNextProgramRecurringRun(_programScheduledRunAt.Value, _programScheduleInterval.Value);
            _programScheduleTimer.Start();
        }
        else
        {
            _programScheduleTimer.Stop();
            _programScheduledRunAt = null;
            _programScheduleInterval = null;
        }

        UpdateProgramScheduleStatus();
    }

    private void RunProgramScheduledTarget()
    {
        try
        {
            _programAutomationExecutionContext = new ProgramAutomationExecutionContext(DateTime.Now, ++_programAutomationRunSequence);
            switch (_cmbProgramScheduleTarget?.SelectedIndex)
            {
                case 1:
                    ConfirmProgramPathMissing();
                    break;
                case 2:
                    PrepareProgramScenario3();
                    break;
                case 3:
                    PrepareProgramScenario1();
                    ConfirmProgramPathMissing();
                    PrepareProgramScenario3();
                    break;
                default:
                    PrepareProgramScenario1();
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Scheduled ProgramExecution preparation failed." + Environment.NewLine + Environment.NewLine + ex.Message,
                "ProgramExecution Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            UpdateProgramStatus("Scheduled ProgramExecution preparation failed: " + ex.Message, DangerColor);
        }
        finally
        {
            _programAutomationExecutionContext = null;
        }
    }

    private void UpdateProgramScheduleStatus()
    {
        if (_lblProgramScheduleStatus is null)
        {
            return;
        }

        if (!_programScheduledRunAt.HasValue)
        {
            _lblProgramScheduleStatus.Text = "ProgramExecution scheduler idle. Choose a mode, then arm the timer.";
            _lblProgramScheduleStatus.ForeColor = MutedText;
            return;
        }

        var remaining = _programScheduledRunAt.Value - DateTime.Now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        _lblProgramScheduleStatus.Text =
            "Armed for " + _programScheduledRunAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetProgramScheduleTargetLabel() +
            " | " + GetArmedProgramScheduleModeLabel() +
            " | remaining " + remaining.ToString(@"dd\.hh\:mm\:ss", CultureInfo.InvariantCulture);
        _lblProgramScheduleStatus.ForeColor = AccentColor;
    }

    private void UpdateProgramAutomationControlState()
    {
        var mode = GetSelectedProgramScheduleMode();
        if (_dtProgramScheduleAt is not null)
        {
            _dtProgramScheduleAt.Enabled = mode != ProgramScheduleMode.EveryNMinutes;
        }

        if (_numProgramScheduleIntervalMinutes is not null)
        {
            _numProgramScheduleIntervalMinutes.Enabled = mode == ProgramScheduleMode.EveryNMinutes;
        }
    }

    private ProgramScheduleMode GetSelectedProgramScheduleMode()
    {
        return _cmbProgramScheduleMode?.SelectedIndex switch
        {
            1 => ProgramScheduleMode.Daily,
            2 => ProgramScheduleMode.EveryNMinutes,
            _ => ProgramScheduleMode.OneTime
        };
    }

    private string GetProgramScheduleTargetLabel()
    {
        return _cmbProgramScheduleTarget?.SelectedIndex switch
        {
            1 => "Scenario 2 - Invalid Path Check",
            2 => "Scenario 3 - Argument Echo Prep",
            3 => "All ProgramExecution Scenarios",
            _ => "Scenario 1 - Successful Run Prep"
        };
    }

    private string GetArmedProgramScheduleModeLabel()
    {
        return _armedProgramScheduleMode switch
        {
            ProgramScheduleMode.Daily => "repeats daily",
            ProgramScheduleMode.EveryNMinutes when _programScheduleInterval.HasValue =>
                "every " + FormatProgramIntervalLabel(_programScheduleInterval.Value),
            _ => "one-time run"
        };
    }

    private static DateTime ResolveNextProgramScheduledRun(DateTime requestedTime)
    {
        var nextRun = requestedTime;
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.AddDays(1);
        }

        return nextRun;
    }

    private static DateTime ResolveNextProgramRecurringRun(DateTime previousScheduledAt, TimeSpan interval)
    {
        var nextRun = previousScheduledAt.Add(interval);
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.Add(interval);
        }

        return nextRun;
    }

    private static string FormatProgramIntervalLabel(TimeSpan interval)
    {
        var totalMinutes = Math.Max(1, Convert.ToInt32(interval.TotalMinutes));
        return totalMinutes.ToString(CultureInfo.InvariantCulture) + " minute" + (totalMinutes == 1 ? string.Empty : "s");
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
        UpdateProgramScenario3Status("Previewed full command line.", SuccessColor);
    }

    private void OpenDbToJsonOutputFolder()
    {
        var folder = EnsureDbToJsonOutputFolder();
        if (folder is null)
        {
            return;
        }

        OpenPath(folder);
    }

    private void PrepareDbToJsonOutputFolder()
    {
        var folder = EnsureDbToJsonOutputFolder();
        if (folder is null)
        {
            return;
        }

        SetDbToJsonStatus("Prepared DBtoJSON output folder: " + folder, SuccessColor);
    }

    private string? EnsureDbToJsonOutputFolder()
    {
        var folder = _txtDbToJsonOutputFolder?.Text.Trim();
        if (string.IsNullOrWhiteSpace(folder))
        {
            MessageBox.Show("Set the DBtoJSON output folder first.", "DBtoJSON", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        try
        {
            Directory.CreateDirectory(folder);
            return folder;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Unable to prepare the DBtoJSON output folder." + Environment.NewLine + Environment.NewLine + ex.Message,
                "DBtoJSON",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return null;
        }
    }

    private DateTime GetDbToJsonScenario1BaseTime()
    {
        return _dbToJsonAutomationExecutionContext?.RunStamp ?? _dtDbToJsonS1CreatedAtBase?.Value ?? _settings.DatetimeBase;
    }

    private DateTime GetDbToJsonScenario2BaseTime()
    {
        return _dbToJsonAutomationExecutionContext?.RunStamp ?? new DateTime(2026, 4, 6, 14, 30, 0);
    }

    private string GetDbToJsonOutputFolderDisplay()
    {
        var folder = _txtDbToJsonOutputFolder?.Text.Trim();
        return string.IsNullOrWhiteSpace(folder) ? "(output folder not set)" : folder;
    }

    private int GetDbToJsonAutomationIdOffset()
    {
        return _dbToJsonAutomationExecutionContext is null
            ? 0
            : _dbToJsonAutomationExecutionContext.RunSequence * 1000;
    }

    private string GetDbToJsonAutomationStatusSuffix()
    {
        return _dbToJsonAutomationExecutionContext is null
            ? string.Empty
            : " Generated by scheduler batch " +
              _dbToJsonAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
              "_r" +
              _dbToJsonAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture) +
              ".";
    }

    private static string BuildDbToJsonPayload(int exportId, bool apiMode)
    {
        return apiMode
            ? "{\"exportId\":" + exportId.ToString(CultureInfo.InvariantCulture) + ",\"destination\":\"api\"}"
            : "{\"exportId\":" + exportId.ToString(CultureInfo.InvariantCulture) + ",\"destination\":\"file\"}";
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

    private void UpdateProgramScenario3Status(string message, Color color)
    {
        if (_lblProgramS3Status is not null)
        {
            _lblProgramS3Status.Text = message;
            _lblProgramS3Status.ForeColor = color;
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

        if (_txtDbToJsonOutputFolder is not null && string.IsNullOrWhiteSpace(_txtDbToJsonOutputFolder.Text))
        {
            _txtDbToJsonOutputFolder.Text = Path.Combine(_settings.OutputRootFolder, "DBtoJSON", "JsonOutput");
        }

        RefreshSqlQueryScenarioTemplates();
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

    private static int ExecuteNonQueryCount(SqlConnection connection, SqlTransaction transaction, string sql)
    {
        using var command = new SqlCommand(sql, connection, transaction);
        return command.ExecuteNonQuery();
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

    private static int DeleteDbToDbRowsByIdRange(SqlConnection connection, SqlTransaction transaction, string tableName, int minimumSourceId, int maximumSourceId)
    {
        using var command = new SqlCommand(
            "DELETE FROM " + tableName + " WHERE SourceId BETWEEN @minId AND @maxId",
            connection,
            transaction);
        command.Parameters.AddWithValue("@minId", minimumSourceId);
        command.Parameters.AddWithValue("@maxId", maximumSourceId);
        return command.ExecuteNonQuery();
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

    private static int DeleteDbToJsonRowsByIdRange(SqlConnection connection, SqlTransaction transaction, string tableName, int minimumExportId, int maximumExportId)
    {
        using var command = new SqlCommand(
            "DELETE FROM " + tableName + " WHERE ExportId BETWEEN @minId AND @maxId",
            connection,
            transaction);
        command.Parameters.AddWithValue("@minId", minimumExportId);
        command.Parameters.AddWithValue("@maxId", maximumExportId);
        return command.ExecuteNonQuery();
    }

    private static int DeleteSqlQueryRowsByIdRange(SqlConnection connection, SqlTransaction transaction, string tableName, int minimumRowId, int maximumRowId)
    {
        using var command = new SqlCommand(
            "DELETE FROM " + tableName + " WHERE RowId BETWEEN @minId AND @maxId",
            connection,
            transaction);
        command.Parameters.AddWithValue("@minId", minimumRowId);
        command.Parameters.AddWithValue("@maxId", maximumRowId);
        return command.ExecuteNonQuery();
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

    private enum DbToDbScheduleMode
    {
        OneTime,
        Daily,
        EveryNMinutes
    }

    private sealed record DbToDbAutomationExecutionContext(DateTime RunStamp, int RunSequence);

    private enum DbToJsonScheduleMode
    {
        OneTime,
        Daily,
        EveryNMinutes
    }

    private sealed record DbToJsonAutomationExecutionContext(DateTime RunStamp, int RunSequence);

    private enum SqlQueryScheduleMode
    {
        OneTime,
        Daily,
        EveryNMinutes
    }

    private sealed record SqlQueryAutomationExecutionContext(DateTime RunStamp, int RunSequence);

    private enum ProgramScheduleMode
    {
        OneTime,
        Daily,
        EveryNMinutes
    }

    private sealed record ProgramAutomationExecutionContext(DateTime RunStamp, int RunSequence);

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
