using System.Diagnostics;
using System.Globalization;
using System.Text;
using Microsoft.Data.SqlClient;

namespace DataSyncer.TestDataGenerator.App;

public partial class Form1 : Form
{
    private readonly Models.AppSettings _settings;
    private readonly Services.ConfigurationService _configurationService;
    private readonly Services.FileLogService _logService;

    private ToolStripStatusLabel? _statusLabel;
    private Label? _lblDashboardConfig;

    private TextBox? _txtSqlConnectionString;
    private TextBox? _txtOutputRootFolder;
    private TextBox? _txtSyncFlagNotProcessed;
    private TextBox? _txtSyncFlagProcessed;
    private TextBox? _txtCsvIdPrefix;
    private TextBox? _txtApiTestEndpoint;
    private TextBox? _txtRemoteFileSharePath;
    private DateTimePicker? _dtDatetimeBase;
    private NumericUpDown? _numCsvIdStart;
    private NumericUpDown? _numDbToDbIdStart;
    private NumericUpDown? _numDbToJsonIdStart;
    private NumericUpDown? _numSqlQueryIdStart;
    private TextBox? _txtFileSyncUploadPrefix;
    private TextBox? _txtFileSyncDownloadPrefix;
    private TextBox? _txtFileSyncTwoWayPrefix;
    private RichTextBox? _logViewer;

    private TextBox? _txtCsvOutputFolder;
    private DateTimePicker? _dtCsvBase;
    private ComboBox? _cmbCsvScenario;
    private Label? _lblCsvStatus;

    private TextBox? _txtCsvS1OutputFilename;
    private NumericUpDown? _numCsvS1RowCount;
    private TextBox? _txtCsvS1KeyStart;
    private CheckBox? _chkCsvS1IncludeComment;
    private TextBox? _txtCsvS1MachineCodes;
    private TextBox? _txtCsvS1Statuses;

    private TextBox? _txtCsvS2OutputFilename;
    private NumericUpDown? _numCsvS2RowCount;
    private TextBox? _txtCsvS2KeyStart;
    private CheckBox? _chkCsvS2OmitComment;

    private TextBox? _txtCsvS3SeedFilename;
    private NumericUpDown? _numCsvS3RowCount;
    private TextBox? _txtCsvS3KeyStart;
    private CheckBox? _chkCsvS3GenerateVariant;
    private TextBox? _txtCsvS3VariantFilename;

    public Form1(
        Models.AppSettings settings,
        Services.ConfigurationService configurationService,
        Services.FileLogService logService)
    {
        _settings = settings;
        _configurationService = configurationService;
        _logService = logService;

        InitializeComponent();
        BuildUi();
        UpdateDashboardSummary();
        WriteStatus("Ready");

        _logService.LogInfo("Application started");
        RefreshLogViewer();
    }

    private void BuildUi()
    {
        SuspendLayout();

        var rootPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(8)
        };
        rootPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        rootPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        rootPanel.Controls.Add(CreateMainTabControl(), 0, 0);

        var statusStrip = new StatusStrip();
        _statusLabel = new ToolStripStatusLabel("Initializing...");
        statusStrip.Items.Add(_statusLabel);
        rootPanel.Controls.Add(statusStrip, 0, 1);

        Controls.Add(rootPanel);
        ResumeLayout();
    }

    private TabControl CreateMainTabControl()
    {
        var tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
            Alignment = TabAlignment.Left,
            SizeMode = TabSizeMode.Fixed,
            ItemSize = new Size(36, 170),
            Multiline = true
        };

        tabControl.TabPages.Add(CreateDashboardTab());
        tabControl.TabPages.Add(CreateCsvToDbTab());
        tabControl.TabPages.Add(CreateScenarioInfoTab("DBtoDB", "DB source and destination row inserter for Scenarios 1-3"));
        tabControl.TabPages.Add(CreateScenarioInfoTab("DBtoJSON", "Export-source row inserter for Scenarios 1-3"));
        tabControl.TabPages.Add(CreateScenarioInfoTab("SQLQuery", "Test table populator for update-query scenarios"));
        tabControl.TabPages.Add(CreateScenarioInfoTab("ProgramExecution", "Path validator and output folder creator"));
        tabControl.TabPages.Add(CreateScenarioInfoTab("FileSyncer", "Local file creator and remote folder checker"));
        tabControl.TabPages.Add(CreateSettingsTab());
        tabControl.TabPages.Add(CreateLogViewerTab());

        return tabControl;
    }

    private TabPage CreateDashboardTab()
    {
        var tab = new TabPage("Home / Dashboard");
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(16)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        panel.Controls.Add(new Label
        {
            Text = "DataSyncer Test Data Generator",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            AutoSize = true
        }, 0, 0);

        panel.Controls.Add(new Label
        {
            Text = "Summary of configured connection strings and output paths.",
            AutoSize = true,
            ForeColor = Color.DimGray
        }, 0, 1);

        _lblDashboardConfig = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(980, 0)
        };
        panel.Controls.Add(_lblDashboardConfig, 0, 2);

        var btnGenerateAll = new Button
        {
            Text = "Generate All",
            AutoSize = true,
            Padding = new Padding(12, 6, 12, 6)
        };
        btnGenerateAll.Click += (_, _) => RunGenerateAllPlaceholder();
        panel.Controls.Add(btnGenerateAll, 0, 3);

        tab.Controls.Add(panel);
        return tab;
    }

    private TabPage CreateCsvToDbTab()
    {
        var tab = new TabPage("CSVtoDB");
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            Padding = new Padding(16),
            AutoScroll = true
        };

        panel.Controls.Add(new Label
        {
            Text = "CSV file generator for Scenarios 1-4",
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            AutoSize = true
        });

        panel.Controls.Add(CreateCsvSharedControlsGroup());
        panel.Controls.Add(CreateCsvScenario1Group());
        panel.Controls.Add(CreateCsvScenario2Group());
        panel.Controls.Add(CreateCsvScenario3Group());
        panel.Controls.Add(CreateCsvScenario4Group());

        tab.Controls.Add(panel);
        return tab;
    }

    private GroupBox CreateCsvSharedControlsGroup()
    {
        var group = new GroupBox { Text = "Shared CSV Controls", Dock = DockStyle.Top, AutoSize = true };
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(10)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddLabel(panel, "Output folder path", 0, 0);
        _txtCsvOutputFolder = new TextBox
        {
            Text = _settings.OutputRootFolder,
            Anchor = AnchorStyles.Left | AnchorStyles.Right
        };
        panel.Controls.Add(_txtCsvOutputFolder, 1, 0);

        AddLabel(panel, "Date/time base", 0, 1);
        _dtCsvBase = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Value = _settings.DatetimeBase,
            Anchor = AnchorStyles.Left
        };
        panel.Controls.Add(_dtCsvBase, 1, 1);

        AddLabel(panel, "Selected scenario", 0, 2);
        _cmbCsvScenario = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left,
            Width = 240
        };
        _cmbCsvScenario.Items.AddRange(new object[]
        {
            "Scenario 1 - Happy Path",
            "Scenario 2 - Optional Column Missing",
            "Scenario 3 - Duplicate Prevention",
            "Scenario 4 - Scheduled Recurring Run"
        });
        _cmbCsvScenario.SelectedIndex = 0;
        panel.Controls.Add(_cmbCsvScenario, 1, 2);

        var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };

        var btnGenerateSelected = new Button { Text = "Generate Selected Scenario", AutoSize = true };
        btnGenerateSelected.Click += (_, _) => GenerateSelectedCsvScenario();
        buttonPanel.Controls.Add(btnGenerateSelected);

        var btnGenerateAll = new Button { Text = "Generate All CSV Scenarios", AutoSize = true };
        btnGenerateAll.Click += (_, _) => GenerateAllCsvScenarios();
        buttonPanel.Controls.Add(btnGenerateAll);

        var btnClearFolder = new Button { Text = "Clear Output Folder", AutoSize = true };
        btnClearFolder.Click += (_, _) => ClearCsvOutputFolder();
        buttonPanel.Controls.Add(btnClearFolder);

        panel.Controls.Add(buttonPanel, 1, 3);

        _lblCsvStatus = new Label
        {
            Text = "Last generated: (none)",
            AutoSize = true,
            ForeColor = Color.DimGray
        };
        panel.Controls.Add(_lblCsvStatus, 1, 4);

        group.Controls.Add(panel);
        return group;
    }

    private GroupBox CreateCsvScenario1Group()
    {
        var group = new GroupBox { Text = "Scenario 1 - Happy Path CSV Import", Dock = DockStyle.Top, AutoSize = true };
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(10)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddLabel(panel, "Output filename", 0, 0);
        _txtCsvS1OutputFilename = new TextBox { Text = "csv_happy_20260406_090000.csv", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS1OutputFilename, 1, 0);

        AddLabel(panel, "Row count (5-500)", 0, 1);
        _numCsvS1RowCount = new NumericUpDown { Minimum = 5, Maximum = 500, Value = 20, Anchor = AnchorStyles.Left };
        panel.Controls.Add(_numCsvS1RowCount, 1, 1);

        AddLabel(panel, "Key start", 0, 2);
        _txtCsvS1KeyStart = new TextBox { Text = _settings.CsvIdPrefix + _settings.CsvIdStart.ToString(CultureInfo.InvariantCulture), Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS1KeyStart, 1, 2);

        AddLabel(panel, "Include Comment column", 0, 3);
        _chkCsvS1IncludeComment = new CheckBox { Checked = true, AutoSize = true };
        panel.Controls.Add(_chkCsvS1IncludeComment, 1, 3);

        AddLabel(panel, "MachineCode values", 0, 4);
        _txtCsvS1MachineCodes = new TextBox { Text = "MC-01, MC-02, MC-03", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS1MachineCodes, 1, 4);

        AddLabel(panel, "Status values", 0, 5);
        _txtCsvS1Statuses = new TextBox { Text = "NEW, RUN, DONE, WAIT", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS1Statuses, 1, 5);

        panel.Controls.Add(new Label
        {
            Text = "Pre-state note: Destination must not already contain these RecordId values. Monitored input folder must contain only this file.",
            AutoSize = true,
            ForeColor = Color.DimGray,
            MaximumSize = new Size(820, 0)
        }, 1, 6);

        group.Controls.Add(panel);
        return group;
    }

    private GroupBox CreateCsvScenario2Group()
    {
        var group = new GroupBox { Text = "Scenario 2 - Optional Column Missing", Dock = DockStyle.Top, AutoSize = true };
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(10)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddLabel(panel, "Output filename", 0, 0);
        _txtCsvS2OutputFilename = new TextBox { Text = "csv_optional_missing_20260406_091000.csv", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS2OutputFilename, 1, 0);

        AddLabel(panel, "Row count (2-100)", 0, 1);
        _numCsvS2RowCount = new NumericUpDown { Minimum = 2, Maximum = 100, Value = 5, Anchor = AnchorStyles.Left };
        panel.Controls.Add(_numCsvS2RowCount, 1, 1);

        AddLabel(panel, "Key start", 0, 2);
        _txtCsvS2KeyStart = new TextBox { Text = "CSV-100021", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS2KeyStart, 1, 2);

        AddLabel(panel, "Omit Comment column", 0, 3);
        _chkCsvS2OmitComment = new CheckBox { Checked = true, AutoSize = true };
        panel.Controls.Add(_chkCsvS2OmitComment, 1, 3);

        panel.Controls.Add(new Label
        {
            Text = "Pre-state note: Mapping must be configured so the missing CSV header is optional, not required.",
            AutoSize = true,
            ForeColor = Color.DimGray,
            MaximumSize = new Size(820, 0)
        }, 1, 4);

        group.Controls.Add(panel);
        return group;
    }

    private GroupBox CreateCsvScenario3Group()
    {
        var group = new GroupBox { Text = "Scenario 3 - Duplicate Prevention", Dock = DockStyle.Top, AutoSize = true };
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(10)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddLabel(panel, "Seed filename", 0, 0);
        _txtCsvS3SeedFilename = new TextBox { Text = "csv_duplicate_seed_20260406_092000.csv", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS3SeedFilename, 1, 0);

        AddLabel(panel, "Row count", 0, 1);
        _numCsvS3RowCount = new NumericUpDown { Minimum = 1, Maximum = 500, Value = 10, Anchor = AnchorStyles.Left };
        panel.Controls.Add(_numCsvS3RowCount, 1, 1);

        AddLabel(panel, "Key start", 0, 2);
        _txtCsvS3KeyStart = new TextBox { Text = "CSV-101001", Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(_txtCsvS3KeyStart, 1, 2);

        AddLabel(panel, "Generate variant file", 0, 3);
        _chkCsvS3GenerateVariant = new CheckBox { Checked = false, AutoSize = true };
        _chkCsvS3GenerateVariant.CheckedChanged += (_, _) => ToggleScenario3VariantState();
        panel.Controls.Add(_chkCsvS3GenerateVariant, 1, 3);

        AddLabel(panel, "Variant filename", 0, 4);
        _txtCsvS3VariantFilename = new TextBox { Text = "csv_duplicate_variant.csv", Anchor = AnchorStyles.Left | AnchorStyles.Right, Enabled = false };
        panel.Controls.Add(_txtCsvS3VariantFilename, 1, 4);

        panel.Controls.Add(new Label
        {
            Text = "Variant behavior: 8 duplicate + 2 new keys based on seed file range.",
            AutoSize = true,
            ForeColor = Color.DimGray,
            MaximumSize = new Size(820, 0)
        }, 1, 5);

        group.Controls.Add(panel);
        return group;
    }

    private GroupBox CreateCsvScenario4Group()
    {
        var group = new GroupBox { Text = "Scenario 4 - Scheduled Recurring Run", Dock = DockStyle.Top, AutoSize = true };
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            Padding = new Padding(10)
        };

        panel.Controls.Add(new Label
        {
            Text = "Generate All 4 Files creates file_A, file_B, file_C, and file_A_REDROP copy with required key ranges and timestamps.",
            AutoSize = true,
            MaximumSize = new Size(980, 0)
        });

        panel.Controls.Add(new Label
        {
            Text = "file_A_yyyyMMdd_100000.csv => CSV-102001 to CSV-102010\n" +
                   "file_B_yyyyMMdd_100100.csv => CSV-102011 to CSV-102020\n" +
                   "file_C_yyyyMMdd_100300.csv => CSV-102021 to CSV-102030\n" +
                   "file_A_REDROP_copy.csv => same keys and content as file_A",
            AutoSize = true,
            ForeColor = Color.DimGray,
            MaximumSize = new Size(980, 0)
        });

        var btnGenerate = new Button { Text = "Generate All 4 Files", AutoSize = true };
        btnGenerate.Click += (_, _) => GenerateCsvScenario4();
        panel.Controls.Add(btnGenerate);

        group.Controls.Add(panel);
        return group;
    }

    private void ToggleScenario3VariantState()
    {
        if (_chkCsvS3GenerateVariant is not null && _txtCsvS3VariantFilename is not null)
        {
            _txtCsvS3VariantFilename.Enabled = _chkCsvS3GenerateVariant.Checked;
        }
    }

    private void GenerateSelectedCsvScenario()
    {
        if (_cmbCsvScenario is null)
        {
            return;
        }

        switch (_cmbCsvScenario.SelectedIndex)
        {
            case 0:
                GenerateCsvScenario1();
                break;
            case 1:
                GenerateCsvScenario2();
                break;
            case 2:
                GenerateCsvScenario3();
                break;
            case 3:
                GenerateCsvScenario4();
                break;
            default:
                MessageBox.Show("Please select a scenario.", "CSVtoDB", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                break;
        }
    }

    private void GenerateAllCsvScenarios()
    {
        GenerateCsvScenario1();
        GenerateCsvScenario2();
        GenerateCsvScenario3();
        GenerateCsvScenario4();
    }

    private void GenerateCsvScenario1()
    {
        if (_txtCsvS1OutputFilename is null ||
            _numCsvS1RowCount is null ||
            _txtCsvS1KeyStart is null ||
            _chkCsvS1IncludeComment is null ||
            _txtCsvS1MachineCodes is null ||
            _txtCsvS1Statuses is null)
        {
            return;
        }

        var outputFileName = _txtCsvS1OutputFilename.Text.Trim();
        if (!ValidateCsvFileName(outputFileName, allowNoPath: false))
        {
            MessageBox.Show("Scenario 1 output filename must end in .csv and must not contain path separators.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!TryParseKeyStart(_txtCsvS1KeyStart.Text.Trim(), out var prefix, out var startNumber))
        {
            MessageBox.Show("Scenario 1 Key start must be non-empty and end with numeric characters, e.g. CSV-100001.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var machineCodes = ParseCommaSeparatedValues(_txtCsvS1MachineCodes.Text);
        var statuses = ParseCommaSeparatedValues(_txtCsvS1Statuses.Text);
        if (machineCodes.Count == 0 || statuses.Count == 0)
        {
            MessageBox.Show("Scenario 1 MachineCode and Status values must be comma-separated with at least one value each.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rows = BuildCsvRows(prefix, startNumber, Decimal.ToInt32(_numCsvS1RowCount.Value), machineCodes, statuses, _chkCsvS1IncludeComment.Checked, GetCsvBaseTime());
        WriteCsvToOutput(outputFileName, rows, _chkCsvS1IncludeComment.Checked, "CSVtoDB Scenario 1");
    }

    private void GenerateCsvScenario2()
    {
        if (_txtCsvS2OutputFilename is null ||
            _numCsvS2RowCount is null ||
            _txtCsvS2KeyStart is null ||
            _chkCsvS2OmitComment is null)
        {
            return;
        }

        var outputFileName = _txtCsvS2OutputFilename.Text.Trim();
        if (!ValidateCsvFileName(outputFileName, allowNoPath: false))
        {
            MessageBox.Show("Scenario 2 output filename must end in .csv.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!TryParseKeyStart(_txtCsvS2KeyStart.Text.Trim(), out var prefix, out var startNumber))
        {
            MessageBox.Show("Scenario 2 Key start must be non-empty and end with numeric characters, e.g. CSV-100021.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rowCount = Decimal.ToInt32(_numCsvS2RowCount.Value);
        var includeComment = !_chkCsvS2OmitComment.Checked;
        var rows = BuildCsvRows(
            prefix,
            startNumber,
            rowCount,
            new List<string> { "MC-01", "MC-02", "MC-03" },
            new List<string> { "NEW", "RUN", "DONE", "WAIT" },
            includeComment,
            GetCsvBaseTime());

        WriteCsvToOutput(outputFileName, rows, includeComment, "CSVtoDB Scenario 2");
    }

    private void GenerateCsvScenario3()
    {
        if (_txtCsvS3SeedFilename is null ||
            _numCsvS3RowCount is null ||
            _txtCsvS3KeyStart is null ||
            _chkCsvS3GenerateVariant is null ||
            _txtCsvS3VariantFilename is null)
        {
            return;
        }

        var seedFileName = _txtCsvS3SeedFilename.Text.Trim();
        if (!ValidateCsvFileName(seedFileName, allowNoPath: false))
        {
            MessageBox.Show("Scenario 3 seed filename must end in .csv.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!TryParseKeyStart(_txtCsvS3KeyStart.Text.Trim(), out var prefix, out var startNumber))
        {
            MessageBox.Show("Scenario 3 Key start must be non-empty and end with numeric characters, e.g. CSV-101001.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rowCount = Decimal.ToInt32(_numCsvS3RowCount.Value);
        var baseTime = GetCsvBaseTime();

        var seedRows = BuildCsvRows(
            prefix,
            startNumber,
            rowCount,
            new List<string> { "MC-01", "MC-02", "MC-03" },
            new List<string> { "NEW", "RUN", "DONE", "WAIT" },
            includeComment: true,
            baseTime);

        WriteCsvToOutput(seedFileName, seedRows, includeComment: true, "CSVtoDB Scenario 3 (seed)");

        if (_chkCsvS3GenerateVariant.Checked)
        {
            var variantFileName = _txtCsvS3VariantFilename.Text.Trim();
            if (!ValidateCsvFileName(variantFileName, allowNoPath: false))
            {
                MessageBox.Show("Scenario 3 variant filename must end in .csv.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var duplicateCount = Math.Min(8, seedRows.Count);
            var variantRows = new List<CsvRow>();

            for (var i = 0; i < duplicateCount; i++)
            {
                variantRows.Add(seedRows[i]);
            }

            var nextId = startNumber + rowCount;
            for (var i = 0; i < 2; i++)
            {
                variantRows.Add(new CsvRow(
                    prefix + (nextId + i).ToString(CultureInfo.InvariantCulture),
                    i % 2 == 0 ? "MC-02" : "MC-03",
                    i % 2 == 0 ? "NEW" : "RUN",
                    baseTime.AddMinutes(duplicateCount + i),
                    "Variant Row " + (duplicateCount + i + 1).ToString(CultureInfo.InvariantCulture)));
            }

            WriteCsvToOutput(variantFileName, variantRows, includeComment: true, "CSVtoDB Scenario 3 (variant)");
        }
    }

    private void GenerateCsvScenario4()
    {
        var baseDate = GetCsvBaseTime().Date;

        var fileAName = "file_A_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100000.csv";
        var fileBName = "file_B_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100100.csv";
        var fileCName = "file_C_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100300.csv";
        var fileRedropName = "file_A_REDROP_copy.csv";

        var rowsA = BuildCsvRows(
            _settings.CsvIdPrefix,
            102001,
            10,
            new List<string> { "MC-01", "MC-02", "MC-03" },
            new List<string> { "NEW", "RUN", "DONE", "WAIT" },
            includeComment: true,
            baseDate.AddHours(10));

        var rowsB = BuildCsvRows(
            _settings.CsvIdPrefix,
            102011,
            10,
            new List<string> { "MC-01", "MC-02", "MC-03" },
            new List<string> { "NEW", "RUN", "DONE", "WAIT" },
            includeComment: true,
            baseDate.AddHours(10).AddMinutes(1));

        var rowsC = BuildCsvRows(
            _settings.CsvIdPrefix,
            102021,
            10,
            new List<string> { "MC-01", "MC-02", "MC-03" },
            new List<string> { "NEW", "RUN", "DONE", "WAIT" },
            includeComment: true,
            baseDate.AddHours(10).AddMinutes(3));

        WriteCsvToOutput(fileAName, rowsA, includeComment: true, "CSVtoDB Scenario 4 (file A)");
        WriteCsvToOutput(fileBName, rowsB, includeComment: true, "CSVtoDB Scenario 4 (file B)");
        WriteCsvToOutput(fileCName, rowsC, includeComment: true, "CSVtoDB Scenario 4 (file C)");
        WriteCsvToOutput(fileRedropName, rowsA, includeComment: true, "CSVtoDB Scenario 4 (file A redrop)");

        var outputFolder = EnsureCsvOutputFolder();
        if (outputFolder is null)
        {
            return;
        }

        var result = MessageBox.Show(
            "Scenario 4 files generated in:\n" + outputFolder + "\n\nOpen folder in Explorer?",
            "CSVtoDB Scenario 4",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information);

        if (result == DialogResult.Yes)
        {
            OpenFolderInExplorer(outputFolder);
        }
    }

    private void ClearCsvOutputFolder()
    {
        var outputFolder = EnsureCsvOutputFolder();
        if (outputFolder is null)
        {
            return;
        }

        var confirm = MessageBox.Show(
            "Delete all .csv files from this folder?\n" + outputFolder,
            "Clear Output Folder",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
        {
            return;
        }

        var deleted = 0;
        foreach (var file in Directory.GetFiles(outputFolder, "*.csv"))
        {
            File.Delete(file);
            deleted++;
        }

        _logService.LogInfo("CSVtoDB: Cleared output folder, deleted " + deleted.ToString(CultureInfo.InvariantCulture) + " csv files");
        SetCsvStatus("(folder cleared)", 0);
        WriteStatus("CSV output folder cleared");
        RefreshLogViewer();
    }

    private List<CsvRow> BuildCsvRows(
        string keyPrefix,
        int startNumber,
        int rowCount,
        IReadOnlyList<string> machineCodes,
        IReadOnlyList<string> statuses,
        bool includeComment,
        DateTime baseTime)
    {
        var rows = new List<CsvRow>(rowCount);
        for (var i = 0; i < rowCount; i++)
        {
            var recordId = keyPrefix + (startNumber + i).ToString(CultureInfo.InvariantCulture);
            var machineCode = machineCodes[i % machineCodes.Count];
            var status = statuses[i % statuses.Count];
            var rowTime = baseTime.AddMinutes(i);
            var comment = includeComment ? "Comment " + (i + 1).ToString(CultureInfo.InvariantCulture) : string.Empty;
            rows.Add(new CsvRow(recordId, machineCode, status, rowTime, comment));
        }

        return rows;
    }

    private void WriteCsvToOutput(string fileName, List<CsvRow> rows, bool includeComment, string contextLabel)
    {
        var outputFolder = EnsureCsvOutputFolder();
        if (outputFolder is null)
        {
            return;
        }

        var fullPath = Path.Combine(outputFolder, fileName);
        var content = BuildCsvContent(rows, includeComment);
        File.WriteAllText(fullPath, content, new UTF8Encoding(false));

        _logService.LogInfo(contextLabel + ": Generated file " + fullPath + " with row count " + rows.Count.ToString(CultureInfo.InvariantCulture));
        SetCsvStatus(fileName, rows.Count);
        WriteStatus(contextLabel + " completed");
        RefreshLogViewer();
    }

    private static string BuildCsvContent(IEnumerable<CsvRow> rows, bool includeComment)
    {
        var sb = new StringBuilder();
        sb.Append("RecordId,MachineCode,Status,CreatedAt");
        if (includeComment)
        {
            sb.Append(",Comment");
        }

        sb.AppendLine();

        foreach (var row in rows)
        {
            sb.Append(EscapeCsv(row.RecordId));
            sb.Append(',');
            sb.Append(EscapeCsv(row.MachineCode));
            sb.Append(',');
            sb.Append(EscapeCsv(row.Status));
            sb.Append(',');
            sb.Append(EscapeCsv(row.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)));
            if (includeComment)
            {
                sb.Append(',');
                sb.Append(EscapeCsv(row.Comment));
            }

            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static string EscapeCsv(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        return value;
    }

    private void SetCsvStatus(string fileName, int rowCount)
    {
        if (_lblCsvStatus is not null)
        {
            _lblCsvStatus.Text = "Last generated: " + fileName + " (rows: " + rowCount.ToString(CultureInfo.InvariantCulture) + ")";
        }
    }

    private DateTime GetCsvBaseTime()
    {
        if (_dtCsvBase is not null)
        {
            return _dtCsvBase.Value;
        }

        return _settings.DatetimeBase;
    }

    private string? EnsureCsvOutputFolder()
    {
        if (_txtCsvOutputFolder is null)
        {
            return null;
        }

        var path = _txtCsvOutputFolder.Text.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show("Output folder path is required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        Directory.CreateDirectory(path);
        return path;
    }

    private static bool ValidateCsvFileName(string fileName, bool allowNoPath)
    {
        if (string.IsNullOrWhiteSpace(fileName) || !fileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!allowNoPath)
        {
            if (fileName.Contains(Path.DirectorySeparatorChar) || fileName.Contains(Path.AltDirectorySeparatorChar))
            {
                return false;
            }
        }

        return true;
    }

    private static List<string> ParseCommaSeparatedValues(string input)
    {
        return input
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(v => v.Trim())
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();
    }

    private static bool TryParseKeyStart(string keyStart, out string prefix, out int number)
    {
        prefix = string.Empty;
        number = 0;

        if (string.IsNullOrWhiteSpace(keyStart))
        {
            return false;
        }

        var index = keyStart.Length - 1;
        while (index >= 0 && char.IsDigit(keyStart[index]))
        {
            index--;
        }

        var numericPart = keyStart[(index + 1)..];
        prefix = keyStart[..(index + 1)];

        if (string.IsNullOrWhiteSpace(prefix) || string.IsNullOrWhiteSpace(numericPart))
        {
            return false;
        }

        return int.TryParse(numericPart, out number);
    }

    private static void OpenFolderInExplorer(string folderPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "\"" + folderPath + "\"",
            UseShellExecute = true
        };

        Process.Start(startInfo);
    }

    private TabPage CreateScenarioInfoTab(string title, string purpose)
    {
        var tab = new TabPage(title);
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(16)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        panel.Controls.Add(new Label
        {
            Text = title,
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            AutoSize = true
        }, 0, 0);

        panel.Controls.Add(new Label
        {
            Text = purpose,
            AutoSize = true,
            ForeColor = Color.DimGray
        }, 0, 1);

        var actions = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        actions.Controls.Add(CreateActionButton("Generate Scenario 1", title + ": Scenario 1 requested"));
        actions.Controls.Add(CreateActionButton("Generate Scenario 2", title + ": Scenario 2 requested"));
        actions.Controls.Add(CreateActionButton("Generate Scenario 3", title + ": Scenario 3 requested"));
        if (title == "ProgramExecution" || title == "FileSyncer")
        {
            actions.Controls.Add(CreateActionButton("Validate Paths", title + ": Path validation requested"));
        }

        panel.Controls.Add(actions, 0, 2);
        tab.Controls.Add(panel);
        return tab;
    }

    private TabPage CreateSettingsTab()
    {
        var tab = new TabPage("Settings");
        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        var stack = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            Padding = new Padding(16)
        };

        stack.Controls.Add(CreateConnectionSettingsGroup());
        stack.Controls.Add(CreateIdRangeGroup());
        stack.Controls.Add(CreateFileSyncPrefixGroup());

        var actions = new FlowLayoutPanel { AutoSize = true };
        var btnTestConnection = new Button { Text = "Test SQL Connection", AutoSize = true };
        btnTestConnection.Click += (_, _) => TestSqlConnection();
        actions.Controls.Add(btnTestConnection);

        var btnSave = new Button { Text = "Save Settings", AutoSize = true };
        btnSave.Click += (_, _) => SaveSettings();
        actions.Controls.Add(btnSave);

        stack.Controls.Add(actions);
        scroll.Controls.Add(stack);
        tab.Controls.Add(scroll);
        return tab;
    }

    private GroupBox CreateConnectionSettingsGroup()
    {
        var group = new GroupBox { Text = "Connection, Paths, and Flags", AutoSize = true, Width = 1020 };
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true, Padding = new Padding(10) };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _txtSqlConnectionString = AddLabeledTextBox(panel, "SqlConnectionString", _settings.SqlConnectionString, 0);
        _txtOutputRootFolder = AddLabeledTextBox(panel, "OutputRootFolder", _settings.OutputRootFolder, 1);
        _txtSyncFlagNotProcessed = AddLabeledTextBox(panel, "SyncFlagNotProcessed", _settings.SyncFlagNotProcessed, 2);
        _txtSyncFlagProcessed = AddLabeledTextBox(panel, "SyncFlagProcessed", _settings.SyncFlagProcessed, 3);
        _txtCsvIdPrefix = AddLabeledTextBox(panel, "CsvIdPrefix", _settings.CsvIdPrefix, 4);
        _txtApiTestEndpoint = AddLabeledTextBox(panel, "ApiTestEndpoint", _settings.ApiTestEndpoint, 5);
        _txtRemoteFileSharePath = AddLabeledTextBox(panel, "RemoteFileSharePath", _settings.RemoteFileSharePath, 6);

        AddLabel(panel, "DatetimeBase", 0, 7);
        _dtDatetimeBase = new DateTimePicker
        {
            Anchor = AnchorStyles.Left,
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Value = _settings.DatetimeBase
        };
        panel.Controls.Add(_dtDatetimeBase, 1, 7);

        group.Controls.Add(panel);
        return group;
    }

    private GroupBox CreateIdRangeGroup()
    {
        var group = new GroupBox { Text = "ID Range Overrides", AutoSize = true, Width = 1020 };
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true, Padding = new Padding(10) };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _numCsvIdStart = AddLabeledNumeric(panel, "CSVtoDB (100000)", _settings.CsvIdStart, 100000, 199999, 0);
        _numDbToDbIdStart = AddLabeledNumeric(panel, "DBtoDB (200000)", _settings.DbToDbIdStart, 200000, 299999, 1);
        _numDbToJsonIdStart = AddLabeledNumeric(panel, "DBtoJSON (300000)", _settings.DbToJsonIdStart, 300000, 399999, 2);
        _numSqlQueryIdStart = AddLabeledNumeric(panel, "SQLQuery (400000)", _settings.SqlQueryIdStart, 400000, 499999, 3);

        group.Controls.Add(panel);
        return group;
    }

    private GroupBox CreateFileSyncPrefixGroup()
    {
        var group = new GroupBox { Text = "FileSyncer Prefix Overrides", AutoSize = true, Width = 1020 };
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true, Padding = new Padding(10) };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _txtFileSyncUploadPrefix = AddLabeledTextBox(panel, "Upload prefix", _settings.FileSyncUploadPrefix, 0);
        _txtFileSyncDownloadPrefix = AddLabeledTextBox(panel, "Download prefix", _settings.FileSyncDownloadPrefix, 1);
        _txtFileSyncTwoWayPrefix = AddLabeledTextBox(panel, "TwoWay prefix", _settings.FileSyncTwoWayPrefix, 2);

        group.Controls.Add(panel);
        return group;
    }

    private TabPage CreateLogViewerTab()
    {
        var tab = new TabPage("Log Viewer");
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(16)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _logViewer = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            Font = new Font("Consolas", 10)
        };
        panel.Controls.Add(_logViewer, 0, 0);

        var btnRefresh = new Button { Text = "Refresh Log", AutoSize = true };
        btnRefresh.Click += (_, _) => RefreshLogViewer();
        panel.Controls.Add(btnRefresh, 0, 1);

        tab.Controls.Add(panel);
        return tab;
    }

    private static TextBox AddLabeledTextBox(TableLayoutPanel panel, string label, string value, int row)
    {
        AddLabel(panel, label, 0, row);
        var txt = new TextBox { Text = value, Anchor = AnchorStyles.Left | AnchorStyles.Right };
        panel.Controls.Add(txt, 1, row);
        return txt;
    }

    private static NumericUpDown AddLabeledNumeric(TableLayoutPanel panel, string label, int value, int min, int max, int row)
    {
        AddLabel(panel, label, 0, row);
        var num = new NumericUpDown { Minimum = min, Maximum = max, Value = Math.Max(min, Math.Min(max, value)), Width = 140 };
        panel.Controls.Add(num, 1, row);
        return num;
    }

    private static void AddLabel(TableLayoutPanel panel, string text, int column, int row)
    {
        panel.Controls.Add(new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left }, column, row);
    }

    private Button CreateActionButton(string text, string logMessage)
    {
        var button = new Button { Text = text, AutoSize = true };
        button.Click += (_, _) =>
        {
            _logService.LogInfo(logMessage);
            WriteStatus(logMessage);
            RefreshLogViewer();
        };
        return button;
    }

    private void SaveSettings()
    {
        if (_txtSqlConnectionString is null ||
            _txtOutputRootFolder is null ||
            _txtSyncFlagNotProcessed is null ||
            _txtSyncFlagProcessed is null ||
            _txtCsvIdPrefix is null ||
            _txtApiTestEndpoint is null ||
            _txtRemoteFileSharePath is null ||
            _dtDatetimeBase is null ||
            _numCsvIdStart is null ||
            _numDbToDbIdStart is null ||
            _numDbToJsonIdStart is null ||
            _numSqlQueryIdStart is null ||
            _txtFileSyncUploadPrefix is null ||
            _txtFileSyncDownloadPrefix is null ||
            _txtFileSyncTwoWayPrefix is null)
        {
            return;
        }

        _settings.SqlConnectionString = _txtSqlConnectionString.Text.Trim();
        _settings.OutputRootFolder = _txtOutputRootFolder.Text.Trim();
        _settings.SyncFlagNotProcessed = _txtSyncFlagNotProcessed.Text.Trim();
        _settings.SyncFlagProcessed = _txtSyncFlagProcessed.Text.Trim();
        _settings.CsvIdPrefix = _txtCsvIdPrefix.Text.Trim();
        _settings.ApiTestEndpoint = _txtApiTestEndpoint.Text.Trim();
        _settings.RemoteFileSharePath = _txtRemoteFileSharePath.Text.Trim();
        _settings.DatetimeBase = _dtDatetimeBase.Value;

        _settings.CsvIdStart = Decimal.ToInt32(_numCsvIdStart.Value);
        _settings.DbToDbIdStart = Decimal.ToInt32(_numDbToDbIdStart.Value);
        _settings.DbToJsonIdStart = Decimal.ToInt32(_numDbToJsonIdStart.Value);
        _settings.SqlQueryIdStart = Decimal.ToInt32(_numSqlQueryIdStart.Value);

        _settings.FileSyncUploadPrefix = _txtFileSyncUploadPrefix.Text.Trim();
        _settings.FileSyncDownloadPrefix = _txtFileSyncDownloadPrefix.Text.Trim();
        _settings.FileSyncTwoWayPrefix = _txtFileSyncTwoWayPrefix.Text.Trim();

        _configurationService.Save(_settings);
        _logService.LogInfo("Settings saved from UI");
        WriteStatus("Settings saved");

        if (_txtCsvOutputFolder is not null)
        {
            _txtCsvOutputFolder.Text = _settings.OutputRootFolder;
        }

        if (_dtCsvBase is not null)
        {
            _dtCsvBase.Value = _settings.DatetimeBase;
        }

        UpdateDashboardSummary();
        RefreshLogViewer();
    }

    private void TestSqlConnection()
    {
        if (_txtSqlConnectionString is null)
        {
            return;
        }

        try
        {
            using var connection = new SqlConnection(_txtSqlConnectionString.Text.Trim());
            connection.Open();
            var message = "SQL connection test succeeded.";
            _logService.LogInfo(message);
            WriteStatus(message);
            RefreshLogViewer();
            MessageBox.Show(message, "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            var message = "SQL connection test failed: " + ex.Message;
            _logService.LogError(message);
            WriteStatus("SQL connection failed");
            RefreshLogViewer();
            MessageBox.Show(message, "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RunGenerateAllPlaceholder()
    {
        _logService.LogInfo("Generate All requested from dashboard");
        _logService.LogInfo("CSVtoDB queued");
        _logService.LogInfo("DBtoDB queued");
        _logService.LogInfo("DBtoJSON queued");
        _logService.LogInfo("SQLQuery queued");
        _logService.LogInfo("ProgramExecution queued");
        _logService.LogInfo("FileSyncer queued");

        WriteStatus("Generate All completed (section triggers only)");
        RefreshLogViewer();
    }

    private void UpdateDashboardSummary()
    {
        if (_lblDashboardConfig is null)
        {
            return;
        }

        _lblDashboardConfig.Text =
            "SQL: " + _settings.SqlConnectionString + Environment.NewLine +
            "OutputRootFolder: " + _settings.OutputRootFolder + Environment.NewLine +
            "ApiTestEndpoint: " + _settings.ApiTestEndpoint + Environment.NewLine +
            "RemoteFileSharePath: " + _settings.RemoteFileSharePath + Environment.NewLine +
            "DatetimeBase: " + _settings.DatetimeBase.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
            "ID Starts => CSVtoDB: " + _settings.CsvIdStart +
            ", DBtoDB: " + _settings.DbToDbIdStart +
            ", DBtoJSON: " + _settings.DbToJsonIdStart +
            ", SQLQuery: " + _settings.SqlQueryIdStart;
    }

    private void WriteStatus(string message)
    {
        if (_statusLabel is not null)
        {
            _statusLabel.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " - " + message;
        }
    }

    private void RefreshLogViewer()
    {
        if (_logViewer is null)
        {
            return;
        }

        _logViewer.Text = _logService.ReadCurrentLog();
        _logViewer.SelectionStart = _logViewer.TextLength;
        _logViewer.ScrollToCaret();
    }

    private sealed record CsvRow(string RecordId, string MachineCode, string Status, DateTime CreatedAt, string Comment);
}
