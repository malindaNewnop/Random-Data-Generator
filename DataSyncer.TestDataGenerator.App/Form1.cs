using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text;
using DataSyncer.TestDataGenerator.App.Models;
using DataSyncer.TestDataGenerator.App.Services;
using Microsoft.Data.SqlClient;

namespace DataSyncer.TestDataGenerator.App;

public partial class Form1 : Form
{
    private const string DashboardPageKey = "dashboard";
    private const string CsvPageKey = "csv";
    private const string DbToDbPageKey = "dbtodb";
    private const string DbToJsonPageKey = "dbtojson";
    private const string SqlQueryPageKey = "sqlquery";
    private const string ProgramExecutionPageKey = "program";
    private const string FileSyncerPageKey = "filesyncer";
    private const string SettingsPageKey = "settings";
    private const string LogViewerPageKey = "logviewer";

    private static readonly Color ShellBackground = Color.FromArgb(243, 246, 249);
    private static readonly Color CardBackground = Color.White;
    private static readonly Color CardBorder = Color.FromArgb(221, 228, 236);
    private static readonly Color NavBackground = Color.FromArgb(15, 23, 42);
    private static readonly Color NavButtonBackground = Color.FromArgb(30, 41, 59);
    private static readonly Color NavButtonSelected = Color.FromArgb(14, 116, 144);
    private static readonly Color AccentColor = Color.FromArgb(14, 116, 144);
    private static readonly Color AccentSoft = Color.FromArgb(224, 242, 254);
    private static readonly Color SuccessColor = Color.FromArgb(5, 150, 105);
    private static readonly Color SuccessSoft = Color.FromArgb(220, 252, 231);
    private static readonly Color WarningColor = Color.FromArgb(180, 83, 9);
    private static readonly Color WarningSoft = Color.FromArgb(254, 243, 199);
    private static readonly Color DangerColor = Color.FromArgb(185, 28, 28);
    private static readonly Color DangerSoft = Color.FromArgb(254, 226, 226);
    private static readonly Color MutedText = Color.FromArgb(71, 85, 105);

    private readonly AppSettings _settings;
    private readonly ConfigurationService _configurationService;
    private readonly FileLogService _logService;
    private readonly ToolTip _toolTip = new();
    private readonly Dictionary<string, NavigationItem> _navButtons = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Control> _pages = new(StringComparer.Ordinal);

    private TableLayoutPanel? _rootLayout;
    private Panel? _heroBanner;
    private TableLayoutPanel? _heroLayout;
    private Control? _heroContent;
    private FlowLayoutPanel? _heroActions;
    private SplitContainer? _mainSplitContainer;
    private FlowLayoutPanel? _navigationStack;
    private Panel? _contentHost;
    private ToolStripStatusLabel? _statusLabel;
    private ToolStripStatusLabel? _currentPageLabel;
    private Label? _heroMetaLabel;

    private Label? _lblDashboardSql;
    private Label? _lblDashboardOutput;
    private Label? _lblDashboardRemote;
    private Label? _lblDashboardRuntime;
    private Label? _lblDashboardSafety;

    private Label? _lblSettingsState;
    private Label? _lblProductionWarning;

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

    private TextBox? _txtCsvOutputFolder;
    private DateTimePicker? _dtCsvBase;
    private ComboBox? _cmbCsvScenario;
    private Label? _lblCsvStatus;
    private ComboBox? _cmbCsvScheduleMode;
    private DateTimePicker? _dtCsvScheduleAt;
    private NumericUpDown? _numCsvScheduleIntervalMinutes;
    private ComboBox? _cmbCsvScheduleTarget;
    private Label? _lblCsvScheduleStatus;
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

    private RichTextBox? _logViewer;
    private readonly System.Windows.Forms.Timer _csvScheduleTimer = new() { Interval = 1000 };
    private DateTime? _csvScheduledRunAt;
    private CsvScheduleMode _armedCsvScheduleMode = CsvScheduleMode.OneTime;
    private TimeSpan? _csvScheduleInterval;
    private CsvAutomationExecutionContext? _csvAutomationExecutionContext;
    private int _csvAutomationRunSequence;
    private bool _settingsDirty;
    private bool _suppressSettingsChangedEvents;

    public Form1(
        AppSettings settings,
        ConfigurationService configurationService,
        FileLogService logService)
    {
        _settings = settings;
        _configurationService = configurationService;
        _logService = logService;

        InitializeComponent();

        DoubleBuffered = true;
        BackColor = ShellBackground;
        Font = new Font("Segoe UI", 10F);

        ConfigureTooltips();
        BuildUi();
        _csvScheduleTimer.Tick += (_, _) => HandleCsvScheduleTick();
        ApplyResponsiveShellLayout();
        LoadSettingsIntoInputs();
        UpdateDashboardSummary();
        UpdateSettingsWarningState();
        NavigateTo(DashboardPageKey);

        _logService.LogInfo("Application shell initialized");
        RefreshLogViewer();
        WriteStatus("Ready");
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        ApplyResponsiveShellLayout();
    }

    private void ConfigureTooltips()
    {
        _toolTip.AutoPopDelay = 10000;
        _toolTip.InitialDelay = 250;
        _toolTip.ReshowDelay = 150;
        _toolTip.ShowAlways = true;
    }

    private void BuildUi()
    {
        SuspendLayout();
        Controls.Clear();

        _rootLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = ShellBackground
        };
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 118F));
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _rootLayout.Controls.Add(CreateHeroBanner(), 0, 0);
        _rootLayout.Controls.Add(CreateMainBody(), 0, 1);
        _rootLayout.Controls.Add(CreateStatusBar(), 0, 2);

        Controls.Add(_rootLayout);
        ResumeLayout(true);
    }

    private Control CreateHeroBanner()
    {
        _heroBanner = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24, 20, 24, 20)
        };
        _heroBanner.Paint += (_, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new LinearGradientBrush(
                _heroBanner.ClientRectangle,
                Color.FromArgb(15, 23, 42),
                Color.FromArgb(14, 116, 144),
                LinearGradientMode.Horizontal);

            e.Graphics.FillRectangle(brush, _heroBanner.ClientRectangle);

            using var glowBrush = new SolidBrush(Color.FromArgb(28, Color.White));
            e.Graphics.FillEllipse(glowBrush, _heroBanner.Width - 180, -20, 220, 220);
            e.Graphics.FillEllipse(glowBrush, _heroBanner.Width - 340, 30, 90, 90);
        };

        _heroLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            BackColor = Color.Transparent
        };
        _heroLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        _heroLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var left = new TableLayoutPanel
        {
            AutoSize = true,
            ColumnCount = 1,
            BackColor = Color.Transparent
        };
        _heroContent = left;

        left.Controls.Add(new Label
        {
            Text = "DataSyncer Test Data Generator",
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 22F, FontStyle.Bold)
        });

        left.Controls.Add(new Label
        {
            Text = "A QA-friendly desktop workspace for configuring data generation, validating environments, and preparing each DataSyncer scenario with less manual setup.",
            AutoSize = true,
            MaximumSize = new Size(820, 0),
            ForeColor = Color.FromArgb(230, 241, 245),
            Margin = new Padding(0, 8, 0, 0)
        });

        _heroMetaLabel = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(840, 0),
            ForeColor = Color.FromArgb(201, 230, 237),
            Margin = new Padding(0, 10, 0, 0)
        };
        left.Controls.Add(_heroMetaLabel);

        _heroLayout.Controls.Add(left, 0, 0);

        _heroActions = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            BackColor = Color.Transparent,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };

        _heroActions.Controls.Add(CreateHeroActionButton("Open Output Root", (_, _) => OpenOrCreateOutputRoot()));
        _heroActions.Controls.Add(CreateHeroActionButton("Settings", (_, _) => NavigateTo(SettingsPageKey)));
        _heroActions.Controls.Add(CreateHeroActionButton("Refresh", (_, _) => RefreshAllUiState()));

        _heroLayout.Controls.Add(_heroActions, 1, 0);
        _heroBanner.Controls.Add(_heroLayout);

        return _heroBanner;
    }

    private Control CreateMainBody()
    {
        _mainSplitContainer = new SplitContainer
        {
            Dock = DockStyle.Fill,
            FixedPanel = FixedPanel.Panel1,
            IsSplitterFixed = false,
            SplitterDistance = 380,
            SplitterWidth = 1,
            BackColor = CardBorder
        };

        _mainSplitContainer.Panel1MinSize = 330;
        _mainSplitContainer.Panel1.BackColor = NavBackground;
        _mainSplitContainer.Panel2.BackColor = ShellBackground;
        _mainSplitContainer.Resize += (_, _) => ApplyResponsiveShellLayout();
        _mainSplitContainer.SplitterMoved += (_, _) => ResizeNavigationButtons();

        _mainSplitContainer.Panel1.Controls.Add(CreateNavigationPanel());

        _contentHost = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20),
            BackColor = ShellBackground
        };
        _mainSplitContainer.Panel2.Controls.Add(_contentHost);

        BuildPages();
        return _mainSplitContainer;
    }

    private Control CreateNavigationPanel()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = NavBackground,
            Padding = new Padding(16, 18, 16, 18)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = Color.Transparent
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var brandCard = new Panel
        {
            Dock = DockStyle.Top,
            BackColor = Color.FromArgb(30, 41, 59),
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 0, 14)
        };
        brandCard.Paint += DrawRoundedCard;

        var brandLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };

        brandLayout.Controls.Add(CreateBadgeLabel("PHASE 1", AccentSoft, AccentColor));
        brandLayout.Controls.Add(new Label
        {
            Text = "Implementation Workspace",
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold),
            Margin = new Padding(0, 12, 0, 0)
        });
        brandLayout.Controls.Add(new Label
        {
            Text = "Shell, settings, dashboard, and log viewer are ready for the next scenario generators.",
            AutoSize = true,
            MaximumSize = new Size(340, 0),
            ForeColor = Color.FromArgb(203, 213, 225),
            Margin = new Padding(0, 6, 0, 0)
        });

        brandCard.Controls.Add(brandLayout);
        layout.Controls.Add(brandCard, 0, 0);

        _navigationStack = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            BackColor = Color.Transparent
        };
        _navigationStack.Resize += (_, _) => ResizeNavigationButtons();

        AddNavigationButton(_navigationStack, DashboardPageKey, "Home / Dashboard", "Overview and quick actions");
        AddNavigationButton(_navigationStack, CsvPageKey, "CSVtoDB", "Scenario workspace");
        AddNavigationButton(_navigationStack, DbToDbPageKey, "DBtoDB", "Scenario workspace");
        AddNavigationButton(_navigationStack, DbToJsonPageKey, "DBtoJSON", "Scenario workspace");
        AddNavigationButton(_navigationStack, SqlQueryPageKey, "SQLQuery", "Scenario workspace");
        AddNavigationButton(_navigationStack, ProgramExecutionPageKey, "ProgramExecution", "Scenario workspace");
        AddNavigationButton(_navigationStack, FileSyncerPageKey, "FileSyncer", "Scenario workspace");
        AddNavigationButton(_navigationStack, SettingsPageKey, "Settings", "Configuration and safety checks");
        AddNavigationButton(_navigationStack, LogViewerPageKey, "Log Viewer", "Run history and trace output");

        layout.Controls.Add(_navigationStack, 0, 1);

        layout.Controls.Add(new Label
        {
            Text = "Responsive shell with scrolling support for smaller displays",
            AutoSize = true,
            ForeColor = Color.FromArgb(148, 163, 184),
            Margin = new Padding(4, 10, 0, 0)
        }, 0, 2);

        panel.Controls.Add(layout);
        return panel;
    }

    private void AddNavigationButton(FlowLayoutPanel host, string key, string title, string subtitle)
    {
        var card = new Panel
        {
            Width = 310,
            Height = 86,
            Margin = new Padding(0, 0, 0, 10),
            Cursor = Cursors.Hand,
            BackColor = NavButtonBackground,
            Padding = new Padding(16, 12, 16, 12)
        };
        card.Paint += DrawRoundedCard;

        var titleLabel = new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold),
            BackColor = Color.Transparent,
            MaximumSize = new Size(340, 0),
            Margin = new Padding(0)
        };

        var subtitleLabel = new Label
        {
            Text = subtitle,
            AutoSize = true,
            ForeColor = Color.FromArgb(191, 219, 254),
            Font = new Font("Segoe UI", 8.75F, FontStyle.Regular),
            BackColor = Color.Transparent,
            MaximumSize = new Size(340, 0),
            Margin = new Padding(0, 6, 0, 0)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.Transparent
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(titleLabel, 0, 0);
        layout.Controls.Add(subtitleLabel, 0, 1);
        card.Controls.Add(layout);

        void Navigate(object? _, EventArgs __) => NavigateTo(key);

        card.Click += Navigate;
        layout.Click += Navigate;
        titleLabel.Click += Navigate;
        subtitleLabel.Click += Navigate;

        _toolTip.SetToolTip(card, "Open " + title);

        _navButtons[key] = new NavigationItem(card, titleLabel, subtitleLabel);
        host.Controls.Add(card);
    }

    private Control CreateStatusBar()
    {
        var statusStrip = new StatusStrip
        {
            SizingGrip = false,
            BackColor = CardBackground,
            GripStyle = ToolStripGripStyle.Hidden
        };

        _currentPageLabel = new ToolStripStatusLabel("Dashboard")
        {
            ForeColor = MutedText
        };

        _statusLabel = new ToolStripStatusLabel("Ready")
        {
            Spring = true,
            TextAlign = ContentAlignment.MiddleRight
        };

        statusStrip.Items.Add(_currentPageLabel);
        statusStrip.Items.Add(new ToolStripStatusLabel(" | ") { ForeColor = CardBorder });
        statusStrip.Items.Add(_statusLabel);

        return statusStrip;
    }

    private void BuildPages()
    {
        if (_contentHost is null)
        {
            return;
        }

        AddPage(DashboardPageKey, CreateDashboardPage());
        AddPage(CsvPageKey, CreateCsvPage());

        AddPage(DbToDbPageKey, CreateDbToDbPage());
        AddPage(DbToJsonPageKey, CreateDbToJsonPage());
        AddPage(SqlQueryPageKey, CreateSqlQueryPage());
        AddPage(ProgramExecutionPageKey, CreateProgramExecutionPage());
        AddPage(FileSyncerPageKey, CreateFileSyncerPage());

        AddPage(SettingsPageKey, CreateSettingsPage());
        AddPage(LogViewerPageKey, CreateLogViewerPage());
    }

    private void AddPage(string key, Control page)
    {
        if (_contentHost is null)
        {
            return;
        }

        page.Dock = DockStyle.Fill;
        page.Visible = false;

        _pages[key] = page;
        _contentHost.Controls.Add(page);
    }

    private Control CreateDashboardPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateDashboardIntroCard());
        stack.Controls.Add(CreateDashboardGenerateAllCard());
        stack.Controls.Add(CreateDashboardHealthCard());
        stack.Controls.Add(CreateDashboardRoadmapCard());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateDashboardIntroCard()
    {
        var card = CreateCard("Workspace Snapshot", "Use the dashboard as the launch point for environment checks, settings maintenance, and the next scenario modules.", out var content);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 6, 0, 0)
        };

        actions.Controls.Add(CreateActionButton("Open Output Folder", (_, _) => OpenOrCreateOutputRoot(), true));
        actions.Controls.Add(CreateActionButton("Go To Settings", (_, _) => NavigateTo(SettingsPageKey)));
        actions.Controls.Add(CreateActionButton("Refresh Health", (_, _) => RefreshAllUiState()));
        actions.Controls.Add(CreateActionButton("Test SQL Connection", (_, _) => TestSqlConnection()));

        content.Controls.Add(actions);
        return card;
    }

    private Control CreateDashboardHealthCard()
    {
        var card = CreateCard("Environment Health", "These cards summarize the current saved configuration from appsettings.json and highlight anything that needs attention before scenario generation.", out var content);

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        var metricCards = new[]
        {
            CreateMetricCard("Primary SQL Connection", out _lblDashboardSql),
            CreateMetricCard("Output Workspace", out _lblDashboardOutput),
            CreateMetricCard("Remote File Share", out _lblDashboardRemote),
            CreateMetricCard("Runtime Defaults", out _lblDashboardRuntime)
        };

        ConfigureResponsiveCardGrid(card, grid, 920, metricCards);

        content.Controls.Add(grid);
        return card;
    }

    private Control CreateDashboardRoadmapCard()
    {
        var card = CreateCard("Current Delivery Progress", "We have the navigation shell and the settings module in place. The remaining scenario pages are scaffolded so we can implement each generator cleanly from here.", out var content);

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        AddRoadmapRow(table, 0, "Application shell and left navigation", "Ready", SuccessSoft, SuccessColor);
        AddRoadmapRow(table, 1, "Dashboard summary cards and quick actions", "Ready", SuccessSoft, SuccessColor);
        AddRoadmapRow(table, 2, "Settings editor with save / reload / test actions", "Ready", SuccessSoft, SuccessColor);
        AddRoadmapRow(table, 3, "Scenario-specific data generation panels", "Scaffolded", AccentSoft, AccentColor);
        AddRoadmapRow(table, 4, "Detailed generation services and validation flows", "Next", WarningSoft, WarningColor);

        _lblDashboardSafety = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            Margin = new Padding(0, 18, 0, 0)
        };

        content.Controls.Add(table);
        content.Controls.Add(_lblDashboardSafety);
        return card;
    }

    private static void AddRoadmapRow(TableLayoutPanel table, int row, string label, string status, Color backColor, Color foreColor)
    {
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.Controls.Add(new Label
        {
            Text = label,
            AutoSize = true,
            Margin = new Padding(0, 0, 12, 12)
        }, 0, row);

        table.Controls.Add(CreateBadgeLabel(status, backColor, foreColor), 1, row);
    }

    private Control CreateCsvPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateCsvOverviewCard());
        stack.Controls.Add(CreateCsvSharedControlsCard());
        stack.Controls.Add(CreateCsvAutomationCard());
        stack.Controls.Add(CreateCsvScenario1Card());
        stack.Controls.Add(CreateCsvScenario2Card());
        stack.Controls.Add(CreateCsvScenario3Card());
        stack.Controls.Add(CreateCsvScenario4Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateCsvOverviewCard()
    {
        var card = CreateCard("CSVtoDB Generator", "Generate deterministic CSV files for all four CSVtoDB scenarios from a single QA-friendly workspace.", out var content);

        var badges = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        badges.Controls.Add(CreateBadgeLabel("READY TO GENERATE", SuccessSoft, SuccessColor));
        badges.Controls.Add(CreateBadgeLabel("4 SCENARIOS", AccentSoft, AccentColor));
        badges.Controls.Add(CreateBadgeLabel("DETERMINISTIC OUTPUT", WarningSoft, WarningColor));
        content.Controls.Add(badges);

        content.Controls.Add(new Label
        {
            Text = "Use the shared controls below to choose the output folder and base timestamp, then generate a single scenario or all CSV scenarios in sequence.",
            AutoSize = true,
            MaximumSize = new Size(1040, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 4, 0, 0)
        });

        return card;
    }

    private Control CreateCsvSharedControlsCard()
    {
        var card = CreateCard("Shared CSV Controls", "Common settings for the CSV generators. These values are pre-filled from the main settings screen but remain editable here for one-off runs.", out var content);

        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Output folder path", 0);
        _txtCsvOutputFolder = new TextBox
        {
            Text = _settings.OutputRootFolder,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "CSV output folder path"
        };
        _toolTip.SetToolTip(_txtCsvOutputFolder, "Folder where the CSVtoDB files will be created.");
        grid.Controls.Add(_txtCsvOutputFolder, 1, 0);

        var outputActions = CreateInlineActionPanel();
        outputActions.Controls.Add(CreateMiniButton("Browse", (_, _) => BrowseForCsvOutputFolder()));
        outputActions.Controls.Add(CreateMiniButton("Open", (_, _) => OpenCsvOutputFolder()));
        grid.Controls.Add(outputActions, 2, 0);

        AddLabelCell(grid, "Date/time base", 1);
        _dtCsvBase = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Value = _settings.DatetimeBase,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "CSV base datetime"
        };
        _toolTip.SetToolTip(_dtCsvBase, "Base timestamp for generated CSV rows.");
        grid.Controls.Add(_dtCsvBase, 1, 1);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 1);

        AddLabelCell(grid, "Selected scenario", 2);
        _cmbCsvScenario = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Width = 320,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "Selected CSV scenario"
        };
        _cmbCsvScenario.Items.AddRange(new object[]
        {
            "Scenario 1 - Happy Path CSV Import",
            "Scenario 2 - Optional Column Missing",
            "Scenario 3 - Duplicate Prevention",
            "Scenario 4 - Scheduled Recurring Run"
        });
        _cmbCsvScenario.SelectedIndex = 0;
        _toolTip.SetToolTip(_cmbCsvScenario, "Choose the scenario for the Generate Selected Scenario button.");
        grid.Controls.Add(_cmbCsvScenario, 1, 2);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 2);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate Selected Scenario", (_, _) => GenerateSelectedCsvScenario(), true));
        actions.Controls.Add(CreateActionButton("Generate All CSV Scenarios", (_, _) => GenerateAllCsvScenarios()));
        actions.Controls.Add(CreateActionButton("Clear Output Folder", (_, _) => ClearCsvOutputFolder()));
        actions.Controls.Add(CreateActionButton("Open Output Folder", (_, _) => OpenCsvOutputFolder()));
        content.Controls.Add(actions);

        _lblCsvStatus = new Label
        {
            Text = "Last generated: (none)",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblCsvStatus);

        return card;
    }

    private Control CreateCsvAutomationCard()
    {
        var card = CreateCard("CSV Automation Timer", "Keep this application open while DataSyncer is running and arm one-time, daily, or interval-based CSV generation. For example, set Every N minutes to 1 to create fresh data every minute automatically.", out var content);

        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Automation mode", 0);
        _cmbCsvScheduleMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "CSV automation mode"
        };
        _cmbCsvScheduleMode.Items.AddRange(new object[]
        {
            "Run Once At Selected Time",
            "Repeat Daily At Selected Time",
            "Repeat Every N Minutes"
        });
        _cmbCsvScheduleMode.SelectedIndex = 0;
        _cmbCsvScheduleMode.SelectedIndexChanged += (_, _) => UpdateCsvAutomationControlState();
        _toolTip.SetToolTip(_cmbCsvScheduleMode, "Choose whether the CSV scheduler runs once, repeats every day, or repeats every N minutes.");
        grid.Controls.Add(_cmbCsvScheduleMode, 1, 0);

        AddLabelCell(grid, "Run at", 1);
        _dtCsvScheduleAt = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Value = DateTime.Now.AddMinutes(5),
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "CSV scheduled run time"
        };
        _toolTip.SetToolTip(_dtCsvScheduleAt, "Date and time used for one-time and daily CSV automation.");
        grid.Controls.Add(_dtCsvScheduleAt, 1, 1);

        AddLabelCell(grid, "Every N minutes", 2);
        _numCsvScheduleIntervalMinutes = CreateCsvNumeric(1, 1440, 1, "Interval in minutes for recurring CSV generation.");
        _numCsvScheduleIntervalMinutes.AccessibleName = "CSV automation interval minutes";
        grid.Controls.Add(_numCsvScheduleIntervalMinutes, 1, 2);

        AddLabelCell(grid, "Timer target", 3);
        _cmbCsvScheduleTarget = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "CSV automation target"
        };
        _cmbCsvScheduleTarget.Items.AddRange(new object[]
        {
            "Generate Selected Scenario",
            "Generate All CSV Scenarios"
        });
        _cmbCsvScheduleTarget.SelectedIndex = 0;
        _toolTip.SetToolTip(_cmbCsvScheduleTarget, "Choose whether the timer runs the selected scenario or all CSV scenarios.");
        grid.Controls.Add(_cmbCsvScheduleTarget, 1, 3);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Arm Timer", (_, _) => ArmCsvSchedule(), true));
        actions.Controls.Add(CreateActionButton("Cancel Timer", (_, _) => CancelCsvSchedule()));
        actions.Controls.Add(CreateActionButton("Run Target Now", (_, _) => RunCsvScheduledTarget()));
        content.Controls.Add(actions);

        _lblCsvScheduleStatus = new Label
        {
            Text = "Scheduler idle. Choose a mode, then arm the timer.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblCsvScheduleStatus);

        content.Controls.Add(CreateInfoNote(
            "Parallel run note",
            "This timer works while the generator app is open, so you can leave it running next to DataSyncer and let it create fresh CSV files automatically at a scheduled time or recurring interval.",
            AccentSoft,
            AccentColor));

        UpdateCsvAutomationControlState();
        return card;
    }

    private Control CreateCsvScenario1Card()
    {
        var card = CreateCard("Scenario 1 - Happy Path CSV Import", "Creates a complete CSV with the optional Comment column included and deterministic row values.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Output filename", 0);
        _txtCsvS1OutputFilename = CreateCsvTextBox("csv_happy_20260406_090000.csv", "Scenario 1 output filename.");
        grid.Controls.Add(_txtCsvS1OutputFilename, 1, 0);

        AddLabelCell(grid, "Row count (5-500)", 1);
        _numCsvS1RowCount = CreateCsvNumeric(5, 500, 20, "Scenario 1 row count.");
        grid.Controls.Add(_numCsvS1RowCount, 1, 1);

        AddLabelCell(grid, "Key start", 2);
        _txtCsvS1KeyStart = CreateCsvTextBox(BuildDailySerialSeed(_settings.DatetimeBase, 1), "Starting Serial value, for example 202604060001.");
        grid.Controls.Add(_txtCsvS1KeyStart, 1, 2);

        AddLabelCell(grid, "Include Comment column", 3);
        _chkCsvS1IncludeComment = new CheckBox
        {
            Checked = true,
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 12),
            AccessibleName = "Include Comment column"
        };
        _toolTip.SetToolTip(_chkCsvS1IncludeComment, "Include the Comment column in Scenario 1 output.");
        grid.Controls.Add(_chkCsvS1IncludeComment, 1, 3);

        AddLabelCell(grid, "MachineCode values", 4);
        _txtCsvS1MachineCodes = CreateCsvTextBox("UPR_HSG", "Comma-separated MachineCode values.");
        grid.Controls.Add(_txtCsvS1MachineCodes, 1, 4);

        AddLabelCell(grid, "Status values", 5);
        _txtCsvS1Statuses = CreateCsvTextBox("OK", "Comma-separated OK/NG values.");
        grid.Controls.Add(_txtCsvS1Statuses, 1, 5);

        content.Controls.Add(grid);
        content.Controls.Add(CreateInfoNote(
            "Pre-state note",
            "Destination must not already contain these RecordId values. The monitored input folder should contain only this file.",
            AccentSoft,
            AccentColor));

        return card;
    }

    private Control CreateCsvScenario2Card()
    {
        var card = CreateCard("Scenario 2 - Optional Column Missing", "Creates a CSV that omits the Comment column so QA can verify optional-header handling.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Output filename", 0);
        _txtCsvS2OutputFilename = CreateCsvTextBox("csv_optional_missing_20260406_091000.csv", "Scenario 2 output filename.");
        grid.Controls.Add(_txtCsvS2OutputFilename, 1, 0);

        AddLabelCell(grid, "Row count (2-100)", 1);
        _numCsvS2RowCount = CreateCsvNumeric(2, 100, 5, "Scenario 2 row count.");
        grid.Controls.Add(_numCsvS2RowCount, 1, 1);

        AddLabelCell(grid, "Key start", 2);
        _txtCsvS2KeyStart = CreateCsvTextBox(BuildDailySerialSeed(_settings.DatetimeBase, 21), "Starting Serial value, for example 202604060021.");
        grid.Controls.Add(_txtCsvS2KeyStart, 1, 2);

        AddLabelCell(grid, "Omit Comment column", 3);
        _chkCsvS2OmitComment = new CheckBox
        {
            Checked = true,
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 12),
            AccessibleName = "Omit Comment column"
        };
        _toolTip.SetToolTip(_chkCsvS2OmitComment, "When checked, the Comment column is omitted from the generated CSV.");
        grid.Controls.Add(_chkCsvS2OmitComment, 1, 3);

        content.Controls.Add(grid);
        content.Controls.Add(CreateInfoNote(
            "Pre-state note",
            "Mapping must be configured so the missing CSV header is optional, not required.",
            WarningSoft,
            WarningColor));

        return card;
    }

    private Control CreateCsvScenario3Card()
    {
        var card = CreateCard("Scenario 3 - Duplicate Prevention", "Creates a seed file and optionally a variant file that mixes duplicate and new keys.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Seed filename", 0);
        _txtCsvS3SeedFilename = CreateCsvTextBox("csv_duplicate_seed_20260406_092000.csv", "Seed file name for the initial run.");
        grid.Controls.Add(_txtCsvS3SeedFilename, 1, 0);

        AddLabelCell(grid, "Row count", 1);
        _numCsvS3RowCount = CreateCsvNumeric(1, 500, 10, "Scenario 3 seed row count.");
        grid.Controls.Add(_numCsvS3RowCount, 1, 1);

        AddLabelCell(grid, "Key start", 2);
        _txtCsvS3KeyStart = CreateCsvTextBox(BuildDailySerialSeed(_settings.DatetimeBase, 1001), "Starting Serial value for the seed file.");
        grid.Controls.Add(_txtCsvS3KeyStart, 1, 2);

        AddLabelCell(grid, "Generate variant file", 3);
        _chkCsvS3GenerateVariant = new CheckBox
        {
            Checked = false,
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 12),
            AccessibleName = "Generate variant file"
        };
        _chkCsvS3GenerateVariant.CheckedChanged += (_, _) => ToggleScenario3VariantState();
        _toolTip.SetToolTip(_chkCsvS3GenerateVariant, "Create a second file with 8 duplicate keys and 2 new keys.");
        grid.Controls.Add(_chkCsvS3GenerateVariant, 1, 3);

        AddLabelCell(grid, "Variant filename", 4);
        _txtCsvS3VariantFilename = CreateCsvTextBox("csv_duplicate_variant.csv", "Variant file name used when Generate variant file is checked.");
        _txtCsvS3VariantFilename.Enabled = false;
        grid.Controls.Add(_txtCsvS3VariantFilename, 1, 4);

        content.Controls.Add(grid);
        content.Controls.Add(CreateInfoNote(
            "Variant behavior",
            "When enabled, the variant file contains 8 duplicate keys from the seed file and 2 new keys for duplicate-prevention testing.",
            AccentSoft,
            AccentColor));

        return card;
    }

    private Control CreateCsvScenario4Card()
    {
        var card = CreateCard("Scenario 4 - Scheduled Recurring Run", "Creates the three scheduled files plus the file_A re-drop copy in one action.", out var content);

        var summary = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        summary.Controls.Add(CreateScenario4Line("file_A_yyyyMMdd_100000.csv", "CSV-102001 to CSV-102010", "Timestamp suffix 100000"));
        summary.Controls.Add(CreateScenario4Line("file_B_yyyyMMdd_100100.csv", "CSV-102011 to CSV-102020", "Timestamp suffix 100100"));
        summary.Controls.Add(CreateScenario4Line("file_C_yyyyMMdd_100300.csv", "CSV-102021 to CSV-102030", "Timestamp suffix 100300"));
        summary.Controls.Add(CreateScenario4Line("file_A_REDROP_yyyyMMdd_100300.csv", "Same keys as file_A", "Identical content to file_A"));
        content.Controls.Add(summary);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 12, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate All 4 Files", (_, _) => GenerateCsvScenario4(), true));
        actions.Controls.Add(CreateActionButton("Open Output Folder", (_, _) => OpenCsvOutputFolder()));
        content.Controls.Add(actions);

        return card;
    }

    private static TableLayoutPanel CreateCsvFormGrid(bool includeActionColumn = false)
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        grid.ColumnCount = includeActionColumn ? 3 : 2;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, includeActionColumn ? 220F : 240F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        if (includeActionColumn)
        {
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        }

        return grid;
    }

    private TextBox CreateCsvTextBox(string value, string tooltip)
    {
        var textBox = new TextBox
        {
            Text = value,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12)
        };
        _toolTip.SetToolTip(textBox, tooltip);
        return textBox;
    }

    private NumericUpDown CreateCsvNumeric(int minimum, int maximum, int value, string tooltip)
    {
        var numeric = new NumericUpDown
        {
            Minimum = minimum,
            Maximum = maximum,
            Value = Math.Max(minimum, Math.Min(maximum, value)),
            Width = 220,
            Margin = new Padding(0, 0, 12, 12)
        };
        _toolTip.SetToolTip(numeric, tooltip);
        return numeric;
    }

    private Control CreateInfoNote(string title, string body, Color background, Color foreground)
    {
        var panel = new Panel
        {
            BackColor = background,
            Dock = DockStyle.Top,
            Padding = new Padding(14),
            Margin = new Padding(0, 8, 0, 0)
        };
        panel.Paint += DrawRoundedCard;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = foreground,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
        });
        layout.Controls.Add(new Label
        {
            Text = body,
            AutoSize = true,
            MaximumSize = new Size(1000, 0),
            ForeColor = foreground,
            Margin = new Padding(0, 6, 0, 0)
        });

        panel.Controls.Add(layout);
        return panel;
    }

    private static Control CreateScenario4Line(string fileName, string keyRange, string note)
    {
        return new Label
        {
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            Text = fileName + " | " + keyRange + " | " + note,
            Margin = new Padding(0, 0, 0, 8)
        };
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
                MessageBox.Show("Please select a CSV scenario first.", "CSVtoDB", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

    private void ArmCsvSchedule()
    {
        if (_dtCsvScheduleAt is null ||
            _cmbCsvScheduleMode is null ||
            _numCsvScheduleIntervalMinutes is null ||
            _cmbCsvScheduleTarget is null)
        {
            return;
        }

        var mode = GetSelectedCsvScheduleMode();
        var scheduledAt = mode switch
        {
            CsvScheduleMode.Daily => ResolveNextCsvScheduledRun(_dtCsvScheduleAt.Value),
            CsvScheduleMode.EveryNMinutes => DateTime.Now.AddMinutes(decimal.ToDouble(_numCsvScheduleIntervalMinutes.Value)),
            _ => _dtCsvScheduleAt.Value
        };

        if (mode == CsvScheduleMode.OneTime && scheduledAt <= DateTime.Now)
        {
            MessageBox.Show(
                "Pick a future time for the CSV scheduler, or switch to a recurring mode to let the app calculate the next run automatically.",
                "CSV Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _armedCsvScheduleMode = mode;
        _csvScheduleInterval = mode == CsvScheduleMode.EveryNMinutes
            ? TimeSpan.FromMinutes(decimal.ToDouble(_numCsvScheduleIntervalMinutes.Value))
            : null;
        _csvScheduledRunAt = scheduledAt;
        _csvScheduleTimer.Start();
        UpdateCsvScheduleStatus();

        _logService.LogInfo(
            "CSV scheduler armed for " +
            scheduledAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetCsvScheduleTargetLabel() +
            " | " + GetArmedCsvScheduleModeLabel());
        RefreshLogViewer();
        WriteStatus("CSV scheduler armed");
    }

    private void CancelCsvSchedule()
    {
        _csvScheduleTimer.Stop();
        _csvScheduledRunAt = null;
        _csvScheduleInterval = null;
        _armedCsvScheduleMode = CsvScheduleMode.OneTime;
        UpdateCsvScheduleStatus();
        _logService.LogInfo("CSV scheduler canceled");
        RefreshLogViewer();
        WriteStatus("CSV scheduler canceled");
    }

    private void HandleCsvScheduleTick()
    {
        if (!_csvScheduledRunAt.HasValue)
        {
            _csvScheduleTimer.Stop();
            UpdateCsvScheduleStatus();
            return;
        }

        if (DateTime.Now < _csvScheduledRunAt.Value)
        {
            UpdateCsvScheduleStatus();
            return;
        }

        RunCsvScheduledTarget();

        if (_armedCsvScheduleMode == CsvScheduleMode.Daily)
        {
            _csvScheduledRunAt = ResolveNextCsvScheduledRun(_csvScheduledRunAt.Value.AddDays(1));
            _csvScheduleTimer.Start();
        }
        else if (_armedCsvScheduleMode == CsvScheduleMode.EveryNMinutes && _csvScheduleInterval.HasValue)
        {
            _csvScheduledRunAt = ResolveNextCsvRecurringRun(_csvScheduledRunAt.Value, _csvScheduleInterval.Value);
            _csvScheduleTimer.Start();
        }
        else
        {
            _csvScheduleTimer.Stop();
            _csvScheduledRunAt = null;
            _csvScheduleInterval = null;
        }

        UpdateCsvScheduleStatus();
    }

    private void RunCsvScheduledTarget()
    {
        _csvAutomationExecutionContext = CreateCsvAutomationExecutionContext();
        try
        {
            var targetLabel = GetCsvScheduleTargetLabel();
            if (_cmbCsvScheduleTarget?.SelectedIndex == 1)
            {
                GenerateAllCsvScenarios();
            }
            else
            {
                GenerateSelectedCsvScenario();
            }

            _logService.LogSuccess("CSV scheduler executed: " + targetLabel);
            RefreshLogViewer();
            WriteStatus("CSV scheduler executed");
        }
        catch (Exception ex)
        {
            _logService.LogError("CSV scheduler failed: " + ex.Message);
            RefreshLogViewer();
            WriteStatus("CSV scheduler failed");

            MessageBox.Show(
                "Scheduled CSV generation failed." + Environment.NewLine + Environment.NewLine + ex.Message,
                "CSV Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            _csvAutomationExecutionContext = null;
        }
    }

    private void UpdateCsvScheduleStatus()
    {
        if (_lblCsvScheduleStatus is null)
        {
            return;
        }

        if (!_csvScheduledRunAt.HasValue)
        {
            _lblCsvScheduleStatus.Text = "Scheduler idle. Choose a mode, then arm the timer.";
            _lblCsvScheduleStatus.ForeColor = MutedText;
            return;
        }

        var remaining = _csvScheduledRunAt.Value - DateTime.Now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        _lblCsvScheduleStatus.Text =
            "Armed for " + _csvScheduledRunAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetCsvScheduleTargetLabel() +
            " | " + GetArmedCsvScheduleModeLabel() +
            Environment.NewLine +
            "Time remaining: " + remaining.ToString(@"dd\.hh\:mm\:ss", CultureInfo.InvariantCulture);
        _lblCsvScheduleStatus.ForeColor = AccentColor;
    }

    private string GetCsvScheduleTargetLabel()
    {
        return _cmbCsvScheduleTarget?.SelectedIndex == 1
            ? "Generate All CSV Scenarios"
            : "Generate Selected Scenario";
    }

    private void UpdateCsvAutomationControlState()
    {
        var mode = GetSelectedCsvScheduleMode();
        if (_dtCsvScheduleAt is not null)
        {
            _dtCsvScheduleAt.Enabled = mode != CsvScheduleMode.EveryNMinutes;
        }

        if (_numCsvScheduleIntervalMinutes is not null)
        {
            _numCsvScheduleIntervalMinutes.Enabled = mode == CsvScheduleMode.EveryNMinutes;
        }
    }

    private CsvScheduleMode GetSelectedCsvScheduleMode()
    {
        return _cmbCsvScheduleMode?.SelectedIndex switch
        {
            1 => CsvScheduleMode.Daily,
            2 => CsvScheduleMode.EveryNMinutes,
            _ => CsvScheduleMode.OneTime
        };
    }

    private string GetArmedCsvScheduleModeLabel()
    {
        return _armedCsvScheduleMode switch
        {
            CsvScheduleMode.Daily => "repeats daily",
            CsvScheduleMode.EveryNMinutes when _csvScheduleInterval.HasValue =>
                "every " + FormatCsvIntervalLabel(_csvScheduleInterval.Value),
            _ => "one-time"
        };
    }

    private static DateTime ResolveNextCsvScheduledRun(DateTime requestedTime)
    {
        while (requestedTime <= DateTime.Now)
        {
            requestedTime = requestedTime.AddDays(1);
        }

        return requestedTime;
    }

    private static DateTime ResolveNextCsvRecurringRun(DateTime previousScheduledAt, TimeSpan interval)
    {
        if (interval <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(interval), "Recurring CSV interval must be greater than zero.");
        }

        var nextRun = previousScheduledAt.Add(interval);
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.Add(interval);
        }

        return nextRun;
    }

    private CsvAutomationExecutionContext CreateCsvAutomationExecutionContext()
    {
        _csvAutomationRunSequence++;
        if (_csvAutomationRunSequence <= 0)
        {
            _csvAutomationRunSequence = 1;
        }

        return new CsvAutomationExecutionContext(DateTime.Now, _csvAutomationRunSequence);
    }

    private static string FormatCsvIntervalLabel(TimeSpan interval)
    {
        var totalMinutes = Math.Max(1, (int)Math.Round(interval.TotalMinutes, MidpointRounding.AwayFromZero));
        return totalMinutes.ToString(CultureInfo.InvariantCulture) + (totalMinutes == 1 ? " minute" : " minutes");
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
        if (!ValidateCsvFileName(outputFileName))
        {
            MessageBox.Show("Scenario 1 output filename must end in .csv and cannot include a folder path.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!TryParseKeyStart(_txtCsvS1KeyStart.Text.Trim(), out var prefix, out var startNumber))
        {
            MessageBox.Show("Scenario 1 key start must end in digits, for example 202604060001 or CSV-100001.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var machineCodes = ParseCommaSeparatedValues(_txtCsvS1MachineCodes.Text);
        var statuses = ParseCommaSeparatedValues(_txtCsvS1Statuses.Text);
        if (machineCodes.Count == 0 || statuses.Count == 0)
        {
            MessageBox.Show("Scenario 1 MachineCode and Status lists must contain at least one comma-separated value.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rows = BuildCsvRows(
            prefix,
            startNumber,
            Decimal.ToInt32(_numCsvS1RowCount.Value),
            machineCodes,
            statuses,
            _chkCsvS1IncludeComment.Checked,
            GetCsvBaseTime());

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
        if (!ValidateCsvFileName(outputFileName))
        {
            MessageBox.Show("Scenario 2 output filename must end in .csv and cannot include a folder path.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!TryParseKeyStart(_txtCsvS2KeyStart.Text.Trim(), out var prefix, out var startNumber))
        {
            MessageBox.Show("Scenario 2 key start must end in digits, for example 202604060021 or CSV-100021.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var includeComment = !_chkCsvS2OmitComment.Checked;
        var rows = BuildCsvRows(
            prefix,
            startNumber,
            Decimal.ToInt32(_numCsvS2RowCount.Value),
            new List<string> { "UPR_HSG" },
            new List<string> { "OK" },
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
        if (!ValidateCsvFileName(seedFileName))
        {
            MessageBox.Show("Scenario 3 seed filename must end in .csv and cannot include a folder path.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!TryParseKeyStart(_txtCsvS3KeyStart.Text.Trim(), out var prefix, out var startNumber))
        {
            MessageBox.Show("Scenario 3 key start must end in digits, for example 202604061001 or CSV-101001.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rowCount = Decimal.ToInt32(_numCsvS3RowCount.Value);
        var baseTime = GetCsvBaseTime();
        var seedRows = BuildCsvRows(
            prefix,
            startNumber,
            rowCount,
            new List<string> { "UPR_HSG" },
            new List<string> { "OK" },
            includeComment: true,
            baseTime);

        WriteCsvToOutput(seedFileName, seedRows, includeComment: true, "CSVtoDB Scenario 3 (seed)");

        if (!_chkCsvS3GenerateVariant.Checked)
        {
            return;
        }

        var variantFileName = _txtCsvS3VariantFilename.Text.Trim();
        if (!ValidateCsvFileName(variantFileName))
        {
            MessageBox.Show("Scenario 3 variant filename must end in .csv and cannot include a folder path.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var duplicateCount = Math.Min(8, seedRows.Count);
        var variantRows = new List<CsvMeasurementRow>(duplicateCount + 2);

        for (var i = 0; i < duplicateCount; i++)
        {
            variantRows.Add(seedRows[i]);
        }

        var nextId = startNumber + rowCount;
        for (var i = 0; i < 2; i++)
        {
            variantRows.Add(BuildCsvMeasurementRow(
                prefix,
                nextId + i,
                "UPR_HSG",
                "OK",
                baseTime.AddMinutes(duplicateCount + i),
                includeComment: true,
                duplicateCount + i));
        }

        WriteCsvToOutput(variantFileName, variantRows, includeComment: true, "CSVtoDB Scenario 3 (variant)");
    }

    private void GenerateCsvScenario4()
    {
        var baseDate = GetCsvBaseTime().Date;
        var outputFolder = EnsureCsvOutputFolder();
        if (outputFolder is null)
        {
            return;
        }

        var fileAName = "file_A_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100000.csv";
        var fileBName = "file_B_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100100.csv";
        var fileCName = "file_C_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100300.csv";
        var fileRedropName = "file_A_REDROP_" + baseDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + "_100300.csv";

        var machineCodes = new List<string> { "UPR_HSG" };
        var statuses = new List<string> { "OK" };

        var rowsA = BuildCsvRows(string.Empty, BuildDailySerialSeedNumber(baseDate, 2001), 10, machineCodes, statuses, includeComment: true, baseDate.AddHours(10));
        var rowsB = BuildCsvRows(string.Empty, BuildDailySerialSeedNumber(baseDate, 2011), 10, machineCodes, statuses, includeComment: true, baseDate.AddHours(10).AddMinutes(1));
        var rowsC = BuildCsvRows(string.Empty, BuildDailySerialSeedNumber(baseDate, 2021), 10, machineCodes, statuses, includeComment: true, baseDate.AddHours(10).AddMinutes(3));

        WriteCsvToOutput(fileAName, rowsA, includeComment: true, "CSVtoDB Scenario 4 (file A)");
        WriteCsvToOutput(fileBName, rowsB, includeComment: true, "CSVtoDB Scenario 4 (file B)");
        WriteCsvToOutput(fileCName, rowsC, includeComment: true, "CSVtoDB Scenario 4 (file C)");
        WriteCsvToOutput(fileRedropName, rowsA, includeComment: true, "CSVtoDB Scenario 4 (file A redrop)");

        if (_csvAutomationExecutionContext is null)
        {
            var result = MessageBox.Show(
                "Scenario 4 files were created in:" + Environment.NewLine + outputFolder + Environment.NewLine + Environment.NewLine + "Open the folder now?",
                "CSVtoDB Scenario 4",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                OpenPath(outputFolder);
            }
        }
    }

    private void ToggleScenario3VariantState()
    {
        if (_chkCsvS3GenerateVariant is not null && _txtCsvS3VariantFilename is not null)
        {
            _txtCsvS3VariantFilename.Enabled = _chkCsvS3GenerateVariant.Checked;
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
            "Delete all .csv files from this folder?" + Environment.NewLine + outputFolder,
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

    private List<CsvMeasurementRow> BuildCsvRows(
        string keyPrefix,
        long startNumber,
        int rowCount,
        IReadOnlyList<string> machineCodes,
        IReadOnlyList<string> statuses,
        bool includeComment,
        DateTime baseTime)
    {
        var rows = new List<CsvMeasurementRow>(rowCount);
        for (var i = 0; i < rowCount; i++)
        {
            var machineCode = machineCodes[i % machineCodes.Count];
            var status = statuses[i % statuses.Count];
            var rowTime = baseTime.AddMinutes(i);
            rows.Add(BuildCsvMeasurementRow(
                keyPrefix,
                startNumber + i,
                machineCode,
                status,
                rowTime,
                includeComment,
                i));
        }

        return rows;
    }

    private CsvMeasurementRow BuildCsvMeasurementRow(
        string keyPrefix,
        long serialNumber,
        string machineCode,
        string status,
        DateTime measuredAt,
        bool includeComment,
        int rowIndex)
    {
        var profile = BuildMeasurementProfile(rowIndex, status);
        var resolvedPrefix = keyPrefix;
        var resolvedSerialNumber = serialNumber;
        if (_csvAutomationExecutionContext is not null)
        {
            var runSuffix =
                _csvAutomationExecutionContext.RunStamp.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) +
                "-R" +
                _csvAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture) +
                "-";

            if (string.IsNullOrWhiteSpace(keyPrefix))
            {
                resolvedSerialNumber = (serialNumber * 1000L) + _csvAutomationExecutionContext.RunSequence;
            }
            else
            {
                resolvedPrefix = keyPrefix + runSuffix;
            }
        }

        var serial = resolvedPrefix + resolvedSerialNumber.ToString(CultureInfo.InvariantCulture);
        var workDate = measuredAt.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
        var measuredAtStamp = measuredAt.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
        var comment = includeComment
            ? "Sample row " + (rowIndex + 1).ToString(CultureInfo.InvariantCulture)
            : string.Empty;

        return new CsvMeasurementRow(
            serial,
            workDate,
            measuredAtStamp,
            machineCode,
            status,
            profile.P1,
            profile.P2,
            profile.P3,
            profile.P4,
            profile.Beta1,
            profile.Beta2,
            profile.Beta3,
            profile.Beta4,
            profile.Beta5,
            profile.Beta6,
            profile.BetaRange,
            profile.BetaAverage,
            comment);
    }

    private static CsvMeasurementProfile BuildMeasurementProfile(int rowIndex, string status)
    {
        var pTemplates = new[]
        {
            new[] { 13.476m, 13.475m, 10.975m, 10.978m },
            new[] { 13.476m, 13.474m, 10.973m, 10.977m },
            new[] { 13.477m, 13.477m, 10.967m, 10.978m },
            new[] { 13.477m, 13.478m, 10.973m, 10.978m },
            new[] { 13.475m, 13.476m, 10.971m, 10.977m },
            new[] { 13.478m, 13.476m, 10.972m, 10.979m }
        };

        var betaTemplates = new[]
        {
            new[] { 38.022m, 38.026m, 38.029m, 38.020m, 38.029m, 38.032m },
            new[] { 37.998m, 38.018m, 37.993m, 38.000m, 38.004m, 37.996m },
            new[] { 37.995m, 38.003m, 37.972m, 37.973m, 38.003m, 37.971m },
            new[] { 38.018m, 38.016m, 37.997m, 37.989m, 37.994m, 37.990m },
            new[] { 38.004m, 38.008m, 38.011m, 38.007m, 38.010m, 38.013m },
            new[] { 37.986m, 37.992m, 37.989m, 37.985m, 37.994m, 37.988m }
        };

        var pValues = pTemplates[rowIndex % pTemplates.Length].ToArray();
        var betaValues = betaTemplates[rowIndex % betaTemplates.Length].ToArray();

        if (status.Contains("NG", StringComparison.OrdinalIgnoreCase))
        {
            pValues[2] = RoundCsvValue(pValues[2] + 0.050m);
            betaValues[^1] = RoundCsvValue(betaValues[^1] + 0.080m);
        }

        var betaRange = RoundCsvValue(betaValues.Max() - betaValues.Min());
        var betaAverage = RoundCsvValue(betaValues.Average());

        return new CsvMeasurementProfile(
            pValues[0],
            pValues[1],
            pValues[2],
            pValues[3],
            betaValues[0],
            betaValues[1],
            betaValues[2],
            betaValues[3],
            betaValues[4],
            betaValues[5],
            betaRange,
            betaAverage);
    }

    private void WriteCsvToOutput(string fileName, List<CsvMeasurementRow> rows, bool includeComment, string contextLabel)
    {
        var outputFolder = EnsureCsvOutputFolder();
        if (outputFolder is null)
        {
            return;
        }

        var resolvedFileName = ResolveCsvOutputFileName(fileName);
        var fullPath = Path.Combine(outputFolder, resolvedFileName);
        File.WriteAllText(fullPath, BuildCsvContent(rows, includeComment), new UTF8Encoding(true));

        _logService.LogSuccess(contextLabel + ": Generated file " + fullPath + " with row count " + rows.Count.ToString(CultureInfo.InvariantCulture));
        SetCsvStatus(resolvedFileName, rows.Count);
        WriteStatus(contextLabel + " completed");
        RefreshLogViewer();
    }

    private static string BuildCsvContent(IEnumerable<CsvMeasurementRow> rows, bool includeComment)
    {
        var builder = new StringBuilder();
        builder.Append("Serial,WorkDate,MeasuredAt,MachineCode,OK/NG,Measure1(P1),Measure2(P2),Measure3(P3),Measure4(P4),Beta1(P5),Beta2(P6),Beta3(P7),Beta4(P8),Beta5(P9),Beta6(P10),BetaRange(P5~P10),BetaAverage(P5~P10)");
        if (includeComment)
        {
            builder.Append(",Comment");
        }

        builder.AppendLine();

        foreach (var row in rows)
        {
            builder.Append(EscapeCsv(row.Serial));
            builder.Append(',');
            builder.Append(EscapeCsv(row.WorkDate));
            builder.Append(',');
            builder.Append(EscapeCsv(row.MeasuredAt));
            builder.Append(',');
            builder.Append(EscapeCsv(row.MachineCode));
            builder.Append(',');
            builder.Append(EscapeCsv(row.Result));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.P1));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.P2));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.P3));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.P4));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.Beta1));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.Beta2));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.Beta3));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.Beta4));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.Beta5));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.Beta6));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.BetaRange));
            builder.Append(',');
            builder.Append(FormatCsvNumber(row.BetaAverage));
            if (includeComment)
            {
                builder.Append(',');
                builder.Append(EscapeCsv(row.Comment));
            }

            builder.AppendLine();
        }

        return builder.ToString();
    }

    private static string FormatCsvNumber(decimal value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static decimal RoundCsvValue(decimal value)
    {
        return decimal.Round(value, 3, MidpointRounding.AwayFromZero);
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
            _lblCsvStatus.ForeColor = SuccessColor;
        }

        RecordRunSummary("CSVtoDB", fileName + " | rows: " + rowCount.ToString(CultureInfo.InvariantCulture), "Success");
    }

    private DateTime GetCsvBaseTime()
    {
        return _csvAutomationExecutionContext?.RunStamp ?? _dtCsvBase?.Value ?? _settings.DatetimeBase;
    }

    private string ResolveCsvOutputFileName(string fileName)
    {
        if (_csvAutomationExecutionContext is null)
        {
            return fileName;
        }

        var fileExtension = Path.GetExtension(fileName);
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
        var suffix =
            _csvAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
            "_r" +
            _csvAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture);

        return fileNameWithoutExtension + "_" + suffix + fileExtension;
    }

    private static string BuildDailySerialSeed(DateTime baseTime, int sequence)
    {
        return BuildDailySerialSeedNumber(baseTime, sequence).ToString(CultureInfo.InvariantCulture);
    }

    private static long BuildDailySerialSeedNumber(DateTime baseTime, int sequence)
    {
        return long.Parse(
            baseTime.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + sequence.ToString("D4", CultureInfo.InvariantCulture),
            CultureInfo.InvariantCulture);
    }

    private string? EnsureCsvOutputFolder()
    {
        var path = _txtCsvOutputFolder?.Text.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show("CSV output folder path is required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        try
        {
            Directory.CreateDirectory(path);
            return path;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Unable to create or access the CSV output folder." + Environment.NewLine + Environment.NewLine + ex.Message,
                "CSVtoDB",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return null;
        }
    }

    private void OpenCsvOutputFolder()
    {
        var outputFolder = EnsureCsvOutputFolder();
        if (outputFolder is null)
        {
            return;
        }

        OpenPath(outputFolder);
    }

    private void BrowseForCsvOutputFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Select the output folder for generated CSV files.",
            UseDescriptionForTitle = true
        };

        if (_txtCsvOutputFolder is not null && Directory.Exists(_txtCsvOutputFolder.Text))
        {
            dialog.SelectedPath = _txtCsvOutputFolder.Text;
        }

        if (dialog.ShowDialog(this) == DialogResult.OK && _txtCsvOutputFolder is not null)
        {
            _txtCsvOutputFolder.Text = dialog.SelectedPath;
        }
    }

    private static bool ValidateCsvFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName) || !fileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return !fileName.Contains(Path.DirectorySeparatorChar) && !fileName.Contains(Path.AltDirectorySeparatorChar);
    }

    private static List<string> ParseCommaSeparatedValues(string input)
    {
        return input
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(value => value.Trim())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();
    }

    private static bool TryParseKeyStart(string keyStart, out string prefix, out long number)
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

        return !string.IsNullOrWhiteSpace(numericPart) &&
               long.TryParse(numericPart, out number);
    }

    private Control CreateScenarioPlaceholderPage(string title, string description, IReadOnlyList<string> nextItems)
    {
        var page = CreateScrollablePage(out var stack);

        var hero = CreateCard(title, description, out var heroContent);
        heroContent.Controls.Add(CreateBadgeLabel("SCAFFOLDED", AccentSoft, AccentColor));

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 10, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Open Settings", (_, _) => NavigateTo(SettingsPageKey), true));
        actions.Controls.Add(CreateActionButton("Back To Dashboard", (_, _) => NavigateTo(DashboardPageKey)));
        heroContent.Controls.Add(actions);
        stack.Controls.Add(hero);

        var details = CreateCard("Planned First Pass", "These are the first building blocks queued for this module so we can keep implementation incremental and testable.", out var detailContent);
        var list = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            Margin = new Padding(0, 6, 0, 0)
        };

        for (var i = 0; i < nextItems.Count; i++)
        {
            list.Controls.Add(new Label
            {
                AutoSize = true,
                MaximumSize = new Size(1040, 0),
                Text = "- " + nextItems[i],
                Margin = new Padding(0, 0, 0, 10)
            });
        }

        detailContent.Controls.Add(list);
        stack.Controls.Add(details);

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateSettingsPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateSettingsOverviewCard());
        stack.Controls.Add(CreateSettingsConnectionCard());
        stack.Controls.Add(CreateSettingsFlagsCard());
        stack.Controls.Add(CreateSettingsIdRangesCard());
        stack.Controls.Add(CreateSettingsFileSyncCard());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateSettingsOverviewCard()
    {
        var card = CreateCard("Settings Module", "All editable runtime configuration lives here. Changes save directly to appsettings.json and take effect in the current UI without restarting the app.", out var content);

        var badges = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        _lblSettingsState = CreateBadgeLabel(string.Empty, SuccessSoft, SuccessColor);
        badges.Controls.Add(_lblSettingsState);

        _lblProductionWarning = CreateBadgeLabel(string.Empty, WarningSoft, WarningColor);
        badges.Controls.Add(_lblProductionWarning);

        content.Controls.Add(badges);

        content.Controls.Add(new Label
        {
            Text = "Config file: " + _configurationService.ConfigurationFilePath,
            AutoSize = true,
            ForeColor = MutedText,
            Margin = new Padding(0, 10, 0, 0)
        });

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 12, 0, 0)
        };

        actions.Controls.Add(CreateActionButton("Save Settings", (_, _) => SaveSettings(), true));
        actions.Controls.Add(CreateActionButton("Reload From File", (_, _) => ReloadSettingsFromDisk()));
        actions.Controls.Add(CreateActionButton("Test SQL Connection", (_, _) => TestSqlConnection()));
        actions.Controls.Add(CreateActionButton("Open Output Folder", (_, _) => OpenOrCreateOutputRoot()));

        content.Controls.Add(actions);
        return card;
    }

    private Control CreateSettingsConnectionCard()
    {
        var card = CreateCard("Connection, Paths, and Runtime Defaults", "These values drive the shared environment used by the generators and validation checks.", out var content);

        var grid = CreateSettingsGrid(3);

        _txtSqlConnectionString = AddTextBoxRow(
            grid,
            0,
            "SQL connection string",
            _settings.SqlConnectionString,
            "Primary SQL Server connection string for DataSyncer test data setup.",
            multiline: true);

        var outputActions = CreateInlineActionPanel();
        outputActions.Controls.Add(CreateMiniButton("Browse", (_, _) => BrowseForOutputFolder()));
        outputActions.Controls.Add(CreateMiniButton("Open", (_, _) => OpenOrCreateOutputRoot()));

        _txtOutputRootFolder = AddTextBoxRow(
            grid,
            1,
            "Output root folder",
            _settings.OutputRootFolder,
            "Root folder where generated CSV, JSON, TXT, and log output will be written.",
            actionControl: outputActions);

        _txtApiTestEndpoint = AddTextBoxRow(
            grid,
            2,
            "API test endpoint",
            _settings.ApiTestEndpoint,
            "Endpoint used later for DBtoJSON API export testing.",
            actionControl: CreateMiniButton("Open URL", (_, _) => OpenConfiguredUrl()));

        _txtRemoteFileSharePath = AddTextBoxRow(
            grid,
            3,
            "Remote file share path",
            _settings.RemoteFileSharePath,
            "UNC path or remote location used by FileSyncer scenarios.",
            actionControl: CreateMiniButton("Check", (_, _) => CheckRemotePath()));

        AddLabelCell(grid, "Base datetime", 4);
        _dtDatetimeBase = new DateTimePicker
        {
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd HH:mm:ss",
            ShowUpDown = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Value = _settings.DatetimeBase,
            AccessibleName = "Base datetime",
            Margin = new Padding(0, 0, 12, 12)
        };
        RegisterSettingsControl(_dtDatetimeBase, "Default timestamp seed used across generator scenarios.");
        grid.Controls.Add(_dtDatetimeBase, 1, 4);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 4);

        content.Controls.Add(grid);
        return card;
    }

    private Control CreateSettingsFlagsCard()
    {
        var card = CreateCard("Processing Flags and Shared Prefixes", "Keep these values aligned with the DataSyncer environment so generated test data uses the same markers as the live jobs under test.", out var content);

        var grid = CreateSettingsGrid(2);

        _txtSyncFlagNotProcessed = AddTextBoxRow(grid, 0, "SyncFlagNotProcessed", _settings.SyncFlagNotProcessed, "Value meaning a row has not yet been processed by DataSyncer.");
        _txtSyncFlagProcessed = AddTextBoxRow(grid, 1, "SyncFlagProcessed", _settings.SyncFlagProcessed, "Value meaning a row has already been processed.");
        _txtCsvIdPrefix = AddTextBoxRow(grid, 2, "CsvIdPrefix", _settings.CsvIdPrefix, "Shared prefix applied to CSVtoDB RecordId values.");

        content.Controls.Add(grid);
        return card;
    }

    private Control CreateSettingsIdRangesCard()
    {
        var card = CreateCard("ID Range Overrides", "These starting points keep scenario runs isolated and make it easier to avoid cross-test contamination between QA runs.", out var content);

        var grid = CreateSettingsGrid(2);

        _numCsvIdStart = AddNumericRow(grid, 0, "CSVtoDB start", _settings.CsvIdStart, 100000, 199999, "Default start for the CSVtoDB ID series.");
        _numDbToDbIdStart = AddNumericRow(grid, 1, "DBtoDB start", _settings.DbToDbIdStart, 200000, 299999, "Default start for the DBtoDB source IDs.");
        _numDbToJsonIdStart = AddNumericRow(grid, 2, "DBtoJSON start", _settings.DbToJsonIdStart, 300000, 399999, "Default start for DBtoJSON export IDs.");
        _numSqlQueryIdStart = AddNumericRow(grid, 3, "SQLQuery start", _settings.SqlQueryIdStart, 400000, 499999, "Default start for SQLQuery row IDs.");

        content.Controls.Add(grid);
        return card;
    }

    private Control CreateSettingsFileSyncCard()
    {
        var card = CreateCard("FileSyncer Prefixes", "These prefixes drive the generated file names for upload, download, and two-way conflict scenarios.", out var content);

        var grid = CreateSettingsGrid(2);

        _txtFileSyncUploadPrefix = AddTextBoxRow(grid, 0, "Upload prefix", _settings.FileSyncUploadPrefix, "Prefix used when creating client-to-server upload files.");
        _txtFileSyncDownloadPrefix = AddTextBoxRow(grid, 1, "Download prefix", _settings.FileSyncDownloadPrefix, "Prefix used when creating server-to-client download files.");
        _txtFileSyncTwoWayPrefix = AddTextBoxRow(grid, 2, "Two-way prefix", _settings.FileSyncTwoWayPrefix, "Prefix used for two-way conflict files.");

        content.Controls.Add(grid);
        return card;
    }

    private Control CreateLogViewerPage()
    {
        var page = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = ShellBackground
        };
        page.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        page.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var actionsCard = CreateCard("Log Viewer", "Review the current application log, clear the live file for a fresh run, or open the log file directly in an external editor.", out var actionContent);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Refresh Log", (_, _) => RefreshLogViewer(), true));
        actions.Controls.Add(CreateActionButton("Open Log File", (_, _) => OpenCurrentLogFile()));
        actions.Controls.Add(CreateActionButton("Save Copy", (_, _) => SaveLogCopy()));
        actions.Controls.Add(CreateActionButton("Clear Current Log", (_, _) => ClearCurrentLog()));

        actionContent.Controls.Add(actions);
        page.Controls.Add(actionsCard, 0, 0);

        var logCard = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = CardBackground,
            Padding = new Padding(18),
            Margin = new Padding(0)
        };
        logCard.Paint += DrawRoundedCard;

        _logViewer = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            BackColor = Color.FromArgb(15, 23, 42),
            ForeColor = Color.Gainsboro,
            Font = new Font("Cascadia Code", 10F),
            DetectUrls = false
        };
        logCard.Controls.Add(_logViewer);

        page.Controls.Add(logCard, 0, 1);
        return page;
    }

    private static TableLayoutPanel CreateSettingsGrid(int columns)
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        grid.ColumnCount = columns;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        if (columns == 3)
        {
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        }

        return grid;
    }

    private TextBox AddTextBoxRow(
        TableLayoutPanel grid,
        int row,
        string label,
        string value,
        string tooltip,
        bool multiline = false,
        Control? actionControl = null)
    {
        AddLabelCell(grid, label, row);

        var textBox = new TextBox
        {
            Text = value,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = label
        };

        if (multiline)
        {
            textBox.Multiline = true;
            textBox.Height = 70;
            textBox.ScrollBars = ScrollBars.Vertical;
        }

        RegisterSettingsControl(textBox, tooltip);
        grid.Controls.Add(textBox, 1, row);

        if (grid.ColumnCount > 2)
        {
            grid.Controls.Add(actionControl ?? new Panel { Width = 1, Height = 1 }, 2, row);
        }

        return textBox;
    }

    private NumericUpDown AddNumericRow(TableLayoutPanel grid, int row, string label, int value, int minimum, int maximum, string tooltip)
    {
        AddLabelCell(grid, label, row);

        var numeric = new NumericUpDown
        {
            Minimum = minimum,
            Maximum = maximum,
            Value = Math.Max(minimum, Math.Min(maximum, value)),
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Width = 220,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = label
        };

        RegisterSettingsControl(numeric, tooltip);
        grid.Controls.Add(numeric, 1, row);
        return numeric;
    }

    private static void AddLabelCell(TableLayoutPanel grid, string text, int row)
    {
        grid.Controls.Add(new Label
        {
            Text = text,
            AutoSize = true,
            MaximumSize = new Size(240, 0),
            ForeColor = Color.FromArgb(30, 41, 59),
            Margin = new Padding(0, 8, 14, 12)
        }, 0, row);
    }

    private static Panel CreateScrollablePage(out FlowLayoutPanel stack)
    {
        var page = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = ShellBackground
        };

        var canvas = new Panel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.Transparent,
            Location = new Point(0, 0),
            Margin = new Padding(0),
            Padding = new Padding(0)
        };

        stack = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0),
            Padding = new Padding(0),
            Location = new Point(0, 0),
            BackColor = Color.Transparent
        };

        canvas.Controls.Add(stack);
        page.Controls.Add(canvas);
        var stackRef = stack;
        var canvasRef = canvas;
        page.Resize += (_, _) => ResizeStackCards(page, canvasRef, stackRef);
        stack.ControlAdded += (_, _) => ResizeStackCards(page, canvasRef, stackRef);
        stack.ControlRemoved += (_, _) => ResizeStackCards(page, canvasRef, stackRef);
        return page;
    }

    private static void ResizeStackCards(Panel page, FlowLayoutPanel stack)
    {
        if (page.Controls.Count == 0 || page.Controls[0] is not Panel canvas)
        {
            return;
        }

        ResizeStackCards(page, canvas, stack);
    }

    private static void ResizeStackCards(Panel page, Panel canvas, FlowLayoutPanel stack)
    {
        var targetWidth = Math.Max(340, page.ClientSize.Width - 8 - SystemInformation.VerticalScrollBarWidth);

        canvas.SuspendLayout();
        stack.SuspendLayout();

        canvas.Width = targetWidth;
        stack.Width = targetWidth;

        foreach (Control control in stack.Controls)
        {
            control.Width = targetWidth;
            control.MaximumSize = new Size(targetWidth, 0);
            UpdateResponsiveCardMetrics(control, targetWidth);
        }

        var preferredHeight = stack.PreferredSize.Height;
        canvas.Height = preferredHeight;
        canvas.Left = Math.Max(0, (page.ClientSize.Width - targetWidth) / 2);
        canvas.Top = 0;
        page.AutoScrollMinSize = new Size(0, preferredHeight + 8);

        stack.ResumeLayout(true);
        canvas.ResumeLayout(true);
    }

    private static Panel CreateCard(string title, string description, out TableLayoutPanel content)
    {
        var card = new Panel
        {
            BackColor = CardBackground,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(22),
            Margin = new Padding(0, 0, 0, 16)
        };
        card.Paint += DrawRoundedCard;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };

        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = Color.FromArgb(15, 23, 42),
            Font = new Font("Segoe UI Semibold", 15F, FontStyle.Bold)
        });

        layout.Controls.Add(new Label
        {
            Text = description,
            AutoSize = true,
            MaximumSize = new Size(1040, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        });

        content = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            Margin = new Padding(0, 6, 0, 0),
            BackColor = Color.Transparent
        };

        layout.Controls.Add(content);
        card.Controls.Add(layout);
        return card;
    }

    private static Panel CreateMetricCard(string title, out Label valueLabel)
    {
        var panel = new Panel
        {
            BackColor = Color.FromArgb(249, 251, 252),
            Dock = DockStyle.Top,
            Padding = new Padding(18),
            Margin = new Padding(0, 0, 14, 14),
            MinimumSize = new Size(220, 150)
        };
        panel.Paint += DrawRoundedCard;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };

        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = MutedText,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
        });

        valueLabel = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(420, 0),
            ForeColor = Color.FromArgb(15, 23, 42),
            Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold),
            Margin = new Padding(0, 10, 0, 0)
        };

        layout.Controls.Add(valueLabel);
        panel.Controls.Add(layout);
        return panel;
    }

    private static void ConfigureResponsiveCardGrid(Control host, TableLayoutPanel grid, int singleColumnThreshold, params Control[] cards)
    {
        void ApplyLayout()
        {
            var useSingleColumn = host.ClientSize.Width < singleColumnThreshold;
            var useFourColumns = cards.Length == 4 && host.ClientSize.Width >= 1380;

            grid.SuspendLayout();
            grid.Controls.Clear();
            grid.RowStyles.Clear();
            grid.ColumnStyles.Clear();

            if (useSingleColumn)
            {
                grid.ColumnCount = 1;
                grid.RowCount = cards.Length;
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                for (var i = 0; i < cards.Length; i++)
                {
                    grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    cards[i].Margin = new Padding(0, 0, 0, i == cards.Length - 1 ? 0 : 14);
                    grid.Controls.Add(cards[i], 0, i);
                }
            }
            else if (useFourColumns)
            {
                grid.ColumnCount = 4;
                grid.RowCount = 1;

                for (var column = 0; column < 4; column++)
                {
                    grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
                }

                grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                for (var i = 0; i < cards.Length; i++)
                {
                    cards[i].Margin = new Padding(i == 0 ? 0 : 6, 0, i == cards.Length - 1 ? 0 : 6, 14);
                    grid.Controls.Add(cards[i], i, 0);
                }
            }
            else
            {
                grid.ColumnCount = 2;
                grid.RowCount = (int)Math.Ceiling(cards.Length / 2D);
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                for (var row = 0; row < grid.RowCount; row++)
                {
                    grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                }

                for (var i = 0; i < cards.Length; i++)
                {
                    var column = i % 2;
                    var row = i / 2;
                    cards[i].Margin = new Padding(column == 0 ? 0 : 6, 0, column == 0 ? 8 : 0, 14);
                    grid.Controls.Add(cards[i], column, row);
                }
            }

            grid.ResumeLayout(true);
        }

        host.Resize += (_, _) => ApplyLayout();
        ApplyLayout();
    }

    private static void UpdateResponsiveCardMetrics(Control control, int containerWidth)
    {
        var contentWidth = Math.Max(220, containerWidth - 56);

        if (control is Label label && label.MaximumSize.Width > 0)
        {
            label.MaximumSize = new Size(contentWidth, 0);
        }

        foreach (Control child in control.Controls)
        {
            UpdateResponsiveCardMetrics(child, contentWidth);
        }
    }

    private Button CreateHeroActionButton(string text, EventHandler onClick)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            ForeColor = NavBackground,
            Margin = new Padding(10, 0, 0, 0),
            Padding = new Padding(14, 8, 14, 8),
            Cursor = Cursors.Hand,
            UseVisualStyleBackColor = false
        };

        button.FlatAppearance.BorderSize = 0;
        button.Click += onClick;
        return button;
    }

    private Button CreateActionButton(string text, EventHandler onClick, bool primary = false)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            Margin = new Padding(0, 0, 10, 10),
            Padding = new Padding(14, 8, 14, 8),
            Cursor = Cursors.Hand,
            BackColor = primary ? AccentColor : CardBackground,
            ForeColor = primary ? Color.White : Color.FromArgb(30, 41, 59),
            UseVisualStyleBackColor = false
        };

        button.FlatAppearance.BorderColor = primary ? AccentColor : CardBorder;
        button.FlatAppearance.BorderSize = primary ? 0 : 1;
        button.Click += onClick;
        return button;
    }

    private static FlowLayoutPanel CreateInlineActionPanel()
    {
        return new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 12)
        };
    }

    private void ApplyResponsiveShellLayout()
    {
        ApplyHeroResponsiveLayout();
        ApplySidebarResponsiveLayout();
        ApplyContentHostResponsiveLayout();
        ResizeNavigationButtons();
    }

    private void ApplyHeroResponsiveLayout()
    {
        if (_rootLayout is null || _heroLayout is null || _heroContent is null || _heroActions is null)
        {
            return;
        }

        var compact = ClientSize.Width < 1360;

        _heroLayout.SuspendLayout();
        _heroActions.SuspendLayout();

        _heroLayout.ColumnStyles.Clear();
        _heroLayout.RowStyles.Clear();

        if (compact)
        {
            _rootLayout.RowStyles[0].Height = 176F;

            _heroLayout.ColumnCount = 1;
            _heroLayout.RowCount = 2;
            _heroLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _heroLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _heroLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _heroLayout.SetColumn(_heroContent, 0);
            _heroLayout.SetRow(_heroContent, 0);
            _heroLayout.SetColumn(_heroActions, 0);
            _heroLayout.SetRow(_heroActions, 1);

            _heroActions.FlowDirection = FlowDirection.LeftToRight;
            _heroActions.WrapContents = true;
            _heroActions.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            _heroActions.Margin = new Padding(0, 14, 0, 0);
            _heroActions.MaximumSize = new Size(Math.Max(320, ClientSize.Width - 120), 0);
        }
        else
        {
            _rootLayout.RowStyles[0].Height = 118F;

            _heroLayout.ColumnCount = 2;
            _heroLayout.RowCount = 1;
            _heroLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _heroLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _heroLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _heroLayout.SetColumn(_heroContent, 0);
            _heroLayout.SetRow(_heroContent, 0);
            _heroLayout.SetColumn(_heroActions, 1);
            _heroLayout.SetRow(_heroActions, 0);

            _heroActions.FlowDirection = FlowDirection.LeftToRight;
            _heroActions.WrapContents = false;
            _heroActions.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            _heroActions.Margin = new Padding(18, 0, 0, 0);
            _heroActions.MaximumSize = Size.Empty;
        }

        _heroActions.ResumeLayout(true);
        _heroLayout.ResumeLayout(true);
    }

    private void ApplySidebarResponsiveLayout()
    {
        if (_mainSplitContainer is null)
        {
            return;
        }

        var availableWidth = _mainSplitContainer.Width;
        if (availableWidth <= 0)
        {
            return;
        }

        var desiredWidth = availableWidth >= 1800 ? 420 :
            availableWidth >= 1500 ? 400 :
            availableWidth >= 1280 ? 380 :
            availableWidth >= 1120 ? 360 :
            330;

        var maxAllowed = Math.Max(_mainSplitContainer.Panel1MinSize, availableWidth - 420);
        var finalWidth = Math.Clamp(desiredWidth, _mainSplitContainer.Panel1MinSize, maxAllowed);

        if (finalWidth > 0 && _mainSplitContainer.SplitterDistance != finalWidth)
        {
            _mainSplitContainer.SplitterDistance = finalWidth;
        }
    }

    private void ApplyContentHostResponsiveLayout()
    {
        if (_contentHost is null || _mainSplitContainer is null)
        {
            return;
        }

        var availableWidth = _mainSplitContainer.Panel2.ClientSize.Width;
        if (availableWidth <= 0)
        {
            return;
        }

        var horizontalPadding = availableWidth >= 1400 ? 24 :
            availableWidth >= 1100 ? 18 :
            12;

        var verticalPadding = availableWidth >= 1100 ? 20 : 12;
        _contentHost.Padding = new Padding(horizontalPadding, verticalPadding, horizontalPadding, verticalPadding);
    }

    private void ResizeNavigationButtons()
    {
        if (_navigationStack is null)
        {
            return;
        }

        var availableWidth = Math.Max(300, _navigationStack.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 4);
        var buttonHeight = availableWidth >= 340 ? 88 : 96;

        foreach (var item in _navButtons.Values)
        {
            item.Container.Width = availableWidth;
            item.Container.Height = buttonHeight;
            item.TitleLabel.MaximumSize = new Size(Math.Max(220, availableWidth - 32), 0);
            item.SubtitleLabel.MaximumSize = new Size(Math.Max(220, availableWidth - 32), 0);
        }
    }

    private Button CreateMiniButton(string text, EventHandler onClick)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            Padding = new Padding(10, 4, 10, 4),
            Margin = new Padding(0, 0, 8, 0),
            Cursor = Cursors.Hand,
            BackColor = Color.FromArgb(248, 250, 252),
            ForeColor = Color.FromArgb(30, 41, 59),
            UseVisualStyleBackColor = false
        };

        button.FlatAppearance.BorderColor = CardBorder;
        button.FlatAppearance.BorderSize = 1;
        button.Click += onClick;
        return button;
    }

    private static Label CreateBadgeLabel(string text, Color backColor, Color foreColor)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            BackColor = backColor,
            ForeColor = foreColor,
            Padding = new Padding(10, 5, 10, 5),
            Margin = new Padding(0, 0, 10, 10),
            Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold)
        };
    }

    private void NavigateTo(string pageKey)
    {
        foreach (var page in _pages.Values)
        {
            page.Visible = false;
        }

        foreach (var entry in _navButtons)
        {
            var selected = string.Equals(entry.Key, pageKey, StringComparison.Ordinal);
            entry.Value.Container.BackColor = selected ? NavButtonSelected : NavButtonBackground;
            entry.Value.TitleLabel.ForeColor = Color.White;
            entry.Value.SubtitleLabel.ForeColor = selected
                ? Color.FromArgb(224, 242, 254)
                : Color.FromArgb(191, 219, 254);
        }

        if (_pages.TryGetValue(pageKey, out var pageControl))
        {
            pageControl.Visible = true;
            pageControl.BringToFront();
        }

        if (_currentPageLabel is not null)
        {
            _currentPageLabel.Text = GetPageTitle(pageKey);
        }

        WriteStatus("Viewing " + GetPageTitle(pageKey));
    }

    private static string GetPageTitle(string pageKey)
    {
        return pageKey switch
        {
            DashboardPageKey => "Home / Dashboard",
            CsvPageKey => "CSVtoDB",
            DbToDbPageKey => "DBtoDB",
            DbToJsonPageKey => "DBtoJSON",
            SqlQueryPageKey => "SQLQuery",
            ProgramExecutionPageKey => "ProgramExecution",
            FileSyncerPageKey => "FileSyncer",
            SettingsPageKey => "Settings",
            LogViewerPageKey => "Log Viewer",
            _ => "Workspace"
        };
    }

    private void LoadSettingsIntoInputs()
    {
        _suppressSettingsChangedEvents = true;

        if (_txtSqlConnectionString is not null)
        {
            _txtSqlConnectionString.Text = _settings.SqlConnectionString;
        }

        if (_txtOutputRootFolder is not null)
        {
            _txtOutputRootFolder.Text = _settings.OutputRootFolder;
        }

        if (_txtSyncFlagNotProcessed is not null)
        {
            _txtSyncFlagNotProcessed.Text = _settings.SyncFlagNotProcessed;
        }

        if (_txtSyncFlagProcessed is not null)
        {
            _txtSyncFlagProcessed.Text = _settings.SyncFlagProcessed;
        }

        if (_txtCsvIdPrefix is not null)
        {
            _txtCsvIdPrefix.Text = _settings.CsvIdPrefix;
        }

        if (_txtApiTestEndpoint is not null)
        {
            _txtApiTestEndpoint.Text = _settings.ApiTestEndpoint;
        }

        if (_txtRemoteFileSharePath is not null)
        {
            _txtRemoteFileSharePath.Text = _settings.RemoteFileSharePath;
        }

        if (_dtDatetimeBase is not null)
        {
            _dtDatetimeBase.Value = _settings.DatetimeBase;
        }

        if (_numCsvIdStart is not null)
        {
            _numCsvIdStart.Value = _settings.CsvIdStart;
        }

        if (_numDbToDbIdStart is not null)
        {
            _numDbToDbIdStart.Value = _settings.DbToDbIdStart;
        }

        if (_numDbToJsonIdStart is not null)
        {
            _numDbToJsonIdStart.Value = _settings.DbToJsonIdStart;
        }

        if (_numSqlQueryIdStart is not null)
        {
            _numSqlQueryIdStart.Value = _settings.SqlQueryIdStart;
        }

        if (_txtFileSyncUploadPrefix is not null)
        {
            _txtFileSyncUploadPrefix.Text = _settings.FileSyncUploadPrefix;
        }

        if (_txtFileSyncDownloadPrefix is not null)
        {
            _txtFileSyncDownloadPrefix.Text = _settings.FileSyncDownloadPrefix;
        }

        if (_txtFileSyncTwoWayPrefix is not null)
        {
            _txtFileSyncTwoWayPrefix.Text = _settings.FileSyncTwoWayPrefix;
        }

        if (_txtCsvOutputFolder is not null)
        {
            _txtCsvOutputFolder.Text = _settings.OutputRootFolder;
        }

        if (_dtCsvBase is not null)
        {
            _dtCsvBase.Value = _settings.DatetimeBase;
        }

        RefreshAdvancedModuleDefaultsFromSettings();

        _settingsDirty = false;
        _suppressSettingsChangedEvents = false;
    }

    private void RegisterSettingsControl(Control control, string tooltip)
    {
        _toolTip.SetToolTip(control, tooltip);

        switch (control)
        {
            case TextBox textBox:
                textBox.TextChanged += SettingsInputChanged;
                break;
            case NumericUpDown numeric:
                numeric.ValueChanged += SettingsInputChanged;
                break;
            case DateTimePicker picker:
                picker.ValueChanged += SettingsInputChanged;
                break;
        }
    }

    private void SettingsInputChanged(object? sender, EventArgs e)
    {
        if (_suppressSettingsChangedEvents)
        {
            return;
        }

        _settingsDirty = true;
        UpdateSettingsWarningState();
    }

    private void SaveSettings()
    {
        if (!TryBuildSettingsFromInputs(out var updatedSettings, out var validationMessage))
        {
            MessageBox.Show(validationMessage, "Settings Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            WriteStatus("Settings validation failed");
            return;
        }

        try
        {
            Directory.CreateDirectory(updatedSettings.OutputRootFolder);
            _configurationService.Save(updatedSettings);
            _settings.CopyFrom(updatedSettings);
            _logService.SetLogFolder(updatedSettings.OutputRootFolder);
            _settingsDirty = false;

            if (_txtCsvOutputFolder is not null)
            {
                _txtCsvOutputFolder.Text = updatedSettings.OutputRootFolder;
            }

            if (_dtCsvBase is not null)
            {
                _dtCsvBase.Value = updatedSettings.DatetimeBase;
            }

            RefreshAdvancedModuleDefaultsFromSettings();
            UpdateDashboardSummary();
            UpdateSettingsWarningState();

            _logService.LogSuccess("Settings saved to " + _configurationService.ConfigurationFilePath);
            RefreshLogViewer();
            WriteStatus("Settings saved");

            MessageBox.Show("Settings saved successfully.", "Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _logService.LogError("Saving settings failed: " + ex.Message);
            RefreshLogViewer();
            WriteStatus("Saving settings failed");

            MessageBox.Show(
                "Unable to save settings." + Environment.NewLine + Environment.NewLine + ex.Message,
                "Settings",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void ReloadSettingsFromDisk()
    {
        if (_settingsDirty)
        {
            var confirm = MessageBox.Show(
                "You have unsaved settings changes. Reloading will discard them. Continue?",
                "Reload Settings",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes)
            {
                return;
            }
        }

        var reloaded = _configurationService.Load();
        _settings.CopyFrom(reloaded);
        _logService.SetLogFolder(reloaded.OutputRootFolder);

        LoadSettingsIntoInputs();
        UpdateDashboardSummary();
        UpdateSettingsWarningState();

        _logService.LogInfo("Settings reloaded from disk");
        RefreshLogViewer();
        WriteStatus("Settings reloaded");
    }

    private bool TryBuildSettingsFromInputs(out AppSettings settings, out string validationMessage)
    {
        settings = new AppSettings();
        validationMessage = string.Empty;

        if (_txtOutputRootFolder is null ||
            _txtSyncFlagNotProcessed is null ||
            _txtSyncFlagProcessed is null ||
            _txtCsvIdPrefix is null ||
            _txtSqlConnectionString is null ||
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
            validationMessage = "Settings controls are not ready yet.";
            return false;
        }

        var outputRoot = _txtOutputRootFolder.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputRoot))
        {
            validationMessage = "Output root folder is required.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(_txtSyncFlagNotProcessed.Text))
        {
            validationMessage = "SyncFlagNotProcessed is required.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(_txtSyncFlagProcessed.Text))
        {
            validationMessage = "SyncFlagProcessed is required.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(_txtCsvIdPrefix.Text))
        {
            validationMessage = "CsvIdPrefix is required.";
            return false;
        }

        var endpoint = _txtApiTestEndpoint.Text.Trim();
        if (!string.IsNullOrWhiteSpace(endpoint) && !Uri.TryCreate(endpoint, UriKind.Absolute, out _))
        {
            validationMessage = "ApiTestEndpoint must be a valid absolute URL.";
            return false;
        }

        settings.SqlConnectionString = _txtSqlConnectionString.Text.Trim();
        settings.OutputRootFolder = outputRoot;
        settings.SyncFlagNotProcessed = _txtSyncFlagNotProcessed.Text.Trim();
        settings.SyncFlagProcessed = _txtSyncFlagProcessed.Text.Trim();
        settings.CsvIdPrefix = _txtCsvIdPrefix.Text.Trim();
        settings.ApiTestEndpoint = endpoint;
        settings.RemoteFileSharePath = _txtRemoteFileSharePath.Text.Trim();
        settings.DatetimeBase = _dtDatetimeBase.Value;
        settings.CsvIdStart = Decimal.ToInt32(_numCsvIdStart.Value);
        settings.DbToDbIdStart = Decimal.ToInt32(_numDbToDbIdStart.Value);
        settings.DbToJsonIdStart = Decimal.ToInt32(_numDbToJsonIdStart.Value);
        settings.SqlQueryIdStart = Decimal.ToInt32(_numSqlQueryIdStart.Value);
        settings.FileSyncUploadPrefix = _txtFileSyncUploadPrefix.Text.Trim();
        settings.FileSyncDownloadPrefix = _txtFileSyncDownloadPrefix.Text.Trim();
        settings.FileSyncTwoWayPrefix = _txtFileSyncTwoWayPrefix.Text.Trim();

        return true;
    }

    private void TestSqlConnection()
    {
        if (_txtSqlConnectionString is null)
        {
            return;
        }

        var connectionString = _txtSqlConnectionString.Text.Trim();
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            MessageBox.Show("Enter a SQL connection string first.", "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            using var connection = new SqlConnection(connectionString);
            connection.Open();

            var message = "SQL connection succeeded. Database: " + connection.Database;
            _logService.LogSuccess(message);
            RefreshLogViewer();
            WriteStatus("SQL connection succeeded");

            MessageBox.Show(message, "Connection Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            var maskedConnection = MaskConnectionString(connectionString);
            var detail = ex.InnerException is null ? ex.Message : ex.Message + Environment.NewLine + ex.InnerException.Message;

            _logService.LogError("SQL connection failed: " + ex.Message);
            RefreshLogViewer();
            WriteStatus("SQL connection failed");

            MessageBox.Show(
                "Connection test failed." + Environment.NewLine + Environment.NewLine +
                "Connection String:" + Environment.NewLine + maskedConnection + Environment.NewLine + Environment.NewLine +
                detail,
                "Connection Test",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void BrowseForOutputFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Select the root output folder for generated test data.",
            UseDescriptionForTitle = true
        };

        if (_txtOutputRootFolder is not null && Directory.Exists(_txtOutputRootFolder.Text))
        {
            dialog.SelectedPath = _txtOutputRootFolder.Text;
        }

        if (dialog.ShowDialog(this) == DialogResult.OK && _txtOutputRootFolder is not null)
        {
            _txtOutputRootFolder.Text = dialog.SelectedPath;
        }
    }

    private void OpenOrCreateOutputRoot()
    {
        var path = _txtOutputRootFolder?.Text.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            path = _settings.OutputRootFolder;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show("Set an output root folder first.", "Output Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            Directory.CreateDirectory(path);
            OpenPath(path);
            WriteStatus("Opened output root");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Unable to create or open the output folder." + Environment.NewLine + Environment.NewLine + ex.Message,
                "Output Folder",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void OpenConfiguredUrl()
    {
        if (_txtApiTestEndpoint is null)
        {
            return;
        }

        var value = _txtApiTestEndpoint.Text.Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            MessageBox.Show("Set the API endpoint first.", "Open URL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri))
        {
            MessageBox.Show("The API endpoint must be a valid absolute URL.", "Open URL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        OpenPath(uri.AbsoluteUri);
    }

    private void CheckRemotePath()
    {
        if (_txtRemoteFileSharePath is null)
        {
            return;
        }

        var path = _txtRemoteFileSharePath.Text.Trim();
        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show("Set a remote path first.", "Remote Path Check", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            var exists = Directory.Exists(path);
            var message = exists
                ? "Remote path is reachable."
                : "Remote path is not reachable right now.";

            if (exists)
            {
                _logService.LogSuccess("Remote path reachable: " + path);
            }
            else
            {
                _logService.LogWarning("Remote path unreachable: " + path);
            }

            RefreshLogViewer();
            WriteStatus(message);

            MessageBox.Show(message + Environment.NewLine + path, "Remote Path Check", MessageBoxButtons.OK, exists ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }
        catch (Exception ex)
        {
            _logService.LogError("Remote path check failed: " + ex.Message);
            RefreshLogViewer();
            WriteStatus("Remote path check failed");

            MessageBox.Show(
                "Unable to check the remote path." + Environment.NewLine + Environment.NewLine + ex.Message,
                "Remote Path Check",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void RefreshAllUiState()
    {
        UpdateDashboardSummary();
        UpdateSettingsWarningState();
        RefreshLogViewer();
        WriteStatus("Refreshed dashboard, settings state, and logs");
    }

    private void UpdateDashboardSummary()
    {
        if (_lblDashboardSql is not null)
        {
            if (string.IsNullOrWhiteSpace(_settings.SqlConnectionString))
            {
                _lblDashboardSql.Text = "Not configured yet";
                _lblDashboardSql.ForeColor = WarningColor;
            }
            else
            {
                _lblDashboardSql.Text = BuildSqlSummary(_settings.SqlConnectionString);
                _lblDashboardSql.ForeColor = _settings.HasPotentialProductionConnection() ? DangerColor : SuccessColor;
            }
        }

        if (_lblDashboardOutput is not null)
        {
            var (text, color) = BuildFolderSummary(_settings.OutputRootFolder, "Output root");
            _lblDashboardOutput.Text = text;
            _lblDashboardOutput.ForeColor = color;
        }

        if (_lblDashboardRemote is not null)
        {
            var remotePath = _settings.RemoteFileSharePath;
            if (string.IsNullOrWhiteSpace(remotePath))
            {
                _lblDashboardRemote.Text = "Not configured yet";
                _lblDashboardRemote.ForeColor = WarningColor;
            }
            else if (Directory.Exists(remotePath))
            {
                _lblDashboardRemote.Text = "Reachable" + Environment.NewLine + remotePath;
                _lblDashboardRemote.ForeColor = SuccessColor;
            }
            else
            {
                _lblDashboardRemote.Text = "Unavailable" + Environment.NewLine + remotePath;
                _lblDashboardRemote.ForeColor = WarningColor;
            }
        }

        if (_lblDashboardRuntime is not null)
        {
            _lblDashboardRuntime.Text =
                "Base time: " + _settings.DatetimeBase.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
                "CSV " + _settings.CsvIdStart.ToString(CultureInfo.InvariantCulture) + "  |  DBtoDB " + _settings.DbToDbIdStart.ToString(CultureInfo.InvariantCulture) + Environment.NewLine +
                "DBtoJSON " + _settings.DbToJsonIdStart.ToString(CultureInfo.InvariantCulture) + "  |  SQLQuery " + _settings.SqlQueryIdStart.ToString(CultureInfo.InvariantCulture);
            _lblDashboardRuntime.ForeColor = Color.FromArgb(30, 41, 59);
        }

        if (_lblDashboardSafety is not null)
        {
            _lblDashboardSafety.Text =
                "Config file: " + _configurationService.ConfigurationFilePath + Environment.NewLine +
                (_settings.HasPotentialProductionConnection()
                    ? "Safety check: review the SQL connection string carefully. It contains prod/prd and may point to a production-like environment."
                    : "Safety check: no production marker was detected in the saved SQL connection string.");

            _lblDashboardSafety.ForeColor = _settings.HasPotentialProductionConnection() ? DangerColor : MutedText;
        }

        if (_heroMetaLabel is not null)
        {
            _heroMetaLabel.Text =
                "Config file: " + _configurationService.ConfigurationFilePath + Environment.NewLine +
                "Current output root: " + _settings.OutputRootFolder;
        }
    }

    private void UpdateSettingsWarningState()
    {
        if (_lblSettingsState is not null)
        {
            _lblSettingsState.Text = _settingsDirty ? "Unsaved changes" : "Saved";
            _lblSettingsState.BackColor = _settingsDirty ? WarningSoft : SuccessSoft;
            _lblSettingsState.ForeColor = _settingsDirty ? WarningColor : SuccessColor;
        }

        if (_lblProductionWarning is not null)
        {
            var currentConnectionString = _txtSqlConnectionString?.Text ?? _settings.SqlConnectionString;
            var hasProductionMarker =
                currentConnectionString.Contains("prod", StringComparison.OrdinalIgnoreCase) ||
                currentConnectionString.Contains("prd", StringComparison.OrdinalIgnoreCase);

            _lblProductionWarning.Text = hasProductionMarker
                ? "Warning: connection string contains prod/prd"
                : "Safety check: no prod/prd marker detected";
            _lblProductionWarning.BackColor = hasProductionMarker ? DangerSoft : AccentSoft;
            _lblProductionWarning.ForeColor = hasProductionMarker ? DangerColor : AccentColor;
        }
    }

    private void RefreshLogViewer()
    {
        if (_logViewer is null)
        {
            return;
        }

        var lines = _logService.ReadCurrentLog()
            .Split(Environment.NewLine, StringSplitOptions.None);

        _logViewer.SuspendLayout();
        _logViewer.Clear();

        foreach (var line in lines)
        {
            if (line.Length == 0)
            {
                continue;
            }

            _logViewer.SelectionColor = GetLogLineColor(line);
            _logViewer.AppendText(line + Environment.NewLine);
        }

        _logViewer.SelectionColor = Color.Gainsboro;
        _logViewer.SelectionStart = _logViewer.TextLength;
        _logViewer.ScrollToCaret();
        _logViewer.ResumeLayout();
    }

    private static Color GetLogLineColor(string line)
    {
        if (line.Contains("[ERROR]", StringComparison.OrdinalIgnoreCase))
        {
            return Color.FromArgb(252, 165, 165);
        }

        if (line.Contains("[WARN]", StringComparison.OrdinalIgnoreCase))
        {
            return Color.FromArgb(253, 230, 138);
        }

        if (line.Contains("[SUCCESS]", StringComparison.OrdinalIgnoreCase))
        {
            return Color.FromArgb(134, 239, 172);
        }

        return Color.FromArgb(226, 232, 240);
    }

    private void ClearCurrentLog()
    {
        var confirm = MessageBox.Show(
            "Clear the current live log file?",
            "Clear Log",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
        {
            return;
        }

        _logService.ClearCurrentLog();
        RefreshLogViewer();
        WriteStatus("Current log cleared");
    }

    private void SaveLogCopy()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*",
            FileName = "DataSyncerTestGen_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".log",
            Title = "Save a copy of the current log"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        File.WriteAllText(dialog.FileName, _logService.ReadCurrentLog());
        WriteStatus("Saved log copy");
    }

    private void OpenCurrentLogFile()
    {
        var path = _logService.GetCurrentLogFilePath();
        if (!File.Exists(path))
        {
            File.WriteAllText(path, string.Empty);
        }

        OpenPath(path);
    }

    private void WriteStatus(string message)
    {
        if (_statusLabel is not null)
        {
            _statusLabel.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "  |  " + message;
        }
    }

    private static string BuildSqlSummary(string connectionString)
    {
        try
        {
            var builder = new SqlConnectionStringBuilder(connectionString);
            var server = string.IsNullOrWhiteSpace(builder.DataSource) ? "(server not set)" : builder.DataSource;
            var database = string.IsNullOrWhiteSpace(builder.InitialCatalog) ? "(database not set)" : builder.InitialCatalog;
            var auth = builder.IntegratedSecurity ? "Windows auth" : "SQL auth";

            return "Configured" + Environment.NewLine +
                   "Server: " + server + Environment.NewLine +
                   "Database: " + database + Environment.NewLine +
                   auth;
        }
        catch
        {
            return "Configured" + Environment.NewLine + MaskConnectionString(connectionString);
        }
    }

    private static (string Text, Color Color) BuildFolderSummary(string path, string label)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return ("Not configured yet", WarningColor);
        }

        if (!Directory.Exists(path))
        {
            return ("Missing folder" + Environment.NewLine + path + Environment.NewLine + label + " will be created when needed.", WarningColor);
        }

        try
        {
            var probeFile = Path.Combine(path, ".write-test-" + Guid.NewGuid().ToString("N") + ".tmp");
            File.WriteAllText(probeFile, "ok");
            File.Delete(probeFile);

            return ("Ready" + Environment.NewLine + path, SuccessColor);
        }
        catch
        {
            return ("Exists but not writable" + Environment.NewLine + path, DangerColor);
        }
    }

    private static string MaskConnectionString(string connectionString)
    {
        try
        {
            var builder = new SqlConnectionStringBuilder(connectionString);
            if (builder.ContainsKey("Password"))
            {
                builder.Password = "******";
            }

            return builder.ConnectionString;
        }
        catch
        {
            return connectionString;
        }
    }

    private static void OpenPath(string path)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true
        });
    }

    private static void DrawRoundedCard(object? sender, PaintEventArgs e)
    {
        if (sender is not Control control)
        {
            return;
        }

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(0, 0, control.Width - 1, control.Height - 1);
        using var path = CreateRoundedRectanglePath(rect, 18);
        using var pen = new Pen(CardBorder);
        e.Graphics.DrawPath(pen, path);
    }

    private static GraphicsPath CreateRoundedRectanglePath(Rectangle bounds, int radius)
    {
        var diameter = radius * 2;
        var path = new GraphicsPath();

        path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();

        return path;
    }

    private sealed record NavigationItem(Panel Container, Label TitleLabel, Label SubtitleLabel);

    private enum CsvScheduleMode
    {
        OneTime,
        Daily,
        EveryNMinutes
    }

    private sealed record CsvAutomationExecutionContext(DateTime RunStamp, int RunSequence);

    private sealed record CsvMeasurementProfile(
        decimal P1,
        decimal P2,
        decimal P3,
        decimal P4,
        decimal Beta1,
        decimal Beta2,
        decimal Beta3,
        decimal Beta4,
        decimal Beta5,
        decimal Beta6,
        decimal BetaRange,
        decimal BetaAverage);

    private sealed record CsvMeasurementRow(
        string Serial,
        string WorkDate,
        string MeasuredAt,
        string MachineCode,
        string Result,
        decimal P1,
        decimal P2,
        decimal P3,
        decimal P4,
        decimal Beta1,
        decimal Beta2,
        decimal Beta3,
        decimal Beta4,
        decimal Beta5,
        decimal Beta6,
        decimal BetaRange,
        decimal BetaAverage,
        string Comment);
}
