using System.Security.Cryptography;
using System.Globalization;
using System.Text.Json;

namespace DataSyncer.TestDataGenerator.App;

public partial class Form1
{
    private CheckBox? _chkGenerateAllCsv;
    private CheckBox? _chkGenerateAllDbToDb;
    private CheckBox? _chkGenerateAllDbToJson;
    private CheckBox? _chkGenerateAllSqlQuery;
    private CheckBox? _chkGenerateAllProgramExecution;
    private CheckBox? _chkGenerateAllFileSyncer;
    private ProgressBar? _progressGenerateAll;
    private Label? _lblGenerateAllStatus;
    private DataGridView? _gridRunSummary;

    private TextBox? _txtFileSyncS1LocalSourceFolder;
    private TextBox? _txtFileSyncRemoteTargetFolder;
    private Label? _lblFileSyncReachability;
    private Label? _lblFileSyncStatus;
    private TextBox? _txtFileSyncS2LocalTargetFolder;
    private TextBox? _txtFileSyncS3LocalFolder;
    private Label? _lblFileSyncConflictSummary;
    private ComboBox? _cmbFileSyncScheduleMode;
    private DateTimePicker? _dtFileSyncScheduleAt;
    private NumericUpDown? _numFileSyncScheduleIntervalMinutes;
    private ComboBox? _cmbFileSyncScheduleTarget;
    private Label? _lblFileSyncScheduleStatus;
    private readonly System.Windows.Forms.Timer _fileSyncScheduleTimer = new() { Interval = 1000 };
    private DateTime? _fileSyncScheduledRunAt;
    private FileSyncScheduleMode _armedFileSyncScheduleMode = FileSyncScheduleMode.OneTime;
    private TimeSpan? _fileSyncScheduleInterval;
    private FileSyncAutomationExecutionContext? _fileSyncAutomationExecutionContext;
    private int _fileSyncAutomationRunSequence;

    private Control CreateDashboardGenerateAllCard()
    {
        var card = CreateCard("Generate All Workspace", "Select the scenario groups to include in a one-click preparation run. Progress and recent results are tracked here for quick QA visibility.", out var content);

        var togglePanel = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        _chkGenerateAllCsv = CreateGenerateAllCheckBox("CSVtoDB");
        _chkGenerateAllDbToDb = CreateGenerateAllCheckBox("DBtoDB");
        _chkGenerateAllDbToJson = CreateGenerateAllCheckBox("DBtoJSON");
        _chkGenerateAllSqlQuery = CreateGenerateAllCheckBox("SQLQuery");
        _chkGenerateAllProgramExecution = CreateGenerateAllCheckBox("ProgramExecution");
        _chkGenerateAllFileSyncer = CreateGenerateAllCheckBox("FileSyncer");

        togglePanel.Controls.Add(_chkGenerateAllCsv);
        togglePanel.Controls.Add(_chkGenerateAllDbToDb);
        togglePanel.Controls.Add(_chkGenerateAllDbToJson);
        togglePanel.Controls.Add(_chkGenerateAllSqlQuery);
        togglePanel.Controls.Add(_chkGenerateAllProgramExecution);
        togglePanel.Controls.Add(_chkGenerateAllFileSyncer);
        content.Controls.Add(togglePanel);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 12, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate All Scenarios", (_, _) => RunGenerateAllScenarios(), true));
        content.Controls.Add(actions);

        _progressGenerateAll = new ProgressBar
        {
            Style = ProgressBarStyle.Continuous,
            Height = 22,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_progressGenerateAll);

        _lblGenerateAllStatus = new Label
        {
            Text = "Generate All is idle.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblGenerateAllStatus);

        _gridRunSummary = new DataGridView
        {
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Height = 220,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 12, 0, 0)
        };
        _gridRunSummary.Columns.Add("Scenario", "Scenario");
        _gridRunSummary.Columns.Add("Details", "Details");
        _gridRunSummary.Columns.Add("Timestamp", "Timestamp");
        _gridRunSummary.Columns.Add("Status", "Status");
        _gridRunSummary.Columns["Timestamp"]!.FillWeight = 26;
        _gridRunSummary.Columns["Status"]!.FillWeight = 20;
        content.Controls.Add(_gridRunSummary);

        return card;
    }

    private CheckBox CreateGenerateAllCheckBox(string label)
    {
        return new CheckBox
        {
            Text = label,
            Checked = true,
            AutoSize = true,
            Margin = new Padding(0, 0, 18, 10)
        };
    }

    private void RunGenerateAllScenarios()
    {
        if (_progressGenerateAll is null || _lblGenerateAllStatus is null)
        {
            return;
        }

        var steps = new List<(string Name, Action Action)>();
        if (_chkGenerateAllCsv?.Checked == true)
        {
            steps.Add(("CSVtoDB", GenerateAllCsvScenarios));
        }
        if (_chkGenerateAllDbToDb?.Checked == true)
        {
            steps.Add(("DBtoDB", () =>
            {
                GenerateDbToDbScenario1();
                GenerateDbToDbScenario2();
                GenerateDbToDbScenario3();
            }));
        }
        if (_chkGenerateAllDbToJson?.Checked == true)
        {
            steps.Add(("DBtoJSON", () =>
            {
                GenerateDbToJsonScenario1();
                GenerateDbToJsonScenario2();
                RefreshDbToJsonUnprocessedCount();
            }));
        }
        if (_chkGenerateAllSqlQuery?.Checked == true)
        {
            steps.Add(("SQLQuery", () =>
            {
                GenerateSqlQueryScenario1();
                PrepareSqlQueryScenario2();
                PrepareSqlQueryScenario3();
            }));
        }
        if (_chkGenerateAllProgramExecution?.Checked == true)
        {
            steps.Add(("ProgramExecution", () =>
            {
                PrepareProgramScenario1();
                ConfirmProgramPathMissing();
                PrepareProgramScenario3();
            }));
        }
        if (_chkGenerateAllFileSyncer?.Checked == true)
        {
            steps.Add(("FileSyncer", () =>
            {
                GenerateFileSyncerScenario1();
                StageFileSyncerScenario2();
                GenerateFileSyncerScenario3();
            }));
        }

        if (steps.Count == 0)
        {
            _lblGenerateAllStatus.Text = "Select at least one scenario group first.";
            _lblGenerateAllStatus.ForeColor = WarningColor;
            return;
        }

        _progressGenerateAll.Minimum = 0;
        _progressGenerateAll.Maximum = steps.Count;
        _progressGenerateAll.Value = 0;
        _lblGenerateAllStatus.Text = "Generate All is running...";
        _lblGenerateAllStatus.ForeColor = AccentColor;

        Cursor = Cursors.WaitCursor;
        try
        {
            for (var i = 0; i < steps.Count; i++)
            {
                _lblGenerateAllStatus.Text = "Running " + steps[i].Name + " (" + (i + 1) + "/" + steps.Count + ")";
                Application.DoEvents();

                try
                {
                    steps[i].Action();
                    RecordRunSummary(steps[i].Name, "Generate All step completed.", "Success");
                }
                catch (Exception ex)
                {
                    RecordRunSummary(steps[i].Name, "Generate All step failed: " + ex.Message, "Error");
                }

                _progressGenerateAll.Value = i + 1;
            }

            _lblGenerateAllStatus.Text = "Generate All completed.";
            _lblGenerateAllStatus.ForeColor = SuccessColor;
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private void RecordRunSummary(string scenario, string details, string status)
    {
        if (_gridRunSummary is null)
        {
            return;
        }

        _gridRunSummary.Rows.Insert(0, scenario, details, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture), status);
        if (_gridRunSummary.Rows.Count > 20)
        {
            _gridRunSummary.Rows.RemoveAt(_gridRunSummary.Rows.Count - 1);
        }
    }

    private Control CreateFileSyncerPage()
    {
        var page = CreateScrollablePage(out var stack);

        stack.Controls.Add(CreateFileSyncerOverviewCard());
        stack.Controls.Add(CreateFileSyncerAutomationCard());
        stack.Controls.Add(CreateFileSyncerScenario1Card());
        stack.Controls.Add(CreateFileSyncerScenario2Card());
        stack.Controls.Add(CreateFileSyncerScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateFileSyncerOverviewCard()
    {
        var card = CreateCard("FileSyncer", "Create upload files locally, stage download files remotely, and prepare two-way conflict files with checksum details for pre-state capture.", out var content);

        content.Controls.Add(CreateInfoNote(
            "Scope note",
            "This workspace prepares local folders and reachable remote path targets such as UNC shares or mounted locations. Native FTP/SFTP session handling, transfer retries, and protocol-specific logging are still validated in DataSyncer during the actual FileSyncer job run.",
            AccentSoft,
            AccentColor));

        _lblFileSyncStatus = new Label
        {
            Text = "No FileSyncer action has run yet.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblFileSyncStatus);
        return card;
    }

    private Control CreateFileSyncerAutomationCard()
    {
        var card = CreateCard("FileSyncer Automation Timer", "Keep this application open and let it refresh FileSyncer test files on a timer. Repeating every N minutes is useful when DataSyncer should keep finding new upload or download candidates.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Automation mode", 0);
        _cmbFileSyncScheduleMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "FileSyncer automation mode"
        };
        _cmbFileSyncScheduleMode.Items.AddRange(new object[]
        {
            "Run Once At Selected Time",
            "Repeat Daily At Selected Time",
            "Repeat Every N Minutes"
        });
        _cmbFileSyncScheduleMode.SelectedIndex = 0;
        _cmbFileSyncScheduleMode.SelectedIndexChanged += (_, _) => UpdateFileSyncAutomationControlState();
        grid.Controls.Add(_cmbFileSyncScheduleMode, 1, 0);

        AddLabelCell(grid, "Run at", 1);
        _dtFileSyncScheduleAt = CreateDateTimeInput(DateTime.Now.AddMinutes(5), "Date and time used for one-time and daily FileSyncer preparation.");
        grid.Controls.Add(_dtFileSyncScheduleAt, 1, 1);

        AddLabelCell(grid, "Every N minutes", 2);
        _numFileSyncScheduleIntervalMinutes = CreateCsvNumeric(1, 1440, 1, "Interval in minutes for recurring FileSyncer preparation.");
        _numFileSyncScheduleIntervalMinutes.AccessibleName = "FileSyncer automation interval minutes";
        grid.Controls.Add(_numFileSyncScheduleIntervalMinutes, 1, 2);

        AddLabelCell(grid, "Timer target", 3);
        _cmbFileSyncScheduleTarget = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(0, 0, 12, 12),
            AccessibleName = "FileSyncer automation target"
        };
        _cmbFileSyncScheduleTarget.Items.AddRange(new object[]
        {
            "Scenario 1 - Upload Files",
            "Scenario 2 - Download Files",
            "Scenario 3 - Two-Way Conflict Files",
            "All FileSyncer Scenarios"
        });
        _cmbFileSyncScheduleTarget.SelectedIndex = 0;
        grid.Controls.Add(_cmbFileSyncScheduleTarget, 1, 3);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Arm Timer", (_, _) => ArmFileSyncSchedule(), true));
        actions.Controls.Add(CreateActionButton("Cancel Timer", (_, _) => CancelFileSyncSchedule()));
        actions.Controls.Add(CreateActionButton("Run Target Now", (_, _) => RunFileSyncScheduledTarget()));
        content.Controls.Add(actions);

        _lblFileSyncScheduleStatus = new Label
        {
            Text = "FileSyncer scheduler idle. Choose a mode, then arm the timer.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblFileSyncScheduleStatus);

        _fileSyncScheduleTimer.Tick -= OnFileSyncScheduleTimerTick;
        _fileSyncScheduleTimer.Tick += OnFileSyncScheduleTimerTick;
        UpdateFileSyncAutomationControlState();
        return card;
    }

    private Control CreateFileSyncerScenario1Card()
    {
        var card = CreateCard("Scenario 1 - Client-to-Server Upload", "Generate the local upload files and verify remote share reachability before the FileSyncer job runs.", out var content);
        var grid = CreateCsvFormGrid(includeActionColumn: true);

        AddLabelCell(grid, "Local source folder", 0);
        _txtFileSyncS1LocalSourceFolder = CreateCsvTextBox(Path.Combine(_settings.OutputRootFolder, "FileSyncer", "Upload"), "Local source folder for upload files.");
        grid.Controls.Add(_txtFileSyncS1LocalSourceFolder, 1, 0);
        var localActions = CreateInlineActionPanel();
        localActions.Controls.Add(CreateMiniButton("Browse", (_, _) => BrowseForFolder(_txtFileSyncS1LocalSourceFolder, "Select upload source folder")));
        localActions.Controls.Add(CreateMiniButton("Open", (_, _) => OpenFolderFromTextBox(_txtFileSyncS1LocalSourceFolder)));
        grid.Controls.Add(localActions, 2, 0);

        AddLabelCell(grid, "Remote target folder", 1);
        _txtFileSyncRemoteTargetFolder = CreateCsvTextBox(_settings.RemoteFileSharePath, "Remote target share path from settings.");
        _txtFileSyncRemoteTargetFolder.ReadOnly = true;
        grid.Controls.Add(_txtFileSyncRemoteTargetFolder, 1, 1);
        grid.Controls.Add(CreateMiniButton("Retry Check", (_, _) => CheckFileSyncRemotePath()), 2, 1);

        _lblFileSyncReachability = CreateInlineStatusLabel("Remote path not checked yet.");
        grid.Controls.Add(_lblFileSyncReachability, 1, 2);
        grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, 2);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate Files", (_, _) => GenerateFileSyncerScenario1(), true));
        actions.Controls.Add(CreateActionButton("Check Remote Path", (_, _) => CheckFileSyncRemotePath()));
        content.Controls.Add(actions);
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use these local files with a Client to Server Only FileSyncer job. After DataSyncer runs, matching files should exist on the remote side and the transfer log should report upload success.",
            AccentSoft,
            AccentColor));
        return card;
    }

    private Control CreateFileSyncerScenario2Card()
    {
        var card = CreateCard("Scenario 2 - Server-to-Client Download", "Stage the download files on the remote share and keep the local target folder empty before the sync test.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Local target folder", 0);
        _txtFileSyncS2LocalTargetFolder = CreateCsvTextBox(Path.Combine(_settings.OutputRootFolder, "FileSyncer", "DownloadLocal"), "Local target folder that should start empty.");
        grid.Controls.Add(_txtFileSyncS2LocalTargetFolder, 1, 0);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Stage Remote Files", (_, _) => StageFileSyncerScenario2(), true));
        actions.Controls.Add(CreateActionButton("Verify Empty / Clear", (_, _) => VerifyOrClearDownloadFolder()));
        actions.Controls.Add(CreateActionButton("Open Local Target", (_, _) => OpenFolderFromTextBox(_txtFileSyncS2LocalTargetFolder)));
        content.Controls.Add(actions);
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Scenario 2 keeps the local target empty and stages fresh remote files first. After DataSyncer runs in Server to Client Only mode, the expected files should appear locally and the logs should confirm download success.",
            AccentSoft,
            AccentColor));
        return card;
    }

    private Control CreateFileSyncerScenario3Card()
    {
        var card = CreateCard("Scenario 3 - Two-Way Sync Conflict Check", "Generate the local-only, remote-only, and conflicting shared files, then show checksum and size details for the shared pair.", out var content);
        var grid = CreateCsvFormGrid();

        AddLabelCell(grid, "Local conflict folder", 0);
        _txtFileSyncS3LocalFolder = CreateCsvTextBox(Path.Combine(_settings.OutputRootFolder, "FileSyncer", "ConflictLocal"), "Local folder used for FileSyncer conflict setup.");
        grid.Controls.Add(_txtFileSyncS3LocalFolder, 1, 0);

        content.Controls.Add(grid);

        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("Generate All Conflict Files", (_, _) => GenerateFileSyncerScenario3(), true));
        actions.Controls.Add(CreateActionButton("Open Local Conflict Folder", (_, _) => OpenFolderFromTextBox(_txtFileSyncS3LocalFolder)));
        actions.Controls.Add(CreateActionButton("Check Remote Path", (_, _) => CheckFileSyncRemotePath()));
        content.Controls.Add(actions);

        _lblFileSyncConflictSummary = new Label
        {
            Text = "Conflict file checksum details will appear here after generation.",
            AutoSize = true,
            MaximumSize = new Size(1020, 0),
            ForeColor = MutedText,
            Margin = new Padding(0, 8, 0, 0)
        };
        content.Controls.Add(_lblFileSyncConflictSummary);
        content.Controls.Add(CreateInfoNote(
            "Expected after job",
            "Use this setup for a Both Way sync run. After DataSyncer finishes, both endpoints should remain consistent with the product's sync rules, and any overwrite or conflict ordering should be visible in the logs.",
            WarningSoft,
            WarningColor));
        return card;
    }

    private void GenerateFileSyncerScenario1()
    {
        var folder = EnsureFileSyncFolder(_txtFileSyncS1LocalSourceFolder, "FileSyncer upload source");
        if (folder is null)
        {
            return;
        }

        var baseTime = GetFileSyncRunStamp();
        var txtTime = baseTime;
        var csvTime = baseTime.AddSeconds(1);
        var jsonTime = baseTime.AddSeconds(2);
        var txtPath = Path.Combine(folder, _settings.FileSyncUploadPrefix + txtTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".txt");
        var csvPath = Path.Combine(folder, _settings.FileSyncUploadPrefix + csvTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".csv");
        var jsonPath = Path.Combine(folder, _settings.FileSyncUploadPrefix + jsonTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".json");

        File.WriteAllText(txtPath, "Upload file generated at " + txtTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + GetFileSyncAutomationStatusSuffix());
        File.WriteAllText(csvPath, BuildFileSyncCsvContent(csvTime, "upload"));
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new[]
        {
            new { id = 1, value = "upload-a", ts = txtTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
            new { id = 2, value = "upload-b", ts = csvTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
            new { id = 3, value = "upload-c", ts = jsonTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) }
        }, new JsonSerializerOptions { WriteIndented = true }));

        File.SetLastWriteTime(txtPath, txtTime);
        File.SetLastWriteTime(csvPath, csvTime);
        File.SetLastWriteTime(jsonPath, jsonTime);

        SetFileSyncStatus("Scenario 1 generated 3 upload files in " + folder + " for batch " + baseTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".", SuccessColor);
    }

    private void StageFileSyncerScenario2()
    {
        var remoteFolder = EnsureRemoteFileSyncFolder();
        if (remoteFolder is null)
        {
            return;
        }

        var localTarget = EnsureFileSyncFolder(_txtFileSyncS2LocalTargetFolder, "FileSyncer download local target");
        if (localTarget is null)
        {
            return;
        }

        var clearedEntries = ClearFileSyncFolderContents(localTarget);
        var baseTime = GetFileSyncRunStamp();
        var txtTime = baseTime;
        var csvTime = baseTime.AddSeconds(1);
        var jsonTime = baseTime.AddSeconds(2);
        var txtPath = Path.Combine(remoteFolder, _settings.FileSyncDownloadPrefix + txtTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".txt");
        var csvPath = Path.Combine(remoteFolder, _settings.FileSyncDownloadPrefix + csvTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".csv");
        var jsonPath = Path.Combine(remoteFolder, _settings.FileSyncDownloadPrefix + jsonTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".json");

        File.WriteAllText(txtPath, "Download file generated at " + txtTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + GetFileSyncAutomationStatusSuffix());
        File.WriteAllText(csvPath, BuildFileSyncCsvContent(csvTime, "download"));
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new[]
        {
            new { id = 1, value = "download-a", ts = txtTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
            new { id = 2, value = "download-b", ts = csvTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
            new { id = 3, value = "download-c", ts = jsonTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) }
        }, new JsonSerializerOptions { WriteIndented = true }));

        File.SetLastWriteTime(txtPath, txtTime);
        File.SetLastWriteTime(csvPath, csvTime);
        File.SetLastWriteTime(jsonPath, jsonTime);

        SetFileSyncStatus("Scenario 2 staged 3 remote download files in " + remoteFolder + ", and cleared " + clearedEntries + " local item(s) from " + localTarget + ".", SuccessColor);
    }

    private void VerifyOrClearDownloadFolder()
    {
        var folder = EnsureFileSyncFolder(_txtFileSyncS2LocalTargetFolder, "FileSyncer download local target");
        if (folder is null)
        {
            return;
        }

        var entries = Directory.GetFileSystemEntries(folder);
        if (entries.Length == 0)
        {
            SetFileSyncStatus("Local target folder is already empty: " + folder, SuccessColor);
            return;
        }

        var confirm = MessageBox.Show(
            "The local target folder contains " + entries.Length + " item(s)." + Environment.NewLine + Environment.NewLine + "Clear it now?",
            "FileSyncer",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
        {
            SetFileSyncStatus("Local target folder was not cleared.", WarningColor);
            return;
        }

        var clearedEntries = ClearFileSyncFolderContents(folder);
        SetFileSyncStatus("Cleared local target folder: " + folder + " (" + clearedEntries + " item(s)).", SuccessColor);
    }

    private void GenerateFileSyncerScenario3()
    {
        var localFolder = EnsureFileSyncFolder(_txtFileSyncS3LocalFolder, "FileSyncer conflict local folder");
        var remoteFolder = EnsureRemoteFileSyncFolder();
        if (localFolder is null || remoteFolder is null)
        {
            return;
        }

        var baseTime = GetFileSyncRunStamp();
        var batchStamp = baseTime.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
        var localOnly = Path.Combine(localFolder, _settings.FileSyncTwoWayPrefix + "local_only_" + batchStamp + ".txt");
        var remoteOnly = Path.Combine(remoteFolder, _settings.FileSyncTwoWayPrefix + "remote_only_" + batchStamp + ".txt");
        var localConflict = Path.Combine(localFolder, _settings.FileSyncTwoWayPrefix + "shared_conflict.txt");
        var remoteConflict = Path.Combine(remoteFolder, _settings.FileSyncTwoWayPrefix + "shared_conflict.txt");

        var localOnlyTime = baseTime;
        var remoteOnlyTime = baseTime.AddSeconds(5);
        var localConflictTime = baseTime.AddSeconds(10);
        var remoteConflictTime = baseTime.AddSeconds(20);

        File.WriteAllText(localOnly, "LOCAL ONLY file - " + localOnlyTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + GetFileSyncAutomationStatusSuffix());
        File.WriteAllText(remoteOnly, "REMOTE ONLY file - " + remoteOnlyTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + GetFileSyncAutomationStatusSuffix());
        File.WriteAllText(localConflict, "LOCAL VERSION " + localConflictTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + GetFileSyncAutomationStatusSuffix());
        File.WriteAllText(remoteConflict, "REMOTE VERSION " + remoteConflictTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + GetFileSyncAutomationStatusSuffix());
        File.SetLastWriteTime(localOnly, localOnlyTime);
        File.SetLastWriteTime(remoteOnly, remoteOnlyTime);
        File.SetLastWriteTime(localConflict, localConflictTime);
        File.SetLastWriteTime(remoteConflict, remoteConflictTime);

        var localHash = ComputeFileChecksum(localConflict);
        var remoteHash = ComputeFileChecksum(remoteConflict);
        var localSize = new FileInfo(localConflict).Length;
        var remoteSize = new FileInfo(remoteConflict).Length;

        if (_lblFileSyncConflictSummary is not null)
        {
            _lblFileSyncConflictSummary.Text =
                "Local conflict file: " + localSize + " bytes | SHA256 " + localHash + " | LastWrite " + localConflictTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
                "Remote conflict file: " + remoteSize + " bytes | SHA256 " + remoteHash + " | LastWrite " + remoteConflictTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            _lblFileSyncConflictSummary.ForeColor = SuccessColor;
        }

        SetFileSyncStatus("Scenario 3 generated local-only, remote-only, and shared-conflict files for batch " + batchStamp + ".", SuccessColor);
    }

    private int ClearFileSyncFolderContents(string folder)
    {
        var cleared = 0;
        foreach (var file in Directory.GetFiles(folder))
        {
            File.Delete(file);
            cleared++;
        }

        foreach (var dir in Directory.GetDirectories(folder))
        {
            Directory.Delete(dir, recursive: true);
            cleared++;
        }

        return cleared;
    }

    private DateTime GetFileSyncRunStamp()
    {
        return _fileSyncAutomationExecutionContext?.RunStamp ?? DateTime.Now;
    }

    private string GetFileSyncAutomationStatusSuffix()
    {
        return _fileSyncAutomationExecutionContext is null
            ? string.Empty
            : " Scheduler batch " +
              _fileSyncAutomationExecutionContext.RunStamp.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) +
              "_r" +
              _fileSyncAutomationExecutionContext.RunSequence.ToString("D3", CultureInfo.InvariantCulture) +
              ".";
    }

    private void OnFileSyncScheduleTimerTick(object? sender, EventArgs e)
    {
        HandleFileSyncScheduleTick();
    }

    private void ArmFileSyncSchedule()
    {
        if (_dtFileSyncScheduleAt is null ||
            _cmbFileSyncScheduleMode is null ||
            _numFileSyncScheduleIntervalMinutes is null)
        {
            return;
        }

        var mode = GetSelectedFileSyncScheduleMode();
        var scheduledAt = mode switch
        {
            FileSyncScheduleMode.Daily => ResolveNextFileSyncScheduledRun(_dtFileSyncScheduleAt.Value),
            FileSyncScheduleMode.EveryNMinutes => DateTime.Now.AddMinutes(decimal.ToDouble(_numFileSyncScheduleIntervalMinutes.Value)),
            _ => _dtFileSyncScheduleAt.Value
        };

        if (mode == FileSyncScheduleMode.OneTime && scheduledAt <= DateTime.Now)
        {
            MessageBox.Show(
                "Choose a time in the future for a one-time FileSyncer automation run.",
                "FileSyncer Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        _armedFileSyncScheduleMode = mode;
        _fileSyncScheduleInterval = mode == FileSyncScheduleMode.EveryNMinutes
            ? TimeSpan.FromMinutes(decimal.ToDouble(_numFileSyncScheduleIntervalMinutes.Value))
            : null;
        _fileSyncScheduledRunAt = scheduledAt;
        _fileSyncScheduleTimer.Start();
        UpdateFileSyncScheduleStatus();
        SetFileSyncStatus(
            "Armed FileSyncer automation for " + scheduledAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetFileSyncScheduleTargetLabel() +
            " | " + GetArmedFileSyncScheduleModeLabel(),
            SuccessColor);
    }

    private void CancelFileSyncSchedule()
    {
        _fileSyncScheduleTimer.Stop();
        _fileSyncScheduledRunAt = null;
        _fileSyncScheduleInterval = null;
        _armedFileSyncScheduleMode = FileSyncScheduleMode.OneTime;
        UpdateFileSyncScheduleStatus();
        SetFileSyncStatus("Cancelled FileSyncer automation timer.", WarningColor);
    }

    private void HandleFileSyncScheduleTick()
    {
        if (!_fileSyncScheduledRunAt.HasValue)
        {
            _fileSyncScheduleTimer.Stop();
            UpdateFileSyncScheduleStatus();
            return;
        }

        if (DateTime.Now < _fileSyncScheduledRunAt.Value)
        {
            UpdateFileSyncScheduleStatus();
            return;
        }

        RunFileSyncScheduledTarget();

        if (_armedFileSyncScheduleMode == FileSyncScheduleMode.Daily)
        {
            _fileSyncScheduledRunAt = ResolveNextFileSyncScheduledRun(_fileSyncScheduledRunAt.Value.AddDays(1));
            _fileSyncScheduleTimer.Start();
        }
        else if (_armedFileSyncScheduleMode == FileSyncScheduleMode.EveryNMinutes && _fileSyncScheduleInterval.HasValue)
        {
            _fileSyncScheduledRunAt = ResolveNextFileSyncRecurringRun(_fileSyncScheduledRunAt.Value, _fileSyncScheduleInterval.Value);
            _fileSyncScheduleTimer.Start();
        }
        else
        {
            _fileSyncScheduleTimer.Stop();
            _fileSyncScheduledRunAt = null;
            _fileSyncScheduleInterval = null;
        }

        UpdateFileSyncScheduleStatus();
    }

    private void RunFileSyncScheduledTarget()
    {
        try
        {
            _fileSyncAutomationExecutionContext = new FileSyncAutomationExecutionContext(DateTime.Now, ++_fileSyncAutomationRunSequence);
            switch (_cmbFileSyncScheduleTarget?.SelectedIndex)
            {
                case 1:
                    StageFileSyncerScenario2();
                    break;
                case 2:
                    GenerateFileSyncerScenario3();
                    break;
                case 3:
                    GenerateFileSyncerScenario1();
                    StageFileSyncerScenario2();
                    GenerateFileSyncerScenario3();
                    break;
                default:
                    GenerateFileSyncerScenario1();
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Scheduled FileSyncer preparation failed." + Environment.NewLine + Environment.NewLine + ex.Message,
                "FileSyncer Scheduler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            SetFileSyncStatus("Scheduled FileSyncer preparation failed: " + ex.Message, DangerColor);
        }
        finally
        {
            _fileSyncAutomationExecutionContext = null;
        }
    }

    private void UpdateFileSyncScheduleStatus()
    {
        if (_lblFileSyncScheduleStatus is null)
        {
            return;
        }

        if (!_fileSyncScheduledRunAt.HasValue)
        {
            _lblFileSyncScheduleStatus.Text = "FileSyncer scheduler idle. Choose a mode, then arm the timer.";
            _lblFileSyncScheduleStatus.ForeColor = MutedText;
            return;
        }

        var remaining = _fileSyncScheduledRunAt.Value - DateTime.Now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        _lblFileSyncScheduleStatus.Text =
            "Armed for " + _fileSyncScheduledRunAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
            " | " + GetFileSyncScheduleTargetLabel() +
            " | " + GetArmedFileSyncScheduleModeLabel() +
            " | remaining " + remaining.ToString(@"dd\.hh\:mm\:ss", CultureInfo.InvariantCulture);
        _lblFileSyncScheduleStatus.ForeColor = AccentColor;
    }

    private void UpdateFileSyncAutomationControlState()
    {
        var mode = GetSelectedFileSyncScheduleMode();
        if (_dtFileSyncScheduleAt is not null)
        {
            _dtFileSyncScheduleAt.Enabled = mode != FileSyncScheduleMode.EveryNMinutes;
        }

        if (_numFileSyncScheduleIntervalMinutes is not null)
        {
            _numFileSyncScheduleIntervalMinutes.Enabled = mode == FileSyncScheduleMode.EveryNMinutes;
        }
    }

    private FileSyncScheduleMode GetSelectedFileSyncScheduleMode()
    {
        return _cmbFileSyncScheduleMode?.SelectedIndex switch
        {
            1 => FileSyncScheduleMode.Daily,
            2 => FileSyncScheduleMode.EveryNMinutes,
            _ => FileSyncScheduleMode.OneTime
        };
    }

    private string GetFileSyncScheduleTargetLabel()
    {
        return _cmbFileSyncScheduleTarget?.SelectedIndex switch
        {
            1 => "Scenario 2 - Download Files",
            2 => "Scenario 3 - Two-Way Conflict Files",
            3 => "All FileSyncer Scenarios",
            _ => "Scenario 1 - Upload Files"
        };
    }

    private string GetArmedFileSyncScheduleModeLabel()
    {
        return _armedFileSyncScheduleMode switch
        {
            FileSyncScheduleMode.Daily => "repeats daily",
            FileSyncScheduleMode.EveryNMinutes when _fileSyncScheduleInterval.HasValue =>
                "every " + FormatFileSyncIntervalLabel(_fileSyncScheduleInterval.Value),
            _ => "one-time run"
        };
    }

    private static DateTime ResolveNextFileSyncScheduledRun(DateTime requestedTime)
    {
        var nextRun = requestedTime;
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.AddDays(1);
        }

        return nextRun;
    }

    private static DateTime ResolveNextFileSyncRecurringRun(DateTime previousScheduledAt, TimeSpan interval)
    {
        var nextRun = previousScheduledAt.Add(interval);
        while (nextRun <= DateTime.Now)
        {
            nextRun = nextRun.Add(interval);
        }

        return nextRun;
    }

    private static string FormatFileSyncIntervalLabel(TimeSpan interval)
    {
        var totalMinutes = Math.Max(1, Convert.ToInt32(interval.TotalMinutes));
        return totalMinutes.ToString(CultureInfo.InvariantCulture) + " minute" + (totalMinutes == 1 ? string.Empty : "s");
    }

    private void CheckFileSyncRemotePath()
    {
        var remoteFolder = _txtFileSyncRemoteTargetFolder?.Text.Trim();
        if (string.IsNullOrWhiteSpace(remoteFolder))
        {
            SetInlineStatus(_lblFileSyncReachability, "Remote path is not configured.", WarningColor);
            return;
        }

        var checkedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
        if (Directory.Exists(remoteFolder))
        {
            SetInlineStatus(_lblFileSyncReachability, "Reachable | checked " + checkedAt, SuccessColor);
            SetFileSyncStatus("Remote path is reachable: " + remoteFolder, SuccessColor);
        }
        else
        {
            SetInlineStatus(_lblFileSyncReachability, "Not reachable | checked " + checkedAt, DangerColor);
            SetFileSyncStatus("Remote path is not reachable: " + remoteFolder, DangerColor);
        }
    }

    private void RefreshFileSyncerDefaultsFromSettings()
    {
        if (_txtFileSyncS1LocalSourceFolder is not null)
        {
            _txtFileSyncS1LocalSourceFolder.Text = Path.Combine(_settings.OutputRootFolder, "FileSyncer", "Upload");
        }

        if (_txtFileSyncRemoteTargetFolder is not null)
        {
            _txtFileSyncRemoteTargetFolder.Text = _settings.RemoteFileSharePath;
        }

        if (_txtFileSyncS2LocalTargetFolder is not null)
        {
            _txtFileSyncS2LocalTargetFolder.Text = Path.Combine(_settings.OutputRootFolder, "FileSyncer", "DownloadLocal");
        }

        if (_txtFileSyncS3LocalFolder is not null)
        {
            _txtFileSyncS3LocalFolder.Text = Path.Combine(_settings.OutputRootFolder, "FileSyncer", "ConflictLocal");
        }
    }

    private string? EnsureFileSyncFolder(TextBox? textBox, string label)
    {
        var folder = textBox?.Text.Trim();
        if (string.IsNullOrWhiteSpace(folder))
        {
            MessageBox.Show(label + " is required.", "FileSyncer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        Directory.CreateDirectory(folder);
        return folder;
    }

    private string? EnsureRemoteFileSyncFolder()
    {
        var remoteFolder = _txtFileSyncRemoteTargetFolder?.Text.Trim();
        if (string.IsNullOrWhiteSpace(remoteFolder))
        {
            MessageBox.Show("Remote target folder is not configured.", "FileSyncer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        try
        {
            Directory.CreateDirectory(remoteFolder);
            return remoteFolder;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Unable to access the remote target folder." + Environment.NewLine + Environment.NewLine + ex.Message,
                "FileSyncer",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return null;
        }
    }

    private void BrowseForFolder(TextBox? textBox, string description)
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = description,
            UseDescriptionForTitle = true
        };

        if (textBox is not null && Directory.Exists(textBox.Text))
        {
            dialog.SelectedPath = textBox.Text;
        }

        if (dialog.ShowDialog(this) == DialogResult.OK && textBox is not null)
        {
            textBox.Text = dialog.SelectedPath;
        }
    }

    private void OpenFolderFromTextBox(TextBox? textBox)
    {
        var folder = EnsureFileSyncFolder(textBox, "Folder");
        if (folder is not null)
        {
            OpenPath(folder);
        }
    }

    private void SetFileSyncStatus(string message, Color color)
    {
        if (_lblFileSyncStatus is not null)
        {
            _lblFileSyncStatus.Text = message;
            _lblFileSyncStatus.ForeColor = color;
        }

        RecordRunSummary("FileSyncer", message, color == SuccessColor ? "Success" : color == DangerColor ? "Error" : "Info");
        _logService.LogInfo("FileSyncer: " + message);
        WriteStatus(message);
        RefreshLogViewer();
    }

    private static string BuildFileSyncCsvContent(DateTime baseTime, string labelPrefix)
    {
        return "RecordId,Value,Timestamp" + Environment.NewLine +
               "1," + labelPrefix + "-Alpha," + baseTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
               "2," + labelPrefix + "-Beta," + baseTime.AddSeconds(1).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
               "3," + labelPrefix + "-Gamma," + baseTime.AddSeconds(2).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
               "4," + labelPrefix + "-Delta," + baseTime.AddSeconds(3).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
               "5," + labelPrefix + "-Epsilon," + baseTime.AddSeconds(4).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
    }

    private static string ComputeFileChecksum(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return Convert.ToHexString(hash);
    }

    private enum FileSyncScheduleMode
    {
        OneTime,
        Daily,
        EveryNMinutes
    }

    private sealed record FileSyncAutomationExecutionContext(DateTime RunStamp, int RunSequence);
}
