using System.Security.Cryptography;
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
            steps.Add(("SQLQuery", GenerateSqlQueryScenario1));
        }
        if (_chkGenerateAllProgramExecution?.Checked == true)
        {
            steps.Add(("ProgramExecution", () =>
            {
                CreateProgramOutputFolder(_txtProgramS1OutputFile, UpdateProgramStatus);
                ValidateProgramScriptPath(_txtProgramS1ScriptPath, UpdateProgramStatus);
                ConfirmProgramPathMissing();
                PreviewProgramCommand();
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
        stack.Controls.Add(CreateFileSyncerScenario1Card());
        stack.Controls.Add(CreateFileSyncerScenario2Card());
        stack.Controls.Add(CreateFileSyncerScenario3Card());

        ResizeStackCards(page, stack);
        return page;
    }

    private Control CreateFileSyncerOverviewCard()
    {
        var card = CreateCard("FileSyncer", "Create upload files locally, stage download files remotely, and prepare two-way conflict files with checksum details for pre-state capture.", out var content);

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
        return card;
    }

    private void GenerateFileSyncerScenario1()
    {
        var folder = EnsureFileSyncFolder(_txtFileSyncS1LocalSourceFolder, "FileSyncer upload source");
        if (folder is null)
        {
            return;
        }

        var txtPath = Path.Combine(folder, _settings.FileSyncUploadPrefix + "20260406_170000.txt");
        var csvPath = Path.Combine(folder, _settings.FileSyncUploadPrefix + "20260406_170001.csv");
        var jsonPath = Path.Combine(folder, _settings.FileSyncUploadPrefix + "20260406_170002.json");

        File.WriteAllText(txtPath, "Upload file generated at 2026-04-06 17:00:00");
        File.WriteAllText(csvPath, BuildFileSyncCsvContent());
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new[]
        {
            new { id = 1, value = "upload-a", ts = "2026-04-06 17:00:00" },
            new { id = 2, value = "upload-b", ts = "2026-04-06 17:00:01" },
            new { id = 3, value = "upload-c", ts = "2026-04-06 17:00:02" }
        }, new JsonSerializerOptions { WriteIndented = true }));

        SetFileSyncStatus("Scenario 1 generated 3 upload files in " + folder, SuccessColor);
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

        var txtPath = Path.Combine(remoteFolder, _settings.FileSyncDownloadPrefix + "20260406_171000.txt");
        var csvPath = Path.Combine(remoteFolder, _settings.FileSyncDownloadPrefix + "20260406_171001.csv");
        var jsonPath = Path.Combine(remoteFolder, _settings.FileSyncDownloadPrefix + "20260406_171002.json");

        File.WriteAllText(txtPath, "Download file generated at 2026-04-06 17:10:00");
        File.WriteAllText(csvPath, BuildFileSyncCsvContent());
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new[]
        {
            new { id = 1, value = "download-a", ts = "2026-04-06 17:10:00" },
            new { id = 2, value = "download-b", ts = "2026-04-06 17:10:01" },
            new { id = 3, value = "download-c", ts = "2026-04-06 17:10:02" }
        }, new JsonSerializerOptions { WriteIndented = true }));

        SetFileSyncStatus("Scenario 2 staged 3 remote download files in " + remoteFolder + ". Local target: " + localTarget, SuccessColor);
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

        foreach (var file in Directory.GetFiles(folder))
        {
            File.Delete(file);
        }
        foreach (var dir in Directory.GetDirectories(folder))
        {
            Directory.Delete(dir, recursive: true);
        }

        SetFileSyncStatus("Cleared local target folder: " + folder, SuccessColor);
    }

    private void GenerateFileSyncerScenario3()
    {
        var localFolder = EnsureFileSyncFolder(_txtFileSyncS3LocalFolder, "FileSyncer conflict local folder");
        var remoteFolder = EnsureRemoteFileSyncFolder();
        if (localFolder is null || remoteFolder is null)
        {
            return;
        }

        var localOnly = Path.Combine(localFolder, _settings.FileSyncTwoWayPrefix + "local_only.txt");
        var remoteOnly = Path.Combine(remoteFolder, _settings.FileSyncTwoWayPrefix + "remote_only.txt");
        var localConflict = Path.Combine(localFolder, _settings.FileSyncTwoWayPrefix + "shared_conflict.txt");
        var remoteConflict = Path.Combine(remoteFolder, _settings.FileSyncTwoWayPrefix + "shared_conflict.txt");

        File.WriteAllText(localOnly, "LOCAL ONLY file - 2026-04-06 17:20:00");
        File.WriteAllText(remoteOnly, "REMOTE ONLY file - 2026-04-06 17:20:00");
        File.WriteAllText(localConflict, "LOCAL VERSION 2026-04-06 17:20:00");
        File.WriteAllText(remoteConflict, "REMOTE VERSION 2026-04-06 17:21:00");

        var localHash = ComputeFileChecksum(localConflict);
        var remoteHash = ComputeFileChecksum(remoteConflict);
        var localSize = new FileInfo(localConflict).Length;
        var remoteSize = new FileInfo(remoteConflict).Length;

        if (_lblFileSyncConflictSummary is not null)
        {
            _lblFileSyncConflictSummary.Text =
                "Local conflict file: " + localSize + " bytes | SHA256 " + localHash + Environment.NewLine +
                "Remote conflict file: " + remoteSize + " bytes | SHA256 " + remoteHash;
            _lblFileSyncConflictSummary.ForeColor = SuccessColor;
        }

        SetFileSyncStatus("Scenario 3 generated all conflict files in the local and remote folders.", SuccessColor);
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

    private static string BuildFileSyncCsvContent()
    {
        return "RecordId,Value,Timestamp" + Environment.NewLine +
               "1,Alpha,2026-04-06 17:00:00" + Environment.NewLine +
               "2,Beta,2026-04-06 17:00:01" + Environment.NewLine +
               "3,Gamma,2026-04-06 17:00:02" + Environment.NewLine +
               "4,Delta,2026-04-06 17:00:03" + Environment.NewLine +
               "5,Epsilon,2026-04-06 17:00:04";
    }

    private static string ComputeFileChecksum(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return Convert.ToHexString(hash);
    }
}
