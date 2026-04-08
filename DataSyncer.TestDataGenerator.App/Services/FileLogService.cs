namespace DataSyncer.TestDataGenerator.App.Services;

public sealed class FileLogService
{
    private const long MaxFileBytes = 10 * 1024 * 1024;
    private readonly object _syncRoot = new();
    private string _logFolder;

    public FileLogService(string logFolder)
    {
        _logFolder = logFolder;
        EnsureLogFolder();
    }

    public void SetLogFolder(string logFolder)
    {
        lock (_syncRoot)
        {
            _logFolder = logFolder;
            EnsureLogFolder();
        }
    }

    public void LogInfo(string message)
    {
        Append("INFO", message);
    }

    public void LogError(string message)
    {
        Append("ERROR", message);
    }

    public void LogWarning(string message)
    {
        Append("WARN", message);
    }

    public void LogSuccess(string message)
    {
        Append("SUCCESS", message);
    }

    public string ReadCurrentLog()
    {
        var path = GetCurrentLogFilePath();
        if (!File.Exists(path))
        {
            return string.Empty;
        }

        return File.ReadAllText(path);
    }

    public string GetCurrentLogFilePath()
    {
        lock (_syncRoot)
        {
            EnsureLogFolder();
            return ResolveLatestLogFilePath(DateTime.Now);
        }
    }

    public void ClearCurrentLog()
    {
        lock (_syncRoot)
        {
            EnsureLogFolder();
            File.WriteAllText(ResolveLatestLogFilePath(DateTime.Now), string.Empty);
        }
    }

    private void Append(string level, string message)
    {
        lock (_syncRoot)
        {
            EnsureLogFolder();
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}";
            File.AppendAllText(ResolveWritableLogFilePath(DateTime.Now), line);
        }
    }

    private void EnsureLogFolder()
    {
        try
        {
            Directory.CreateDirectory(_logFolder);
        }
        catch (Exception ex) when (
            ex is DirectoryNotFoundException ||
            ex is IOException ||
            ex is UnauthorizedAccessException ||
            ex is ArgumentException ||
            ex is NotSupportedException)
        {
            _logFolder = ResolveFallbackLogFolder();
            Directory.CreateDirectory(_logFolder);
        }

        PruneExpiredLogs();
    }

    private static string ResolveFallbackLogFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DataSyncerTestGen",
            "Logs");
    }

    private void PruneExpiredLogs()
    {
        var cutoff = DateTime.Today.AddDays(-7);
        foreach (var file in Directory.GetFiles(_logFolder, "DataSyncerTestGen_*.log"))
        {
            if (File.GetLastWriteTime(file) < cutoff)
            {
                File.Delete(file);
            }
        }
    }

    private string ResolveWritableLogFilePath(DateTime now)
    {
        var baseName = $"DataSyncerTestGen_{now:yyyyMMdd}";
        var index = 0;

        while (true)
        {
            var suffix = index == 0 ? string.Empty : "_" + index;
            var path = Path.Combine(_logFolder, baseName + suffix + ".log");

            if (!File.Exists(path))
            {
                return path;
            }

            var info = new FileInfo(path);
            if (info.Length < MaxFileBytes)
            {
                return path;
            }

            index++;
        }
    }

    private string ResolveLatestLogFilePath(DateTime now)
    {
        var baseName = $"DataSyncerTestGen_{now:yyyyMMdd}";
        var matches = Directory.GetFiles(_logFolder, baseName + "*.log")
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (matches.Count == 0)
        {
            return Path.Combine(_logFolder, baseName + ".log");
        }

        return matches[^1];
    }
}
