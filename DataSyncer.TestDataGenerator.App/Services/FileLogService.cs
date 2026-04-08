namespace DataSyncer.TestDataGenerator.App.Services;

public sealed class FileLogService
{
    private readonly string _logFolder;

    public FileLogService(string logFolder)
    {
        _logFolder = logFolder;
        Directory.CreateDirectory(_logFolder);
    }

    public void LogInfo(string message)
    {
        Append("INFO", message);
    }

    public void LogError(string message)
    {
        Append("ERROR", message);
    }

    public string ReadCurrentLog()
    {
        var path = CurrentLogFilePath();
        if (!File.Exists(path))
        {
            return string.Empty;
        }

        return File.ReadAllText(path);
    }

    private void Append(string level, string message)
    {
        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}";
        File.AppendAllText(CurrentLogFilePath(), line);
    }

    private string CurrentLogFilePath()
    {
        return Path.Combine(_logFolder, $"app-{DateTime.Now:yyyyMMdd}.log");
    }
}