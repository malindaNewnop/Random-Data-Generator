namespace DataSyncer.TestDataGenerator.App.Models;

public sealed class AppSettings
{
    public const string LegacyDefaultOutputRootFolder = @"D:\DataSyncerTest\GeneratedData";

    public string SqlConnectionString { get; set; } = "Server=localhost;Database=DSTest;Trusted_Connection=True;TrustServerCertificate=True;";

    public string OutputRootFolder { get; set; } = GetRecommendedOutputRootFolder();

    public string SyncFlagNotProcessed { get; set; } = "N";

    public string SyncFlagProcessed { get; set; } = "Y";

    public string CsvIdPrefix { get; set; } = "CSV-";

    public string ApiTestEndpoint { get; set; } = "http://localhost:5000/api/test";

    public string RemoteFileSharePath { get; set; } = @"\\server\share\DSTest";

    public DateTime DatetimeBase { get; set; } = new DateTime(2026, 4, 6, 9, 0, 0);

    public int CsvIdStart { get; set; } = 100001;

    public int DbToDbIdStart { get; set; } = 200001;

    public int DbToJsonIdStart { get; set; } = 300001;

    public int SqlQueryIdStart { get; set; } = 400001;

    public string FileSyncUploadPrefix { get; set; } = "upload_";

    public string FileSyncDownloadPrefix { get; set; } = "download_";

    public string FileSyncTwoWayPrefix { get; set; } = "twoway_";

    public static string GetRecommendedOutputRootFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DataSyncerTestGen",
            "GeneratedData");
    }

    public bool NormalizeOutputRootFolder()
    {
        var normalizedPath = Environment.ExpandEnvironmentVariables(OutputRootFolder?.Trim() ?? string.Empty);
        var changed = !string.Equals(OutputRootFolder, normalizedPath, StringComparison.Ordinal);
        OutputRootFolder = normalizedPath;

        if (string.IsNullOrWhiteSpace(OutputRootFolder))
        {
            OutputRootFolder = GetRecommendedOutputRootFolder();
            return true;
        }

        if (!string.Equals(OutputRootFolder, LegacyDefaultOutputRootFolder, StringComparison.OrdinalIgnoreCase))
        {
            return changed;
        }

        if (IsRootAccessible(OutputRootFolder))
        {
            return changed;
        }

        OutputRootFolder = GetRecommendedOutputRootFolder();
        return true;
    }

    public void CopyFrom(AppSettings source)
    {
        ArgumentNullException.ThrowIfNull(source);

        SqlConnectionString = source.SqlConnectionString;
        OutputRootFolder = source.OutputRootFolder;
        SyncFlagNotProcessed = source.SyncFlagNotProcessed;
        SyncFlagProcessed = source.SyncFlagProcessed;
        CsvIdPrefix = source.CsvIdPrefix;
        ApiTestEndpoint = source.ApiTestEndpoint;
        RemoteFileSharePath = source.RemoteFileSharePath;
        DatetimeBase = source.DatetimeBase;
        CsvIdStart = source.CsvIdStart;
        DbToDbIdStart = source.DbToDbIdStart;
        DbToJsonIdStart = source.DbToJsonIdStart;
        SqlQueryIdStart = source.SqlQueryIdStart;
        FileSyncUploadPrefix = source.FileSyncUploadPrefix;
        FileSyncDownloadPrefix = source.FileSyncDownloadPrefix;
        FileSyncTwoWayPrefix = source.FileSyncTwoWayPrefix;
    }

    public bool HasPotentialProductionConnection()
    {
        return SqlConnectionString.Contains("prod", StringComparison.OrdinalIgnoreCase) ||
               SqlConnectionString.Contains("prd", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsRootAccessible(string path)
    {
        try
        {
            if (path.StartsWith(@"\\", StringComparison.Ordinal))
            {
                return true;
            }

            var root = Path.GetPathRoot(path);
            if (string.IsNullOrWhiteSpace(root))
            {
                return true;
            }

            return Directory.Exists(root);
        }
        catch
        {
            return false;
        }
    }
}
