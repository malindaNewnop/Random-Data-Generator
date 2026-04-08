namespace DataSyncer.TestDataGenerator.App.Models;

public sealed class AppSettings
{
    public string SqlConnectionString { get; set; } = "Server=localhost;Database=DSTest;Trusted_Connection=True;TrustServerCertificate=True;";

    public string OutputRootFolder { get; set; } = @"D:\DataSyncerTest\GeneratedData";

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
}