using System.Text.Json;
using DataSyncer.TestDataGenerator.App.Models;

namespace DataSyncer.TestDataGenerator.App.Services;

public sealed class ConfigurationService
{
    private const string FileName = "appsettings.json";
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _appRoot;
    private readonly string _settingsRoot;

    public ConfigurationService(string appRoot)
    {
        _appRoot = appRoot;
        _settingsRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DataSyncerTestGen");
    }

    public string ConfigurationFilePath => ResolveWritablePath();

    public AppSettings Load()
    {
        EnsureConfigurationFileExists();

        var filePath = ResolveWritablePath();

        var raw = File.ReadAllText(filePath);
        var parsed = JsonSerializer.Deserialize<AppSettings>(raw);
        var settings = parsed ?? new AppSettings();

        var changed = settings.NormalizeSqlConnectionString();

        if (settings.NormalizeOutputRootFolder())
        {
            changed = true;
        }

        if (changed)
        {
            Save(settings);
        }

        return settings;
    }

    public void Save(AppSettings settings)
    {
        var filePath = ResolveWritablePath();
        Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(filePath, json);
    }

    private void EnsureConfigurationFileExists()
    {
        var writablePath = ResolveWritablePath();
        if (File.Exists(writablePath))
        {
            return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(writablePath)!);

        var packagedPath = ResolvePackagedPath();
        if (File.Exists(packagedPath))
        {
            File.Copy(packagedPath, writablePath, overwrite: false);
            return;
        }

        var json = JsonSerializer.Serialize(new AppSettings(), JsonOptions);
        File.WriteAllText(writablePath, json);
    }

    private string ResolveWritablePath()
    {
        return Path.Combine(_settingsRoot, FileName);
    }

    private string ResolvePackagedPath()
    {
        return Path.Combine(_appRoot, FileName);
    }
}
