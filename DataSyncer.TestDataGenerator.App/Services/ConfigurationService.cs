using System.Text.Json;
using DataSyncer.TestDataGenerator.App.Models;

namespace DataSyncer.TestDataGenerator.App.Services;

public sealed class ConfigurationService
{
    private const string FileName = "appsettings.json";
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _appRoot;

    public ConfigurationService(string appRoot)
    {
        _appRoot = appRoot;
    }

    public string ConfigurationFilePath => ResolvePath();

    public AppSettings Load()
    {
        var filePath = ResolvePath();
        if (!File.Exists(filePath))
        {
            var defaults = new AppSettings();
            Save(defaults);
            return defaults;
        }

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
        var filePath = ResolvePath();
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(filePath, json);
    }

    private string ResolvePath()
    {
        return Path.Combine(_appRoot, FileName);
    }
}
