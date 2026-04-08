namespace DataSyncer.TestDataGenerator.App;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();

        var configurationService = new Services.ConfigurationService(AppContext.BaseDirectory);
        var settings = configurationService.Load();
        var logService = new Services.FileLogService(settings.OutputRootFolder);

        Application.Run(new Form1(settings, configurationService, logService));
    }    
}
