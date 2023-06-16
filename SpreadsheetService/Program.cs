using SpreadsheetService;
using SpreadsheetService.GoogleSheets;

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<Worker>();
        services.AddSingleton<GoogleSheetsConnection>();
    })
    .Build();

await host.RunAsync();