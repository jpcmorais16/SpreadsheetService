using System.Data.Common;
using SpreadsheetService.GoogleSheets;
using System.Data.SqlClient;
using MySqlConnector;
using Dapper;

namespace SpreadsheetService;

public class Worker : BackgroundService
{
    private readonly ILogger<Worker> _logger;
    private readonly string _connectionString;
    private readonly GoogleSheetsConnection _googleSheetsConnection;
    private readonly string _configurationSpreadsheetId;
    
    public Worker(ILogger<Worker> logger, IConfiguration configuration, GoogleSheetsConnection googleSheetsConnection)
    {
        _logger = logger;
        _googleSheetsConnection = googleSheetsConnection;
        _connectionString = configuration.GetConnectionString("DefaultConnection");
        _configurationSpreadsheetId = configuration.GetSection("SpreadsheetIds")["Config"];
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);

            var spreadsheetsToUpdate = GetSpreadsheets();

            foreach (var spreadsheet in spreadsheetsToUpdate)
            {
                await UpdateSpreadsheet(spreadsheet);
            }
            
            await UpdateDatabase();

            await Task.Delay(10000, stoppingToken);
        }
    }

    private List<Spreadsheet> GetSpreadsheets()
    {
        var spreadsheets = _googleSheetsConnection.GetRowsFromSpreadsheet(_configurationSpreadsheetId, "a");

        List<Spreadsheet> spreadsheetList = new List<Spreadsheet>();

        foreach (var (table, spreadsheetId, page) in spreadsheets)
        {
            if (string.IsNullOrEmpty(table) || string.IsNullOrEmpty(spreadsheetId) ||
                string.IsNullOrEmpty(page)) continue;
            
            
            var spreadsheet = new Spreadsheet()
            {
                Table = table,
                GoogleSheetsId = spreadsheetId,
                Page = page,
            };

            spreadsheetList.Add(spreadsheet);
        }

        return spreadsheetList;
    }

    private async Task UpdateSpreadsheet(Spreadsheet spreadsheet)
    {
        // read from the table of this spreadsheet
        using (var connection = new MySqlConnection(_connectionString))
        {
            var query = $"SELECT * FROM {spreadsheet.Table};";
            
            //var parameters = new List<MySqlParameter>() { new MySqlParameter("@checklistId", spreadsheet.Table) };

            await connection.OpenAsync();

            //var result = (await connection.QueryAsync(query, parameters)).AsList();

            var result = new List<List<object>>();

            await using (var command = new MySqlCommand(query, connection))
            {
                //command.Parameters.AddWithValue("@table", spreadsheet.Table);
                
                // Execute the query and obtain a SqlDataReader
                await using (var reader = await command.ExecuteReaderAsync())
                {
                    if (reader.HasRows)
                    {
                        while (await reader.ReadAsync())
                        {
                            var subResult = new List<object>();
                            
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                subResult.Add(reader.GetValue(i));
                            }
                            
                            result.Add(subResult);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No rows found.");
                    }
                }
            }

            await connection.CloseAsync();
            
            // upload the values to the google sheets spreadsheet
            
            if(result != null)
                await _googleSheetsConnection.UpdateSpreadsheet(spreadsheet.GoogleSheetsId, spreadsheet.Page, result);
            
            else
            {
                _logger.LogInformation("Table {table} empty at: {time}", spreadsheet.Table, DateTimeOffset.Now);
            }
        }
    }

    private async Task UpdateDatabase()
    {
        
    }
}