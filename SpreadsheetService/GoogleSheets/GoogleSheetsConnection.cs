using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace SpreadsheetService.GoogleSheets
{
    public class GoogleSheetsConnection
    {
        private List<SheetsService> _sheetsServiceList = new List<SheetsService>();
        private Tuple<int, int> _baseRange;

        public GoogleSheetsConnection(IConfiguration configuration)
        {
            GoogleSheetsServiceAccountCredentials credentials1 = configuration.GetSection("acc1").Get<GoogleSheetsServiceAccountCredentials>();


            var xCred1 = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(credentials1.client_email)
            {
                Scopes = new[] {
                SheetsService.Scope.Spreadsheets
            }
            }.FromPrivateKey(credentials1.private_key));
            

            var sheetsService1 = new SheetsService(
                new BaseClientService.Initializer()
                {
                    HttpClientInitializer = xCred1,
                }
            );
            
            _sheetsServiceList.Add(sheetsService1);
            
            _baseRange = new Tuple<int, int>(500, 15);
        }

        public Dictionary<string, List<string>> GetColumnsFromSpreadsheet(string spreadsheetId, string page, List<string> columns)
        {

            var response = MakeGetRequest(spreadsheetId, page);

            List<string> columnsFromSpreadsheet = response.Values[0].Select(o => o.ToString()).ToList();

            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();
            
            foreach(var c in columns)
            {
                result.Add(c, new List<string>());
            }

            
            foreach (List<object> row in response.Values.Skip(1))
            {
                List<string> data = new List<string>();

                for(int i = 0; i < row.Count; i++)
                {
                    if (result.Keys.Contains(columnsFromSpreadsheet[i]))
                        result[columnsFromSpreadsheet[i]].Add(row[i].ToString());

                }

            }

            return result;
        }

        public async Task<string> AppendRowToSpreadsheet(string spreadsheetId, string page, List<string> valuesToAppend)
        {
            List<object> list = new List<object>();

            foreach(string value in valuesToAppend)
            {
                list.Add(value);
            }

            var range = SpreadsheetRangeHelper.GetAppendRequestRange(1, valuesToAppend.Count, page);
            var valuerange = new ValueRange();
            valuerange.Values = new List<IList<object>>() { list };


            return await MakeAppendRequest(spreadsheetId, valuerange, range);
        }

        public async Task UpdateSpreadsheet(string spreadsheetId, string page, List<List<dynamic>>? valuesToAppend)
        {
            var list = new List<IList<object>>();

            if (valuesToAppend == null)
                throw new Exception();

            foreach(var valueList in valuesToAppend)
            {
                var sublist = new List<object>();
                
                foreach (var value in valueList)
                {
                    sublist.Add(value ?? "");
                }
                
                list.Add(sublist);
            }

            var range = SpreadsheetRangeHelper.GetUpdateRequestRange(valuesToAppend[0].Count , valuesToAppend.Count, page);
            
            var valuerange = new ValueRange();
            
            valuerange.Values = list;
            
            await MakeUpdateRequest(spreadsheetId, valuerange, range);
        }

        public void UpdateSingleCell(string spreadsheetId, string page, int column, int row, string value)
        {
            List<object> list = new List<object> { value };

            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>>() { list };

            var range = SpreadsheetRangeHelper.GetColumnUpdateRequestRange(row, column, page);
            MakeUpdateRequest(spreadsheetId, valueRange, range);
        }
        
        
        public List<(string?, string?, string?)> GetRowsFromSpreadsheet(string spreadsheetId, string page)
        {
            var response = MakeGetRequest(spreadsheetId, page);

            var columnsFromSpreadsheet = response.Values[0].Select(o => o.ToString()).ToList();

            var result = new List<(string?, string?, string?)>();
            
            foreach (var row in response.Values.Skip(1))
            {
                var table = row[0] as string;
                var spreadsheetIdToUpdate = row[1] as string;
                var pageToUpdate = row[2] as string;
                
                result.Add((table, spreadsheetIdToUpdate, pageToUpdate));
            }
            
            return result;
        }

        private ValueRange MakeGetRequest(string spreadsheetId, string page)
        {
            ValueRange? response = null;
            List<SheetsService>? localServiceList = _sheetsServiceList;

            for (int i = 0; i < _sheetsServiceList.Count; i++)
            {       
                try
                {
                    var rand = new Random().Next(0, _sheetsServiceList.Count - 1);
                    var service = localServiceList[rand];
                    

                    response = service.Spreadsheets.Values.Get(spreadsheetId,
                    SpreadsheetRangeHelper.GetReadRequestRange(1, _baseRange.Item1, 1, _baseRange.Item2, page)).Execute();

                    break;
                }
                catch
                {
                    continue;
                }
            }

            if (response == null)
                throw new Exception("Muitas requests!");

            return response; 
        }

        private async Task<string> MakeAppendRequest(string spreadsheetId, ValueRange valuerange, string range)
        {
            AppendValuesResponse? response = null;
            List<SheetsService>? localServiceList = _sheetsServiceList;
            

            var service = localServiceList[0];

            var appendRequest = service.Spreadsheets.Values.Append(valuerange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

            response = await appendRequest.ExecuteAsync();

            return response.Updates.UpdatedRange;
        }

        private async Task MakeUpdateRequest(string spreadsheetId, ValueRange valuerange, string range)
        {
            var service = _sheetsServiceList[0];
            
            var updateRequest = service.Spreadsheets.Values.Update(valuerange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            await updateRequest.ExecuteAsync();
        }
    }
}
