using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

namespace csv2GoogleSheet
{
    internal class Program
    {
        static int resRows;
        static void Main(string[] args)
        {
            string csvFilePath = "01_r0024_0001.csv";
            if (args.Length > 0)
            {
                csvFilePath = args[0];
            }
            string apiKeyFilePath = ConfigurationManager.AppSettings["GoogleSheetApiKey"] ?? "impactful-bee-392504-8202288b3779.json";
            string spreadsheetId = ConfigurationManager.AppSettings["GoogleSheetId"] ?? "1pwjLpTe-Vh8ONGtKpCculTZpI6szm_0MVVpfHV9oonA";
            string sheetName = ConfigurationManager.AppSettings["GoogleSheetName"] ?? "Sheet1";

            GoogleCredential credential;
            using (var stream = new FileStream(apiKeyFilePath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "csv2GoogleSheet"
            });
            List<string[]> csvData = LoadCsvData(csvFilePath);            
            List<string[]> processedData = ProcessCsvData(csvData);   
            string range = $"{sheetName}!A1:Z";
            var data = ConvertToValueRange(processedData);
            try
            {
                AppendDataToSheet(service, data, spreadsheetId, range);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        static List<string[]> LoadCsvData(string csvFilePath)
        {
            List<string[]> csvData = new List<string[]>();
            using (var reader = new StreamReader(csvFilePath))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    string[] values = line.Split(',');
                    csvData.Add(values);
                }
            }
            return csvData;
        }
        /// <summary>
        /// Remove double quotes from a string
        /// </summary>
        /// <param name="csvData"></param>
        /// <returns></returns>
        static List<string[]> ProcessCsvData(List<string[]> csvData)
        {
            List<string[]> processedData = new List<string[]>();
            foreach (var row in csvData)
            {
                string[] processedRow = new string[row.Length];
                for (int i = 0; i < row.Length; i++)
                {
                    string field = row[i];
                    if (field.StartsWith("\"") && field.EndsWith("\"") && field.Length >= 2)
                    {
                        processedRow[i] = field.Substring(1, field.Length - 2);
                    }
                    else
                    {
                        processedRow[i] = field;
                    }
                }
                processedData.Add(processedRow);
            }
            return processedData;
        }

        static List<IList<object>> ConvertToValueRange(List<string[]> processedData)
        {
            var data = new List<IList<object>>();
            foreach (var row in processedData)
            {
                data.Add(new List<object>(row));
            }
            return data;
        }
        static void AppendDataToSheet(SheetsService service, List<IList<object>> data, string spreadsheetId, string range)
        {
            var valueRange = new ValueRange { Values = data };
            SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            AppendValuesResponse response = request.Execute();
            resRows = (int)response.Updates.UpdatedRows;
        }

    }
}
