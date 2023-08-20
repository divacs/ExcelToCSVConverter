using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using ClosedXML.Excel;

namespace ExcelToCSVConverterNamespace
{
    public class ExcelToCSVConverter
    {
        private readonly HttpClient _httpClient;

        public ExcelToCSVConverter(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<bool> ConvertExcelToCSVAsync(string url, string csvFilePath)
        {
            try
            {
                ThreadPool.SetMinThreads(50, 50);

                using var httpClient = _httpClient;
                httpClient.Timeout = TimeSpan.FromMinutes(1);

                var html = await httpClient.GetStringAsync(url);
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(html);

                var excelLinkNode = htmlDocument.DocumentNode.SelectSingleNode("//a[@title='Worldwide Rig Count Jul 2023.xlsx']");
                if (excelLinkNode != null)
                {
                    string excelLink = excelLinkNode.GetAttributeValue("href", "");

                    var excelStream = await httpClient.GetStreamAsync("https://bakerhughesrigcount.gcs-web.com" + excelLink);

                    using var workbook = new XLWorkbook(excelStream);
                    var worksheet = workbook.Worksheet(1);

                    int rowCount = worksheet.RowsUsed().Count();

                    DateTime twoYearsAgo = DateTime.Now.AddYears(-2);
                    int startRow = 2;
                    int endRow = rowCount;

                    // Tražimo redove koji pripadaju poslednje 2 godine
                    for (int row = startRow; row <= rowCount; row++)
                    {
                        var dateCell = worksheet.Cell(row, 1);
                        if (DateTime.TryParse(dateCell.GetValue<string>(), out DateTime rigDate))
                        {
                            if (rigDate >= twoYearsAgo)
                            {
                                startRow = row;
                                break;
                            }
                        }
                    }

                    // Kreiramo CSV fajl i upisujemo podatke
                    using var csvFile = new StreamWriter(csvFilePath);
                    for (int row = startRow; row <= endRow; row++)
                    {
                        var csvLine = string.Join(",", worksheet.Row(row).Cells());
                        csvFile.WriteLine(csvLine);
                    }

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
    }

    public class Program
    {
        static async Task Main(string[] args)
        {
            // Kreiranje instance HttpClient i ExcelToCSVConverter
            using var httpClient = new HttpClient();
            var converter = new ExcelToCSVConverter(httpClient);

            // Poziv metode za konverziju
            bool conversionResult = await converter.ConvertExcelToCSVAsync(
                "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TestRigCounts.csv")
            );

            if (conversionResult)
            {
                Console.WriteLine("Konverzija uspešno završena.");
            }
            else
            {
                Console.WriteLine("Došlo je do greške pri konverziji.");
            }
        }
    }
}
