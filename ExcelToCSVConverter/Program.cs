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
                // Set the minimum number of threads in the ThreadPool
                ThreadPool.SetMinThreads(50, 50);

                // Create and configure an instance of HttpClient
                using var httpClient = _httpClient;
                httpClient.Timeout = TimeSpan.FromMinutes(1);

                // Download the HTML content from the provided URL
                var html = await httpClient.GetStringAsync(url);
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(html);

                // Find the anchor element containing the specified text
                var excelLinkNode = htmlDocument.DocumentNode.SelectSingleNode("//a[contains(text(), 'Worldwide Rig Counts - Current & Historical Data')]");
                if (excelLinkNode != null)
                {
                    // Get the 'href' attribute value from the anchor element
                    string excelLink = excelLinkNode.GetAttributeValue("href", "");

                    // Download the Excel file using the obtained link
                    var excelStream = await httpClient.GetStreamAsync("https://bakerhughesrigcount.gcs-web.com" + excelLink);

                    // Load the Excel file into a ClosedXML workbook
                    using var workbook = new XLWorkbook(excelStream);
                    var worksheet = workbook.Worksheet(1);

                    int rowCount = worksheet.RowsUsed().Count();

                    DateTime twoYearsAgo = DateTime.Now.AddYears(-2);
                    int startRow = 2;
                    int endRow = rowCount;

                    // Find rows that belong to the last 2 years
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

                    // Create a CSV file and write data into it
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
            // Create an instance of HttpClient and ExcelToCSVConverter
            using var httpClient = new HttpClient();
            var converter = new ExcelToCSVConverter(httpClient);

            // Call the method for conversion
            bool conversionResult = await converter.ConvertExcelToCSVAsync(
                "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TestRigCounts.csv")
            );

            if (conversionResult)
            {
                Console.WriteLine("Conversion completed successfully.");
            }
            else
            {
                Console.WriteLine("An error occurred during conversion.");
            }
        }
    }
}
