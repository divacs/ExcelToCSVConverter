using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using ClosedXML.Excel;

namespace ExcelToCSVConverter
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                ThreadPool.SetMinThreads(50, 50);
                string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl";
                string csvFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "RigCounts.csv"); // Fajl će biti sačuvan na desktopu

                using var httpClient = new HttpClient();
                httpClient.Timeout = TimeSpan.FromMinutes(1); // Povećavamo timeout 

                var html = await httpClient.GetStringAsync(url);
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(html);

                var excelLinkNode = htmlDocument.DocumentNode.SelectSingleNode("//a[@title='Worldwide Rig Count Jul 2023.xlsx']");
                if (excelLinkNode != null)
                {
                    string excelLink = excelLinkNode.GetAttributeValue("href", "");
                    Console.WriteLine("Excel link: " + excelLink);

                    var excelStream = await httpClient.GetStreamAsync("https://bakerhughesrigcount.gcs-web.com" + excelLink);

                    using var workbook = new XLWorkbook(excelStream);
                    var worksheet = workbook.Worksheet(1);

                    int rowCount = worksheet.RowsUsed().Count();

                    // Tražimo redove koji pripadaju poslednje 2 godine
                    DateTime twoYearsAgo = DateTime.Now.AddYears(-2);
                    int startRow = 2; // Prvi red sa podacima (preskačemo zaglavlje)
                    int endRow = rowCount; // Zadnji red sa podacima

                    for (int row = startRow; row <= rowCount; row++)
                    {
                        var dateCell = worksheet.Cell(row, 1);
                        if (DateTime.TryParse(dateCell.GetValue<string>(), out DateTime rigDate))
                        {
                            if (rigDate >= twoYearsAgo)
                            {
                                startRow = row; // Postavljamo početni red na prvi red poslednje 2 godine
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

                    Console.WriteLine("CSV fajl je uspešno sačuvan na putanji: " + csvFilePath);
                }
                else
                {
                    Console.WriteLine("Nije pronađen link do Excel fajla.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Došlo je do greške: " + ex.Message);
            }
        }
    }
}
