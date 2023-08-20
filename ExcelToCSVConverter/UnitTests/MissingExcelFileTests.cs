using System.Net.Http;
using System.Threading.Tasks;
using Moq;
using Xunit;

namespace ExcelToCSVConverterNamespace.Tests
{
    // Ovaj test simulira situaciju kada Excel fajl nije pronađen ili nije mogao da se preuzme.
    // Korišćenjem httpClientMock simulira se prazan odgovor za preuzimanje sadržaja
    // Excel fajla kako bi se proverilo kako metoda reaguje na ovakvu situaciju.
    // Test proverava da li se povratna vrednost metode ConvertExcelToCSVAsync
    // postavlja na false, što ukazuje na neuspešnu konverziju.
    public class MissingExcelFileTests
    {
        [Fact]
        public async Task ConvertExcelToCSVAsync_Failure_MissingExcelFile()
        {
            // Arrange
            var httpClientMock = new Mock<HttpClient>();
            var converter = new ExcelToCSVConverter(httpClientMock.Object);

            // Simulate a scenario where the Excel file is missing
            httpClientMock.Setup(mock => mock.GetStringAsync(It.IsAny<string>())).ReturnsAsync("<html>...</html>");
            httpClientMock.Setup(mock => mock.GetStreamAsync(It.IsAny<string>())).ReturnsAsync((System.IO.Stream)null);

            // Act
            bool conversionResult = await converter.ConvertExcelToCSVAsync(
                "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl",
                System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), "TestRigCounts.csv")
            );

            // Assert
            Assert.False(conversionResult);
        }
    }
}
