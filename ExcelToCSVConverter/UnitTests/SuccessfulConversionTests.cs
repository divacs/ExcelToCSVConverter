using System.Net.Http;
using System.Threading.Tasks;
using Moq;
using Xunit;

namespace ExcelToCSVConverterNamespace.Tests
{
    // Ovaj test simulira uspešnu konverziju Excel fajla u CSV format.
    // Koristi se httpClientMock kako bi se simuliralo uspešno preuzimanje HTML-a i
    // sadržaja Excel fajla. Test proverava da li se povratna vrednost metode
    // ConvertExcelToCSVAsync postavlja na true, što ukazuje na uspešnu konverziju.
    public class SuccessfulConversionTests
    {
        [Fact]
        public async Task ConvertExcelToCSVAsync_Success()
        {
            // Arrange
            var httpClientMock = new Mock<HttpClient>();
            var converter = new ExcelToCSVConverter(httpClientMock.Object);

            // Simulate a successful conversion
            httpClientMock.Setup(mock => mock.GetStringAsync(It.IsAny<string>())).ReturnsAsync("<html>...</html>");
            httpClientMock.Setup(mock => mock.GetStreamAsync(It.IsAny<string>())).ReturnsAsync(new System.IO.MemoryStream());

            // Act
            bool conversionResult = await converter.ConvertExcelToCSVAsync(
                "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl",
                System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), "TestRigCounts.csv")
            );

            // Assert
            Assert.True(conversionResult);
        }
    }
}
