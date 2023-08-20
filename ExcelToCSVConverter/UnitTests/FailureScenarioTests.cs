using System.Net.Http;
using System.Threading.Tasks;
using Moq;
using Xunit;

namespace ExcelToCSVConverterNamespace.Tests
{
    // Ovaj test simulira scenarij u kojem link ka Excel fajlu nije pronađen na veb sajtu.
    // Koristi se isti httpClientMock, ali sada se simulira da se ne nalazi link ka Excel fajlu.
    // Test proverava da li se povratna vrednost metode ConvertExcelToCSVAsync postavlja
    // na false, što ukazuje na neuspešnu konverziju.
    public class FailureScenarioTests
    {
        [Fact]
        public async Task ConvertExcelToCSVAsync_Failure_NoExcelLink()
        {
            // Arrange
            var httpClientMock = new Mock<HttpClient>();
            var converter = new ExcelToCSVConverter(httpClientMock.Object);

            // Simulate a failed conversion due to no Excel link found
            httpClientMock.Setup(mock => mock.GetStringAsync(It.IsAny<string>())).ReturnsAsync("<html>...</html>");
            // Simulate that no Excel link is found

            // Act
            bool conversionResult = await converter.ConvertExcelToCSVAsync(
                "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=12345&p=irol-rigcountsintl",
                System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), "TestRigCounts.csv")
            );

            // Assert
            Assert.False(conversionResult);
        }
    }
}
