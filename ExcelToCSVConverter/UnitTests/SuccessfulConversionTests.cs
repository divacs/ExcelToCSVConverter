using System.Net.Http;
using System.Threading.Tasks;
using Moq;
using Xunit;

namespace ExcelToCSVConverterNamespace.Tests
{
    // This test simulates a successful conversion of an Excel file into CSV format.
    // The httpClientMock is used to simulate a successful download of HTML and
    // the content of the Excel file. The test verifies whether the return value of the
    // ConvertExcelToCSVAsync method is set to true, indicating a successful conversion.
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
