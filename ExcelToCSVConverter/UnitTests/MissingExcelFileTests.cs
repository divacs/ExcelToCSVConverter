using System.Net.Http;
using System.Threading.Tasks;
using Moq;
using Xunit;

namespace ExcelToCSVConverterNamespace.Tests
{
    // This test simulates a situation where the Excel file is not found or could not be downloaded.
    // Using httpClientMock, an empty response for downloading the content of the Excel file is simulated
    // to verify how the method reacts to such a situation.
    // The test checks whether the return value of the ConvertExcelToCSVAsync method
    // is set to false, indicating an unsuccessful conversion.
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
