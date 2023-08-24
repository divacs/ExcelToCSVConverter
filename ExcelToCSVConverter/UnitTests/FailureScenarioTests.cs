using System.Net.Http;
using System.Threading.Tasks;
using Moq;
using Xunit;

namespace ExcelToCSVConverterNamespace.Tests
{
    // This test simulates a scenario where the link to the Excel file is not found on the website.
    // The same httpClientMock is used, but now it simulates that there is no link to the Excel file.
    // The test checks whether the return value of the ConvertExcelToCSVAsync method is set to
    // false, indicating an unsuccessful conversion.
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
