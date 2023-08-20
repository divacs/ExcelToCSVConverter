# Excel to CSV Converter

This console application downloads an Excel file from a specified website, extracts data, and converts it into a CSV (comma-separated values) file. The application uses C# .NET 6/5, HttpClient for web requests, HtmlAgilityPack for HTML parsing, and ClosedXML for working with Excel files.

## Installation

1. Clone this repository to your local machine using: git clone https://github.com/divacs/ExcelToCSVConverter.git
2. Open the solution in your preferred development environment (Visual Studio, Visual Studio Code, etc.).

3. Make sure to restore NuGet packages and ensure you have .NET 6/5 SDK installed.

## Usage

1. Open the `Program.cs` file and set the `url` variable to the URL of the website containing the Excel file.

2. Run the application. The program will download the Excel file, extract data from it, and save it as a CSV file on your desktop.

## Configuration

- The `ThreadPool.SetMinThreads(50, 50)` call is used to set the minimum number of threads in the ThreadPool. Adjust these values according to your system's capabilities and requirements.

- The `httpClient.Timeout` is set to control the maximum time the program will wait for a web request to complete. Modify the timeout value as needed.

## Dependencies

- HtmlAgilityPack: Used for parsing HTML content.
- ClosedXML: Used for working with Excel files.

## Contributing

Feel free to contribute to this project by opening issues or pull requests.

## License

For more details and explanations, you can refer to the source code in `Program.cs`.



