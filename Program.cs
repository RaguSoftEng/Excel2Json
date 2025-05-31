using System.Text;
using Excel2Json;

// Register code page provider for legacy Excel encodings
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var excelInput = @"D:\Test\Inventory.xlsx"; // ExcelFilePath
var OutputPath = @"D:\Test"; // Output dir

try
{
    Console.WriteLine("Processing......");

    var fileName = ExcelToJson.ConvertExcelToJsonStreaming(excelInput, OutputPath);

    Console.WriteLine("Successfully converted.");
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}