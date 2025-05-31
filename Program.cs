using Excel2Json;

var excelInput = @"D:\Test\Inventory.xlsx"; // ExcelFilePath
var OutputPath = @"D:\Test"; // Output dir

try
{
    Console.WriteLine("Processing......");

    var (fileName, json) = ExcelToJson.ConvertExcelToJson(excelInput);
    File.WriteAllText($"{OutputPath}\\{fileName}", json);

    Console.WriteLine("Successfully converted.");
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}