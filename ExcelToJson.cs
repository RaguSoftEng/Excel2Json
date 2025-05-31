using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel2Json;
internal class ExcelToJson
{
    public static (string FileName, string Json) ConvertExcelToJson(string filePath)
    {
        try
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new ArgumentException("Invalid file path.");

            var fileName = Path.GetFileNameWithoutExtension(filePath);

            using SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = doc.WorkbookPart ?? throw new Exception("Invalid Document.");
            var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault() ?? throw new Exception("Invalid Document.");
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault() ?? throw new Exception("Invalid Document.");

            var rows = sheetData.Elements<Row>().ToList();
            if (rows.Count == 0) throw new Exception("Excel sheet is empty.");

            var headers = rows.First().Elements<Cell>()
            .Select(cell => GetCellValue(cell, workbookPart))
            .Where(header => !string.IsNullOrWhiteSpace(header))
            .ToArray();


            if (headers.Length == 0) throw new Exception("No headers found in the first row.");

            var data = rows.Skip(1)
                .Select(row => row.Elements<Cell>()
                    .Select((cell, index) => new
                    {
                        Key = headers.ElementAtOrDefault(index) ?? $"Column{index}",
                        Value = GetCellValue(cell, workbookPart)
                    })
                    .Where(x => x.Key != null && !string.IsNullOrWhiteSpace(x.Value))
                    .ToDictionary(x => x.Key, x => x.Value))
                .ToList();

            var result = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });

            return ($"{fileName}.Json", result);
        }
        catch (Exception)
        {
            throw;
        }
    }


    private static string? GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        try
        {
            if (cell == null || cell.CellValue == null) return null;

            return cell.DataType?.Value == CellValues.SharedString
                ? workbookPart.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>()
                    .ElementAtOrDefault(int.Parse(cell.CellValue.Text))?.InnerText ?? null
                : cell.CellValue.Text;
        }
        catch (Exception)
        {
            throw;
        }
    }
}
