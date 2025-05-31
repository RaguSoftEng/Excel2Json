using System.Text;
using System.Text.Json;
using ExcelDataReader;

namespace Excel2Json;
internal class ExcelToJson
{
    // Streaming conversion using ExcelDataReader and Utf8JsonWriter
    public static string ConvertExcelToJsonStreaming(string filePath, string outputDir)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            throw new ArgumentException("Invalid file path.");
        if (string.IsNullOrWhiteSpace(outputDir) || !Directory.Exists(outputDir))
            throw new ArgumentException("Invalid output directory.");

        var fileName = Path.GetFileNameWithoutExtension(filePath) + ".json";
        var outputPath = Path.Combine(outputDir, fileName);

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        using var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var jsonWriter = new Utf8JsonWriter(fs, new JsonWriterOptions { Indented = true });

        if (!reader.Read())
            throw new Exception("Excel sheet is empty.");

        // Read headers
        // This assumes the first row contains headers, if not, it will generate default column names
        // If the first row is empty, it will create default headers like Column0, Column1, etc.
        // This approach allows for handling cases where the first row may not contain valid headers, ensuring that the JSON output still has meaningful keys.
        var headers = new List<string>();
        for (int i = 0; i < reader.FieldCount; i++)
        {
            var header = reader.GetValue(i)?.ToString();
            if (string.IsNullOrWhiteSpace(header))
                header = $"Column{i}";
            headers.Add(header);
        }
        if (headers.Count == 0)
            throw new Exception("No headers found in the first row.");

        jsonWriter.WriteStartArray();
        long rowCount = 0;
        while (reader.Read())
        {
            jsonWriter.WriteStartObject();
            for (int i = 0; i < headers.Count; i++)
            {
                var value = reader.GetValue(i)?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                    jsonWriter.WriteString(headers[i], value);
            }
            jsonWriter.WriteEndObject();
            rowCount++;
            if (rowCount % 1000 == 0)
                Console.Write($"\rProcessed {rowCount} rows...");
        }
        jsonWriter.WriteEndArray();
        jsonWriter.Flush();
        Console.WriteLine($"\rProcessed {rowCount} rows. Done.         ");
        return fileName;
    }
}
