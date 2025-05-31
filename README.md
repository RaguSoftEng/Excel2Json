# **Excel to JSON Converter - .NET Core**

## **Overview**
This repository provides a **dynamic solution** for converting **Excel files to JSON** using **.NET Core and OpenXML**. It extracts data efficiently and transforms it into structured JSON format without requiring third-party dependencies.

## **Features**
- ✅ **Dynamic Column Mapping** – Automatically detects headers and assigns values.
- ✅ **No External Libraries** – Uses **pure .NET Core & OpenXML** for compatibility.
- ✅ **Null & Error Handling** – Ensures clean and accurate JSON output.
- ✅ **Optimized for Large Files** – Uses **streaming techniques** to minimize memory usage.

## **How It Works**
1. **Reads an Excel file** using OpenXML.
2. **Extracts headers dynamically** from the first row.
3. **Maps each row to JSON objects** based on column names.
4. **Handles missing values** gracefully.
5. **Saves JSON output** for easy processing.

## **Getting Started**

### **Prerequisites**
- .NET 9  
- OpenXML

### **Installation**
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/excel-to-json.git
   cd excel-to-json
   ```
2. Open in Visual Studio or any .NET-supported IDE.
3. Build and run the project.

### **Usage**

Modify Program.cs to specify the Excel file path & Output dir:
```csharp
var excelInput = @"D:\Test\Inventory.xlsx"; // ExcelFilePath
var OutputPath = @"D:\Test"; // Output dir
```

Run the program:
```bash
dotnet run
```

Example output:
``` JSON
[
  {
    "Item ID": "101",
    "Name": "Laptop",
    "Category": "Electronics",
    "Stock": "25",
    "Price": "1200"
  }
]
```

## **Contributing**
Contributions are welcome! Feel free to submit a pull request or suggest improvements.