# Excel2Json Application Specification

## Overview
Excel2Json is a .NET 9 console application that converts Excel spreadsheet files (.xlsx) into JSON format. It is designed to process a single Excel file and output a JSON file representing the data in the first worksheet.

## Core Features
- **Excel to JSON Conversion**: Reads an Excel (.xlsx) file and serializes its first worksheet's data into a JSON array of objects.
- **Header Mapping**: Uses the first row of the worksheet as property names (headers) for the JSON objects.
- **Data Extraction**: Each subsequent row is converted into a JSON object, mapping cell values to the corresponding headers.
- **File Output**: Writes the resulting JSON to a file in a specified output directory, naming the file after the original Excel file.
- **Error Handling**: Provides basic error messages for invalid file paths, empty sheets, or missing headers.

## User Scenarios
- **Single File Conversion**: A user specifies the path to an Excel file and an output directory. The application processes the file and creates a JSON file in the output directory.
- **Data Integration**: Users can use the generated JSON for further processing, integration, or import into other systems.
- **Error Feedback**: If the input file is invalid or the Excel sheet is empty, the user receives a descriptive error message.

## Usage Flow
1. User sets the input Excel file path and output directory in the application.
2. The application reads the first worksheet, extracts headers and data, and serializes them to JSON.
3. The JSON file is saved to the output directory with a name based on the Excel file.
4. Success or error messages are displayed in the console.

## Limitations
- Only the first worksheet is processed.
- Only supports .xlsx files (OpenXML format).
- Requires valid headers in the first row.
- No support for advanced Excel features (formulas, formatting, multiple sheets).

## Example
Given an Excel file `Inventory.xlsx` with the following content:

| Name   | Quantity | Price |
|--------|----------|-------|
| Apple  | 10       | 1.2   |
| Banana | 5        | 0.8   |

The output JSON will be:
```
[
  { "Name": "Apple", "Quantity": "10", "Price": "1.2" },
  { "Name": "Banana", "Quantity": "5", "Price": "0.8" }
]
```
