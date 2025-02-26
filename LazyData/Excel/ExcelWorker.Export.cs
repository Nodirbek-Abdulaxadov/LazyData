using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LazyData.Excel;

public static partial class ExcelWorker
{
    /// <summary>
    /// Generates a Excel from a collection of data and returns it as a MemoryStream.
    /// </summary>
    /// <typeparam name="T">The type of data objects.</typeparam>
    /// <param name="data">The collection of data to be converted into a Excel.</param>
    /// <returns>A MemoryStream containing the generated Excel.</returns>
    public static MemoryStream ToExcelStream<T>(this IEnumerable<T> source) where T : class
    {
        MemoryStream memoryStream = new();

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = PluralizationHelper.ToPlural(typeof(T).Name) };
            sheets.Append(sheet);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            // Assuming T is a class with properties
            var properties = typeof(T).GetProperties();

            // Add header row
            Row headerRow = new();
            foreach (var property in properties)
            {
                Cell headerCell = new() { DataType = CellValues.String, CellValue = new CellValue(property.Name) };
                headerRow.AppendChild(headerCell);
            }
            sheetData!.AppendChild(headerRow);

            // Add data rows
            foreach (var item in source)
            {
                Row dataRow = new();
                foreach (var property in properties)
                {
                    var value = property.GetValue(item);
                    CellValues dataType = CellValues.String;
                    string cellValue = string.Empty;

                    if (value != null)
                    {
                        if (value is int || value is long || value is short || value is byte)
                        {
                            dataType = CellValues.Number;
                            cellValue = value.ToString()!;
                        }
                        else if (value is double || value is float || value is decimal)
                        {
                            dataType = CellValues.Number;
                            cellValue = Convert.ToDouble(value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else if (value is bool boolValue)
                        {
                            dataType = CellValues.Boolean;
                            cellValue = boolValue ? "1" : "0";
                        }
                        else if (value is DateTime dateTimeValue)
                        {
                            dataType = CellValues.Date;
                            cellValue = dateTimeValue.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            cellValue = value.ToString()!;
                        }
                    }

                    Cell dataCell = new() { DataType = dataType, CellValue = new(cellValue) };
                    dataRow.AppendChild(dataCell);
                }
                sheetData.AppendChild(dataRow);
            }

            workbookPart.Workbook.Save();
        }

        // Reset stream position to the beginning before returning it
        memoryStream.Position = 0;
        return memoryStream;
    }

    /// <summary>
    /// Converts a given data collection into a Excel byte array.
    /// </summary>
    /// <typeparam name="T">The type of data objects.</typeparam>
    /// <param name="data">The collection of data to be converted into a Excel.</param>
    /// <returns>A byte array containing the generated Excel.</returns>
    public static byte[] ToExcelByteArray<T>(this IEnumerable<T> source) where T : class
    {
        using var memoryStream = new MemoryStream();
        string tempPath = Path.GetTempFileName();

        try
        {
            SaveAsExcelFile(source, tempPath);
            return File.ReadAllBytes(tempPath);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    /// <summary>
    /// Saves a given data collection as a Excel file at the specified path.
    /// </summary>
    /// <typeparam name="T">The type of data objects.</typeparam>
    /// <param name="data">The collection of data to be written to the Excel.</param>
    /// <param name="filePath">The file path where the Excel will be saved.</param>
    public static void SaveAsExcelFile<T>(this IEnumerable<T> source, string path) where T : class
    {
        using MemoryStream stream = ToExcelStream(source);
        using FileStream fileStream = new(path, FileMode.Create);
        stream.WriteTo(fileStream);
    }
}