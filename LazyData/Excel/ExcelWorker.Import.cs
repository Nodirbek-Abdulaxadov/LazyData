using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LazyData.Excel;

public static class ExcelHelper
{
    /// <summary>
    /// Reads an Excel stream and converts it into a List of objects of type T.
    /// </summary>
    public static List<T> FromExcelStream<T>(this MemoryStream stream, bool skipFirstRow = true) where T : class, new()
    {
        return ReadExcel<T>(SpreadsheetDocument.Open(stream, false), skipFirstRow);
    }

    /// <summary>
    /// Reads an Excel file from a byte array and converts it into a List of objects of type T.
    /// </summary>
    public static List<T> FromExcelByteArray<T>(this byte[] bytes, bool skipFirstRow = true) where T : class, new()
    {
        using MemoryStream stream = new(bytes);
        return ReadExcel<T>(SpreadsheetDocument.Open(stream, false), skipFirstRow);
    }

    /// <summary>
    /// Reads an Excel file from a file path and converts it into a List of objects of type T.
    /// </summary>
    public static List<T> FromExcelFilePath<T>(this IEnumerable<T> source, string path, bool skipFirstRow = true) where T : class, new()
    {
        using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
        source = ReadExcel<T>(document, skipFirstRow);
        return [.. source];
    }

    /// <summary>
    /// Generic method to extract data from an Excel document into a list of objects.
    /// </summary>
    private static List<T> ReadExcel<T>(SpreadsheetDocument document, bool skipFirstRow) where T : class, new()
    {
        List<T> dataList = [];

        // Get the first worksheet
        WorkbookPart workbookPart = document.WorkbookPart!;
        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        // Read header row to map column indexes to properties
        Dictionary<int, PropertyInfo> propertyMap = [];
        Row headerRow = sheetData.Elements<Row>().FirstOrDefault()!;

        if (headerRow != null)
        {
            PropertyInfo[] properties = typeof(T).GetProperties();

            int colIndex = 0;
            foreach (Cell cell in headerRow.Elements<Cell>())
            {
                string headerValue = GetCellValue(workbookPart, cell);
                PropertyInfo property = properties.FirstOrDefault(p => p.Name.Equals(headerValue, StringComparison.OrdinalIgnoreCase))!;

                if (property != null)
                {
                    propertyMap[colIndex] = property;
                }

                colIndex++;
            }
        }

        // Read data rows
        foreach (Row row in sheetData.Elements<Row>().Skip(skipFirstRow ? 1 : 0))
        {
            T obj = new();
            int colIndex = 0;

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (propertyMap.TryGetValue(colIndex, out PropertyInfo? property))
                {
                    string cellValue = GetCellValue(workbookPart, cell);
                    object convertedValue = ConvertValue(cellValue, property.PropertyType);
                    property.SetValue(obj, convertedValue);
                }
                colIndex++;
            }

            dataList.Add(obj);
        }

        return dataList;
    }

    /// <summary>
    /// Extracts the value of a cell, handling shared strings.
    /// </summary>
    private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
    {
        if (cell.CellValue == null) return string.Empty;

        string value = cell.CellValue.Text;

        // Handle shared strings
        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            return workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>()
                   .ElementAt(int.Parse(value)).InnerText;
        }

        return value;
    }

    /// <summary>
    /// Converts an Excel cell value to the appropriate property type.
    /// </summary>
    private static object ConvertValue(string value, Type targetType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return targetType.IsValueType ? Activator.CreateInstance(targetType)! : null!;

        if (targetType == typeof(int)) return int.Parse(value);
        if (targetType == typeof(long)) return long.Parse(value);
        if (targetType == typeof(double)) return double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        if (targetType == typeof(decimal)) return decimal.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        if (targetType == typeof(bool)) return value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase);
        if (targetType == typeof(DateTime)) return DateTime.FromOADate(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture));

        return value;
    }
}