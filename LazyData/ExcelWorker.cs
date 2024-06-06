using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LazyData;

public static class ExcelWorker
{
    public static MemoryStream ToStream<T>(IEnumerable<T> source) where T : class
    {
        MemoryStream memoryStream = new MemoryStream();

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = PluralizationHelper.ToPlural(typeof(T).Name) };
            sheets.Append(sheet);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            // Assuming T is a class with properties
            var properties = typeof(T).GetProperties();

            // Add header row
            Row headerRow = new Row();
            foreach (var property in properties)
            {
                Cell headerCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(property.Name) };
                headerRow.AppendChild(headerCell);
            }
            sheetData!.AppendChild(headerRow);

            // Add data rows
            foreach (var item in source)
            {
                Row dataRow = new Row();
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
                        else if (value is bool)
                        {
                            dataType = CellValues.Boolean;
                            cellValue = (bool)value ? "1" : "0";
                        }
                        else if (value is DateTime)
                        {
                            dataType = CellValues.Date;
                            cellValue = ((DateTime)value).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            cellValue = value.ToString()!;
                        }
                    }

                    Cell dataCell = new Cell() { DataType = dataType, CellValue = new CellValue(cellValue) };
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

    public static void SaveAsExcelFile<T>(this IEnumerable<T> source, string path) where T : class
    {
        // Generate Excel document
        using (MemoryStream stream = ToStream(source))
        {
            // Save Excel document to file
            using (FileStream fileStream = new FileStream(path, FileMode.Create))
            {
                stream.WriteTo(fileStream);
            }
        }
    }
}