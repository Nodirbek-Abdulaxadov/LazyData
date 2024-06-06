using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LazyData;

public static class WordWorker
{
    public static MemoryStream GenerateWordDocument<T>(IEnumerable<T> data)
    {
        MemoryStream memoryStream = new MemoryStream();

        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // Add title
            Paragraph titleParagraph = new Paragraph(new Run(new Text(PluralizationHelper.ToPlural(typeof(T).Name))));
            body.AppendChild(titleParagraph);

            // Add table
            Table table = new Table();
            TableProperties tableProperties = new TableProperties(new TableLayout { Type = TableLayoutValues.Fixed });
            table.AppendChild(tableProperties);

            var properties = typeof(T).GetProperties();

            // Add table header
            TableRow headerRow = new TableRow();
            foreach (var columnHeader in properties)
            {
                TableCell headerCell = new TableCell(new Paragraph(new Run(new Text(columnHeader.Name))));
                headerRow.AppendChild(headerCell);
            }
            table.AppendChild(headerRow);

            // Add table data
            foreach (var item in data)
            {
                TableRow dataRow = new TableRow();
                foreach (var cellData in properties.Select(property => property.GetValue(item)?.ToString()))
                {
                    TableCell dataCell = new TableCell(new Paragraph(new Run(new Text(cellData ?? ""))));
                    dataRow.AppendChild(dataCell);
                }
                table.AppendChild(dataRow);
            }

            body.AppendChild(table);
        }

        return memoryStream;
    }

    public static void SaveAsWordFile<T>(IEnumerable<T> data, string filePath)
    {
        // Generate Word document
        using (MemoryStream stream = GenerateWordDocument(data))
        {
            // Save Word document to file
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
                stream.WriteTo(fileStream);
            }
        }
    }
}