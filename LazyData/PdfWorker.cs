using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Reflection;

namespace LazyData;

public static class PDFWorker
{
    /// <summary>
    /// Generates a PDF from a collection of data and returns it as a MemoryStream.
    /// </summary>
    /// <typeparam name="T">The type of data objects.</typeparam>
    /// <param name="data">The collection of data to be converted into a PDF.</param>
    /// <returns>A MemoryStream containing the generated PDF.</returns>
    public static MemoryStream ToPdfStrem<T>(this IEnumerable<T> data) where T : class
    {
        MemoryStream memoryStream = new();
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");

        SaveAsPdfFile(data, tempPath);

        using (FileStream fileStream = new(tempPath, FileMode.Open, FileAccess.Read))
        {
            fileStream.CopyTo(memoryStream);
        }

        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }

        memoryStream.Position = 0;
        return memoryStream;
    }

    /// <summary>
    /// Converts a given data collection into a PDF byte array.
    /// </summary>
    /// <typeparam name="T">The type of data objects.</typeparam>
    /// <param name="data">The collection of data to be converted into a PDF.</param>
    /// <returns>A byte array containing the generated PDF.</returns>
    public static byte[] ToPdfByteArray<T>(this IEnumerable<T> data) where T : class
    {
        using var memoryStream = new MemoryStream();
        string tempPath = Path.GetTempFileName();

        try
        {
            SaveAsPdfFile(data, tempPath);
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
    /// Saves a given data collection as a PDF file at the specified path.
    /// </summary>
    /// <typeparam name="T">The type of data objects.</typeparam>
    /// <param name="data">The collection of data to be written to the PDF.</param>
    /// <param name="filePath">The file path where the PDF will be saved.</param>
    public static void SaveAsPdfFile<T>(this IEnumerable<T> data, string filePath) where T : class
    {
        QuestPDF.Settings.License = LicenseType.Community;

        var properties = typeof(T).GetProperties();

        var checkData = CanPortrait(properties, data);

        PageSize pageSize = PageSizes.A3;

        if (checkData.Item1)
        {
            pageSize = PageSizes.A4;
        }

        Document.Create(container =>
        {
            _ = container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.MarginHorizontal(15);
                page.MarginVertical(20);

                page.Header()
                    .Height(30)
                    .AlignCenter()
                    .AlignMiddle()
                    .Text(PluralizationHelper.ToPlural(typeof(T).Name))
                    .FontSize(14);

                page.Content()
                    .AlignCenter()
                    .Column(column =>
                    {
                        column.Spacing(20);

                        column.Item().Border(1).Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                foreach (var prop in checkData.Item2)
                                {
                                    if (prop.Value == 0 || !data.Any())
                                    {
                                        columns.RelativeColumn();
                                    }
                                    else
                                    {
                                        columns.ConstantColumn(prop.Value);
                                    }
                                }
                            });

                            foreach (var prop in properties)
                            {
                                table.Cell().Element(HeaderStyle).Text(prop.Name).Bold().FontSize(11);
                            }

                            foreach (var item in data)
                            {
                                foreach (var prop in properties)
                                {
                                    var value = prop.GetValue(item);
#pragma warning disable CS0618 // Type or member is obsolete
                                    table.Cell().Element(CellStyle).Text(value ?? "").FontSize(10);
#pragma warning restore CS0618 // Type or member is obsolete
                                }
                            }

                            static IContainer HeaderStyle(IContainer x) => x.Border(1).Padding(5);
                            static IContainer CellStyle(IContainer x) => x.Border(0.5f).Padding(2);
                        });
                    });

                page.Footer()
                    .Height(20)
                    .AlignCenter()
                    .AlignMiddle()
                    .Text(text =>
                    {
                        text.DefaultTextStyle(TextStyle.Default.FontSize(8));
                        text.CurrentPageNumber();
                        text.Span(" / ");
                        text.TotalPages();
                    });
            });
        })
        .GeneratePdf(filePath);
    }

    private static (bool, Dictionary<PropertyInfo, float>) CanPortrait<T>(PropertyInfo[] properties, IEnumerable<T> data) where T : class
    {
        Dictionary<PropertyInfo, float> sizes = new();

        int width = 0;
        foreach (var property in properties)
        {
            int maxLength = 0;

            foreach (var item in data)
            {
                var value = property.GetValue(item)?.ToString();
                if (value != null && value.Length > maxLength)
                {
                    maxLength = value.Length;
                }
            }

            if (property.PropertyType == typeof(string))
            {
                maxLength = 0;
            }

            width += maxLength * 8;
            sizes.Add(property, maxLength * 10);
        }

        return (width <= 720, sizes);
    }
}