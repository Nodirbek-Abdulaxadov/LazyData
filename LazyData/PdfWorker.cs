using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Reflection;

namespace LazyData;

public static class PDFWorker
{
    public static MemoryStream GeneratePDFDocument<T>(IEnumerable<T> data) where T : class
    {
        MemoryStream memoryStream = new MemoryStream();
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");

        // Saqlash PDF fayliga
        SaveAsPdfFile(data, tempPath);

        // PDF faylni oqish
        using (FileStream fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read))
        {
            fileStream.CopyTo(memoryStream);
        }

        // Vaqtinchalik faylni o'chirish
        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }

        memoryStream.Position = 0;
        return memoryStream;
    }

    public static void SaveAsPdfFile<T>(IEnumerable<T> data, string filePath) where T : class
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