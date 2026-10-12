using CustomerApi.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CustomerApi.Services;

/// <summary>
/// Service responsible for generating Word (.docx) documents using DocumentFormat.OpenXml
/// to export hotsynced customer data from Azure SQL.
/// </summary>
public class WordExportService
{
    private const string BrandNavy = "0C1E3D";
    private const string AccentCyan = "0072CE";
    private const string LightGray = "F4F6F8";
    private const string BorderGray = "D0D7DE";

    /// <summary>
    /// Generates a comprehensive Word (.docx) document containing executive summary,
    /// fleet overview table, and detailed profiles for all customers.
    /// </summary>
    /// <param name="customers">The list of customers to include.</param>
    /// <returns>A byte array representing the generated .docx file.</returns>
    public byte[] ExportCustomersToWord(IEnumerable<Customer> customers)
    {
        var list = customers.ToList();
        using var memoryStream = new MemoryStream();

        using (var wordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();

            // 1. Title Banner
            body.Append(CreateTitleParagraph("FENNEX HOTSYNC FLEET REPORT"));
            body.Append(CreateSubtitleParagraph("Customer Operations & Offshore Asset Inventory Export"));
            body.Append(CreateMetadataParagraph(
                $"Generated: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC  |  Total Enterprises: {list.Count}  |  Source: Azure SQL dbo.Customers"
            ));
            body.Append(CreateHorizontalDivider());

            // 2. Executive Fleet Summary
            body.Append(CreateHeading1("1. Executive Fleet Summary"));

            int totalRigs = list.Sum(c => c.NumberOfRigs ?? 0);
            int totalJackUps = list.Sum(c => c.NumberOfJackUps ?? 0);
            int totalModus = list.Sum(c => c.NumberOfModus ?? 0);
            int totalUnits = totalRigs + totalJackUps + totalModus;

            body.Append(CreateSummaryTable(list.Count, totalRigs, totalJackUps, totalModus, totalUnits));
            body.Append(new Paragraph(new Run(new Text(string.Empty)))); // spacing

            // 3. Tabular Overview
            body.Append(CreateHeading1("2. Consolidated Customer Fleet Table"));
            body.Append(CreateCustomerOverviewTable(list));
            body.Append(new Paragraph(new Run(new Text(string.Empty)))); // spacing

            // 4. Detailed Customer Profiles
            body.Append(CreateHeading1("3. Customer Detailed Profiles"));

            int customerIndex = 1;
            foreach (var customer in list)
            {
                body.Append(CreateHeading2($"{customerIndex++}. {customer.Name ?? "Unnamed Enterprise"} (SQL #{customer.Id})"));
                body.Append(CreateCustomerDetailTable(customer));
                body.Append(new Paragraph(new Run(new Text(string.Empty)))); // spacing
            }

            // 5. Confidentiality Footer
            body.Append(CreateHorizontalDivider());
            body.Append(CreateFooterParagraph("CONFIDENTIAL & PROPRIETARY — FENNEX HOTSYNC SYSTEM"));

            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }

        return memoryStream.ToArray();
    }

    /// <summary>
    /// Generates a single customer's detailed executive profile in Word (.docx) format.
    /// </summary>
    public byte[] ExportSingleCustomerToWord(Customer customer)
    {
        using var memoryStream = new MemoryStream();

        using (var wordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();

            body.Append(CreateTitleParagraph("CUSTOMER PROFILE REPORT"));
            body.Append(CreateSubtitleParagraph(customer.Name ?? $"Customer #{customer.Id}"));
            body.Append(CreateMetadataParagraph(
                $"SQL ID: #{customer.Id}  |  SharePoint ID: {customer.SharePointId}  |  Generated: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC"
            ));
            body.Append(CreateHorizontalDivider());

            body.Append(CreateHeading1("Enterprise Overview"));
            body.Append(CreateCustomerDetailTable(customer));

            body.Append(new Paragraph(new Run(new Text(string.Empty))));
            body.Append(CreateHorizontalDivider());
            body.Append(CreateFooterParagraph("CONFIDENTIAL & PROPRIETARY — FENNEX HOTSYNC SYSTEM"));

            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }

        return memoryStream.ToArray();
    }

    // -------------------------------------------------------------------------
    // Helper Methods: Paragraphs & Headings
    // -------------------------------------------------------------------------

    private static Paragraph CreateTitleParagraph(string text)
    {
        var run = new Run(new Text(text))
        {
            RunProperties = new RunProperties(
                new Bold(),
                new Color { Val = BrandNavy },
                new FontSize { Val = "44" }, // 22pt
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
            )
        };

        return new Paragraph(run)
        {
            ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "200", After = "80" }
            )
        };
    }

    private static Paragraph CreateSubtitleParagraph(string text)
    {
        var run = new Run(new Text(text))
        {
            RunProperties = new RunProperties(
                new Color { Val = AccentCyan },
                new FontSize { Val = "26" }, // 13pt
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
            )
        };

        return new Paragraph(run)
        {
            ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "120" }
            )
        };
    }

    private static Paragraph CreateMetadataParagraph(string text)
    {
        var run = new Run(new Text(text))
        {
            RunProperties = new RunProperties(
                new Italic(),
                new Color { Val = "6E7781" },
                new FontSize { Val = "18" }, // 9pt
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
            )
        };

        return new Paragraph(run)
        {
            ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "180" }
            )
        };
    }

    private static Paragraph CreateHeading1(string text)
    {
        var run = new Run(new Text(text))
        {
            RunProperties = new RunProperties(
                new Bold(),
                new Color { Val = BrandNavy },
                new FontSize { Val = "28" }, // 14pt
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
            )
        };

        return new Paragraph(run)
        {
            ParagraphProperties = new ParagraphProperties(
                new SpacingBetweenLines { Before = "260", After = "120" }
            )
        };
    }

    private static Paragraph CreateHeading2(string text)
    {
        var run = new Run(new Text(text))
        {
            RunProperties = new RunProperties(
                new Bold(),
                new Color { Val = AccentCyan },
                new FontSize { Val = "24" }, // 12pt
                new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
            )
        };

        return new Paragraph(run)
        {
            ParagraphProperties = new ParagraphProperties(
                new SpacingBetweenLines { Before = "180", After = "80" }
            )
        };
    }

    private static Paragraph CreateFooterParagraph(string text)
    {
        var run = new Run(new Text(text))
        {
            RunProperties = new RunProperties(
                new Italic(),
                new Color { Val = "8C959F" },
                new FontSize { Val = "16" } // 8pt
            )
        };

        return new Paragraph(run)
        {
            ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "120", After = "120" }
            )
        };
    }

    private static Paragraph CreateHorizontalDivider()
    {
        return new Paragraph
        {
            ParagraphProperties = new ParagraphProperties(
                new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Size = 6, Color = BorderGray }),
                new SpacingBetweenLines { After = "200" }
            )
        };
    }

    // -------------------------------------------------------------------------
    // Helper Methods: Tables
    // -------------------------------------------------------------------------

    private static Table CreateSummaryTable(int totalCustomers, int totalRigs, int totalJackUps, int totalModus, int totalUnits)
    {
        var table = new Table();
        table.Append(CreateTableProperties("9500")); // ~100% width

        var headerRow = new TableRow();
        headerRow.Append(CreateTableCell("Total Customers", isHeader: true, width: "1900"));
        headerRow.Append(CreateTableCell("Oil Rigs", isHeader: true, width: "1900"));
        headerRow.Append(CreateTableCell("Jack-Up Units", isHeader: true, width: "1900"));
        headerRow.Append(CreateTableCell("MODU Vessels", isHeader: true, width: "1900"));
        headerRow.Append(CreateTableCell("Total Fleet", isHeader: true, width: "1900"));
        table.Append(headerRow);

        var dataRow = new TableRow();
        dataRow.Append(CreateTableCell(totalCustomers.ToString(), isHeader: false, width: "1900", isBold: true, alignCenter: true));
        dataRow.Append(CreateTableCell(totalRigs.ToString(), isHeader: false, width: "1900", isBold: true, alignCenter: true));
        dataRow.Append(CreateTableCell(totalJackUps.ToString(), isHeader: false, width: "1900", isBold: true, alignCenter: true));
        dataRow.Append(CreateTableCell(totalModus.ToString(), isHeader: false, width: "1900", isBold: true, alignCenter: true));
        dataRow.Append(CreateTableCell(totalUnits.ToString(), isHeader: false, width: "1900", isBold: true, alignCenter: true, highlight: true));
        table.Append(dataRow);

        return table;
    }

    private static Table CreateCustomerOverviewTable(List<Customer> customers)
    {
        var table = new Table();
        table.Append(CreateTableProperties("9500"));

        var headerRow = new TableRow();
        headerRow.Append(CreateTableCell("ID", isHeader: true, width: "700"));
        headerRow.Append(CreateTableCell("Company Name", isHeader: true, width: "2600"));
        headerRow.Append(CreateTableCell("Address", isHeader: true, width: "2600"));
        headerRow.Append(CreateTableCell("Rigs", isHeader: true, width: "900", alignCenter: true));
        headerRow.Append(CreateTableCell("Jack-Ups", isHeader: true, width: "900", alignCenter: true));
        headerRow.Append(CreateTableCell("MODUs", isHeader: true, width: "900", alignCenter: true));
        headerRow.Append(CreateTableCell("Total", isHeader: true, width: "900", alignCenter: true));
        table.Append(headerRow);

        bool zebra = false;
        foreach (var c in customers)
        {
            zebra = !zebra;
            int rigs = c.NumberOfRigs ?? 0;
            int jackUps = c.NumberOfJackUps ?? 0;
            int modus = c.NumberOfModus ?? 0;
            int total = rigs + jackUps + modus;

            var row = new TableRow();
            row.Append(CreateTableCell($"#{c.Id}", isHeader: false, width: "700", isZebra: zebra));
            row.Append(CreateTableCell(c.Name ?? "—", isHeader: false, width: "2600", isBold: true, isZebra: zebra));
            row.Append(CreateTableCell(c.Address ?? "—", isHeader: false, width: "2600", isZebra: zebra));
            row.Append(CreateTableCell(rigs.ToString(), isHeader: false, width: "900", alignCenter: true, isZebra: zebra));
            row.Append(CreateTableCell(jackUps.ToString(), isHeader: false, width: "900", alignCenter: true, isZebra: zebra));
            row.Append(CreateTableCell(modus.ToString(), isHeader: false, width: "900", alignCenter: true, isZebra: zebra));
            row.Append(CreateTableCell(total.ToString(), isHeader: false, width: "900", alignCenter: true, isBold: true, isZebra: zebra));
            table.Append(row);
        }

        return table;
    }

    private static Table CreateCustomerDetailTable(Customer customer)
    {
        var table = new Table();
        table.Append(CreateTableProperties("9500"));

        int rigs = customer.NumberOfRigs ?? 0;
        int jackUps = customer.NumberOfJackUps ?? 0;
        int modus = customer.NumberOfModus ?? 0;
        int total = rigs + jackUps + modus;

        AddDetailRow(table, "SharePoint List ID", $"SP-{customer.SharePointId}");
        AddDetailRow(table, "SQL Primary Key", $"#{customer.Id}");
        AddDetailRow(table, "Headquarters Address", customer.Address ?? "Not recorded");
        AddDetailRow(table, "Fleet Asset Summary", $"{rigs} Oil Rigs | {jackUps} Jack-Up Units | {modus} MODU Vessels (Total: {total})");
        AddDetailRow(table, "Site / Operational URLs", customer.SiteUrls ?? "None provided");
        AddDetailRow(table, "SharePoint Last Modified", customer.SpLastModified.HasValue ? customer.SpLastModified.Value.ToString("yyyy-MM-dd HH:mm:ss") : "—");
        AddDetailRow(table, "Last Synced to Azure SQL", customer.SyncedAt.ToString("yyyy-MM-dd HH:mm:ss"));
        AddDetailRow(table, "SQL Record Created At", customer.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"));

        return table;
    }

    private static void AddDetailRow(Table table, string label, string value)
    {
        var row = new TableRow();
        row.Append(CreateTableCell(label, isHeader: false, width: "2800", isBold: true, isZebra: true));
        row.Append(CreateTableCell(value, isHeader: false, width: "6700"));
        table.Append(row);
    }

    private static TableProperties CreateTableProperties(string width)
    {
        return new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Dxa, Width = width },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = BorderGray },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = BorderGray },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = BorderGray },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = BorderGray },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = BorderGray },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = BorderGray }
            )
        );
    }

    private static TableCell CreateTableCell(
        string text,
        bool isHeader = false,
        string width = "2000",
        bool isBold = false,
        bool alignCenter = false,
        bool isZebra = false,
        bool highlight = false)
    {
        var runProps = new RunProperties(
            new FontSize { Val = isHeader ? "20" : "18" },
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }
        );

        if (isHeader)
        {
            runProps.Append(new Bold());
            runProps.Append(new Color { Val = "FFFFFF" });
        }
        else if (isBold)
        {
            runProps.Append(new Bold());
        }

        var paraProps = new ParagraphProperties(
            new SpacingBetweenLines { Before = "80", After = "80" }
        );

        if (alignCenter)
        {
            paraProps.Append(new Justification { Val = JustificationValues.Center });
        }

        var para = new Paragraph(paraProps, new Run(runProps, new Text(text ?? string.Empty)));

        var cellProps = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width },
            new TableCellMargin(
                new TopMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "140", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "140", Type = TableWidthUnitValues.Dxa }
            )
        );

        if (isHeader)
        {
            cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = BrandNavy });
        }
        else if (highlight)
        {
            cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "E8F4FD" });
        }
        else if (isZebra)
        {
            cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = LightGray });
        }

        return new TableCell(cellProps, para);
    }
}
