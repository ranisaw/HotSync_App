using CustomerApi.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CustomerApi.Services;

/// <summary>
/// Service responsible for generating Excel (.xlsx) spreadsheets using DocumentFormat.OpenXml
/// to export hotsynced customer data from Azure SQL.
/// </summary>
public class ExcelExportService
{
    /// <summary>
    /// Generates a formatted Excel (.xlsx) workbook containing customer data in a tabular layout.
    /// </summary>
    /// <param name="customers">List of customers to export.</param>
    /// <returns>A byte array containing the .xlsx file.</returns>
    public byte[] ExportCustomersToExcel(IEnumerable<Customer> customers)
    {
        using var memoryStream = new MemoryStream();

        using (var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add Stylesheet Part (Fonts, Fills, Borders, CellFormats)
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateStylesheet();
            stylesPart.Stylesheet.Save();

            // Add Worksheet Part
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();

            // Configure Column Widths for readability
            var columns = new Columns();
            columns.Append(new Column { Min = 1, Max = 1, Width = 8, CustomWidth = true });   // ID
            columns.Append(new Column { Min = 2, Max = 2, Width = 14, CustomWidth = true });  // SP ID
            columns.Append(new Column { Min = 3, Max = 3, Width = 28, CustomWidth = true });  // Company Name
            columns.Append(new Column { Min = 4, Max = 4, Width = 36, CustomWidth = true });  // Address
            columns.Append(new Column { Min = 5, Max = 5, Width = 14, CustomWidth = true });  // Oil Rigs
            columns.Append(new Column { Min = 6, Max = 6, Width = 14, CustomWidth = true });  // Jack-Ups
            columns.Append(new Column { Min = 7, Max = 7, Width = 14, CustomWidth = true });  // MODUs
            columns.Append(new Column { Min = 8, Max = 8, Width = 16, CustomWidth = true });  // Total Fleet
            columns.Append(new Column { Min = 9, Max = 9, Width = 40, CustomWidth = true });  // Site URLs
            columns.Append(new Column { Min = 10, Max = 10, Width = 22, CustomWidth = true }); // SP Modified
            columns.Append(new Column { Min = 11, Max = 11, Width = 22, CustomWidth = true }); // Synced At
            columns.Append(new Column { Min = 12, Max = 12, Width = 22, CustomWidth = true }); // Created At

            var worksheet = new Worksheet();
            worksheet.Append(columns);
            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;

            // Define Sheet in Workbook
            var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Hotsynced Customers"
            };
            sheets.Append(sheet);

            // Row 1: Header Row
            uint rowIndex = 1;
            var headerRow = new Row { RowIndex = rowIndex };

            string[] headers =
            [
                "SQL Id",
                "SharePoint ID",
                "Company Name",
                "Address",
                "Oil Rigs",
                "Jack-Up Units",
                "MODU Vessels",
                "Total Fleet",
                "Site URLs",
                "SP Last Modified",
                "Synced At",
                "Created At"
            ];

            for (int i = 0; i < headers.Length; i++)
            {
                headerRow.Append(CreateTextCell(headers[i], styleIndex: 1)); // Style 1: Bold with Navy Fill
            }
            sheetData.Append(headerRow);

            // Data Rows
            foreach (var customer in customers)
            {
                rowIndex++;
                var row = new Row { RowIndex = rowIndex };

                uint styleIdx = (rowIndex % 2 == 0) ? 2U : 3U; // Alternating zebra styles

                int rigs = customer.NumberOfRigs ?? 0;
                int jackUps = customer.NumberOfJackUps ?? 0;
                int modus = customer.NumberOfModus ?? 0;
                int totalFleet = rigs + jackUps + modus;

                row.Append(CreateNumberCell(customer.Id, styleIdx));
                row.Append(CreateNumberCell(customer.SharePointId, styleIdx));
                row.Append(CreateTextCell(customer.Name ?? string.Empty, styleIdx));
                row.Append(CreateTextCell(customer.Address ?? string.Empty, styleIdx));
                row.Append(CreateNumberCell(rigs, styleIdx));
                row.Append(CreateNumberCell(jackUps, styleIdx));
                row.Append(CreateNumberCell(modus, styleIdx));
                row.Append(CreateNumberCell(totalFleet, styleIdx));
                row.Append(CreateTextCell(customer.SiteUrls ?? string.Empty, styleIdx));
                row.Append(CreateTextCell(customer.SpLastModified.HasValue ? customer.SpLastModified.Value.ToString("yyyy-MM-dd HH:mm:ss") : "—", styleIdx));
                row.Append(CreateTextCell(customer.SyncedAt.ToString("yyyy-MM-dd HH:mm:ss"), styleIdx));
                row.Append(CreateTextCell(customer.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"), styleIdx));

                sheetData.Append(row);
            }

            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }

        return memoryStream.ToArray();
    }

    private static Cell CreateTextCell(string text, uint styleIndex)
    {
        return new Cell
        {
            DataType = CellValues.InlineString,
            StyleIndex = styleIndex,
            InlineString = new InlineString(new Text(text ?? string.Empty))
        };
    }

    private static Cell CreateNumberCell(int number, uint styleIndex)
    {
        return new Cell
        {
            DataType = CellValues.Number,
            StyleIndex = styleIndex,
            CellValue = new CellValue(number.ToString())
        };
    }

    private static Stylesheet CreateStylesheet()
    {
        return new Stylesheet(
            new Fonts(
                // Index 0: Normal Font (Calibri 11)
                new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" }),
                // Index 1: Header Font (Calibri 11 Bold White)
                new Font(
                    new Bold(),
                    new FontSize { Val = 11 },
                    new Color { Rgb = "FFFFFFFF" },
                    new FontName { Val = "Calibri" }
                )
            ),
            new Fills(
                // Index 0: Default None
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                // Index 1: Default Gray125
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                // Index 2: Header Deep Navy Fill #0C1E3D
                new Fill(new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = "FF0C1E3D" }
                }),
                // Index 3: Zebra Light Blue-Gray Fill #F0F4FA
                new Fill(new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = "FFF0F4FA" }
                })
            ),
            new Borders(
                // Index 0: Default border
                new Border(),
                // Index 1: Thin border all sides
                new Border(
                    new LeftBorder(new Color { Rgb = "FFD0D7DE" }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color { Rgb = "FFD0D7DE" }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color { Rgb = "FFD0D7DE" }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color { Rgb = "FFD0D7DE" }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder()
                )
            ),
            new CellFormats(
                // Index 0: Default
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0 },
                // Index 1: Header (Bold, White font, Navy fill, thin border)
                new CellFormat
                {
                    FontId = 1,
                    FillId = 2,
                    BorderId = 1,
                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = true,
                    Alignment = new Alignment { Vertical = VerticalAlignmentValues.Center }
                },
                // Index 2: Data row zebra light fill
                new CellFormat
                {
                    FontId = 0,
                    FillId = 3,
                    BorderId = 1,
                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = true,
                    Alignment = new Alignment { Vertical = VerticalAlignmentValues.Center }
                },
                // Index 3: Data row white fill
                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 1,
                    ApplyFont = true,
                    ApplyBorder = true,
                    Alignment = new Alignment { Vertical = VerticalAlignmentValues.Center }
                }
            )
        );
    }
}
