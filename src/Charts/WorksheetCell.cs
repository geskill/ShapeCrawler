using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class WorksheetCell(EmbeddedPackagePart embeddedPackagePart, string sheetName, string address)
{
    internal void UpdateValue(string value)
    {
        UpdateValue(value, CellValues.Number);
    }

    internal void UpdateValue(string value, CellValues type)
    {
        var stream = embeddedPackagePart.GetStream();
        var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, true);
        var xSheet = sdkSpreadsheetDocument.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>()
            .First(xSheet => xSheet.Name == sheetName);
        var sdkWorksheetPart = (WorksheetPart)sdkSpreadsheetDocument.WorkbookPart!.GetPartById(xSheet.Id!);
        var xCells = sdkWorksheetPart.Worksheet!.Descendants<Cell>();
        var xCell = xCells.FirstOrDefault(xCell => xCell.CellReference == address);

        if (xCell != null)
        {
            xCell.DataType = new EnumValue<CellValues>(type);
            xCell.CellValue = new CellValue(value);
        }
        else
        {
            var xWorksheet = sdkWorksheetPart.Worksheet;
            var xSheetData = xWorksheet.Elements<SheetData>().First();
            var rowNumberStr = Regex.Match(address, @"\d+", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
            var rowNumber = int.Parse(rowNumberStr, NumberStyles.Number, NumberFormatInfo.InvariantInfo);
            var xRow = xSheetData.Elements<Row>().First(r => r.RowIndex! == rowNumber);
            var newXCell = new Cell
            {
                CellReference = address,
                DataType = new EnumValue<CellValues>(type),
                CellValue = new CellValue(value)
            };

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            var refCell = xRow.Elements<Cell>().FirstOrDefault(cell =>
                string.Compare(cell.CellReference!.Value, address, true, CultureInfo.InvariantCulture) > 0);

            xRow.InsertBefore(newXCell, refCell);
        }

        sdkSpreadsheetDocument.Dispose();
        stream.Close();
    }
}