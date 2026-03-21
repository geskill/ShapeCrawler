using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents table.
/// </summary>
public interface ITable
{
    /// <summary>
    ///     Gets table columns.
    /// </summary>
    ITableColumnCollection Columns { get; }

    /// <summary>
    ///     Gets table rows.
    /// </summary>
    ITableRowCollection Rows { get; }

    /// <summary>
    ///     Gets or sets the table style.
    /// </summary>
    ITableStyle TableStyle { get; set; }

    /// <summary>
    ///     Gets the table style options.
    /// </summary>
    ITableStyleOptions StyleOptions { get; }

    /// <summary>
    ///     Gets cell by row and column indexes.
    /// </summary>
    ITableCell this[int rowIndex, int columnIndex] { get; }

    /// <summary>
    ///     Merge neighbor cells.
    /// </summary>
    void MergeCells(ITableCell cell1, ITableCell cell2);

    /// <summary>
    ///     Updates table fill.
    /// </summary>
    void UpdateFill(string colorHex);

    /// <summary>
    ///     Gets table cell by row and column numbers.
    /// </summary>
    ITableCell Cell(int rowNumber, int columnNumber);
}

internal sealed class Table(
    TableRowCollection rows,
    TableColumnCollection columns,
    TableStyleOptions styleOptions,
    P.GraphicFrame pGraphicFrame) : ITable
{
    private ITableStyle? tableStyle;

    private A.Table ATable => pGraphicFrame.GetFirstChild<A.Graphic>()!.GraphicData!.GetFirstChild<A.Table>()!;

    public ITableColumnCollection Columns => columns;

    public ITableRowCollection Rows => rows;

    public ITableStyle TableStyle
    {
        get => GetTableStyle();
        set => SetTableStyle(value);
    }

    public ITableStyleOptions StyleOptions => styleOptions;

    public ITableCell this[int rowIndex, int columnIndex] => Rows[rowIndex].Cells[columnIndex];

    public void UpdateFill(string colorHex)
    {
        throw new NotImplementedException();
    }

    public ITableCell Cell(int rowNumber, int columnNumber)
    {
        return Rows[rowNumber - 1].Cells[columnNumber - 1];
    }

    public void MergeCells(ITableCell cell1, ITableCell cell2)
    {
        var cell1Internal = (TableCell)cell1;
        var cell2Internal = (TableCell)cell2;
        if (cell1Internal == cell2Internal)
        {
            throw new SCException("Cannot merge the same cells.");
        }

        var minRowIndex = cell1Internal.RowIndex < cell2Internal.RowIndex
            ? cell1Internal.RowIndex
            : cell2Internal.RowIndex;
        var maxRowIndex = cell1Internal.RowIndex > cell2Internal.RowIndex
            ? cell1Internal.RowIndex
            : cell2Internal.RowIndex;
        var minColIndex = cell1Internal.ColumnIndex < cell2Internal.ColumnIndex
            ? cell1Internal.ColumnIndex
            : cell2Internal.ColumnIndex;
        var maxColIndex = cell1Internal.ColumnIndex > cell2Internal.ColumnIndex
            ? cell1Internal.ColumnIndex
            : cell2Internal.ColumnIndex;

        var aTableRows = ATable.Elements<A.TableRow>().ToList();
        if (minColIndex != maxColIndex)
        {
            MergeHorizontal(maxColIndex, minColIndex, minRowIndex, maxRowIndex, aTableRows);
        }

        if (minRowIndex != maxRowIndex)
        {
            MergeVertically(maxRowIndex, minRowIndex, aTableRows, minColIndex, maxColIndex);
        }

        RemoveColumnIfNeeded(aTableRows);
        RemoveRowOnDemand();
    }

    public void Remove()
    {
        pGraphicFrame.Remove();
    }

    public void SetVideo(Stream video)
    {
        throw new NotImplementedException();
    }

    private static bool IsParagraphEmpty(A.Paragraph aParagraph)
    {
        return aParagraph.Descendants<A.Text>().All(t => string.IsNullOrEmpty(t.Text));
    }

    private static void DeleteTableCells(int colIdx, int deleteColumnCount, List<A.TableRow> aTableRows)
    {
        foreach (var aTblRow in aTableRows)
        {
            var removeCells = aTblRow.Elements<A.TableCell>().Skip(colIdx + 1).Take(deleteColumnCount).ToList();
            foreach (var aTblCell in removeCells)
            {
                aTblCell.Remove();
            }
        }
    }

    private void SetTableStyle(ITableStyle style)
    {
        ATable.TableProperties!.GetFirstChild<A.TableStyleId>()!.Text = ((TableStyle)style).Guid;
        tableStyle = style;
    }

    private ITableStyle GetTableStyle()
    {
        if (tableStyle is null)
        {
            var aTableStyleId = ATable.TableProperties!.GetFirstChild<A.TableStyleId>()!.Text;
            var style = CommonTableStyles.GetTableStyleByGuid(aTableStyleId)!;
            tableStyle = style;
        }

        return tableStyle;
    }

    private void RemoveRowOnDemand()
    {
        var rowIdx = 0;

        while (rowIdx < Rows.Count)
        {
            var cells = Rows[rowIdx].Cells.OfType<TableCell>().ToList();
            var firstCell = cells[0];
            var firstCellSpan = firstCell.ATableCell.RowSpan?.Value;

            if (firstCellSpan > 1 && cells.All(cell => cell.ATableCell.RowSpan?.Value == firstCellSpan))
            {
                var deleteRowsCount = firstCellSpan.Value - 1;
                var targetRow = (TableRow)Rows[rowIdx];
                var newHeight = targetRow.Height;
                foreach (var row in Rows.Skip(rowIdx + 1).Take(deleteRowsCount))
                {
                    ((TableRow)row).ATableRow.Remove();
                    newHeight += row.Height;
                }

                // Update the target row height directly without affecting the table shape size
                targetRow.SetHeight(newHeight);

                rowIdx += firstCellSpan.Value;
            }
            else
            {
                rowIdx++;
            }
        }
    }

    private void MergeVertically(
        int bottomIndex,
        int topRowIndex,
        List<A.TableRow> aTableRows,
        int leftColIndex,
        int rightColIndex)
    {
        var verticalMergingCount = bottomIndex - topRowIndex + 1;
        var numMergingCells = rightColIndex - leftColIndex + 1;
        var horizontalCells =
            aTableRows[topRowIndex].Elements<A.TableCell>().Skip(leftColIndex).Take(numMergingCells);
        foreach (var aTblCell in horizontalCells)
        {
            aTblCell.RowSpan = new Int32Value(verticalMergingCount);
        }

        // Set vertical merging flag
        foreach (var aTableRow in aTableRows.Skip(topRowIndex + 1).Take(bottomIndex - topRowIndex))
        {
            foreach (var aTc in aTableRow.Elements<A.TableCell>().Skip(leftColIndex)
                         .Take(rightColIndex - leftColIndex + 1))
            {
                aTc.VerticalMerge = new BooleanValue(true);
                MergeParagraphs(topRowIndex, leftColIndex, aTc);
            }
        }
    }

    private void MergeParagraphs(int minRowIndex, int minColIndex, A.TableCell aTblCell)
    {
        var mergedCellTextBody = ((TableCell)this[minRowIndex, minColIndex]).ATableCell.TextBody;
        var hasMoreOnePara = false;
        var aParagraphsWithARun = aTblCell.TextBody!.Elements<A.Paragraph>().Where(p => !IsParagraphEmpty(p));
        foreach (var aParagraph in aParagraphsWithARun)
        {
            mergedCellTextBody!.Append(aParagraph.CloneNode(true));
            hasMoreOnePara = true;
        }

        if (!hasMoreOnePara)
        {
            return;
        }

        foreach (var aParagraph in mergedCellTextBody!.Elements<A.Paragraph>().Where(IsParagraphEmpty))
        {
            aParagraph.Remove();
        }
    }

    private void MergeHorizontal(
        int maxColIndex,
        int minColIndex,
        int minRowIndex,
        int maxRowIndex,
        List<A.TableRow> aTableRows)
    {
        var horizontalMergingCount = maxColIndex - minColIndex + 1;
        for (var rowIdx = minRowIndex; rowIdx <= maxRowIndex; rowIdx++)
        {
            A.TableCell[] rowATblCells = [.. aTableRows[rowIdx].Elements<A.TableCell>()];
            var firstMergingCell = rowATblCells[minColIndex];
            firstMergingCell.GridSpan = new Int32Value(horizontalMergingCount);
            var nextMergingCells =
                new Span<A.TableCell>(rowATblCells, minColIndex + 1, horizontalMergingCount - 1);
            foreach (var aTblCell in nextMergingCells)
            {
                aTblCell.HorizontalMerge = new BooleanValue(true);
                MergeParagraphs(minRowIndex, minColIndex, aTblCell);
            }
        }
    }

    private void RemoveColumnIfNeeded(List<A.TableRow> aTableRows)
    {
        var colIdx = 0;
        while (colIdx < Columns.Count)
        {
            var topColumnCell = ((TableRow)Rows[0]).ATableRow.Elements<A.TableCell>().ToList()[colIdx];
            var topColumnCellSpan = topColumnCell.GridSpan?.Value;
            var nextBottomColumnCells = Rows
                .Select(row => ((TableRow)row).ATableRow.Elements<A.TableCell>().ToList()[colIdx]).ToList();
            var sameGridSpan = nextBottomColumnCells.All(c => c.GridSpan?.Value == topColumnCellSpan);
            if (topColumnCellSpan > 1 && sameGridSpan)
            {
                colIdx += ProcessColumnsWithSameGridSpan(colIdx, topColumnCellSpan.Value, aTableRows);
            }
            else
            {
                colIdx++;
            }
        }
    }

    private int ProcessColumnsWithSameGridSpan(int colIdx, int topColumnCellSpan, List<A.TableRow> aTableRows)
    {
        var deleteColumnCount = topColumnCellSpan - 1;

        // Delete a:gridCol elements and append width of deleting column to merged column
        DeleteAndUpdateGridColumns(colIdx, deleteColumnCount);

        // Delete a:tc elements
        DeleteTableCells(colIdx, deleteColumnCount, aTableRows);

        return topColumnCellSpan;
    }

    private void DeleteAndUpdateGridColumns(int colIdx, int deleteColumnCount)
    {
        for (var i = 0; i < deleteColumnCount; i++)
        {
            var column = (Column)Columns[colIdx + 1 + i];
            column.AGridColumn.Remove();
            Columns[colIdx].Width += column.Width;
        }
    }
}