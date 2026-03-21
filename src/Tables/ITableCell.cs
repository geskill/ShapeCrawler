using ShapeCrawler.Drawing;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a table cell.
/// </summary>
public interface ITableCell
{
    /// <summary>
    ///     Gets the shape text.
    /// </summary>
    ITextBox TextBox { get; }

    /// <summary>
    ///     Gets a value indicating whether cell belongs to merged cell.
    /// </summary>
    bool IsMergedCell { get; }

    /// <summary>
    ///     Gets Shape Fill of the cell.
    /// </summary>
    IShapeFill Fill { get; }

    /// <summary>
    ///     Gets the Top Border.
    /// </summary>
    IBorder TopBorder { get; }

    /// <summary>
    ///     Gets the Bottom Border.
    /// </summary>
    IBorder BottomBorder { get; }

    /// <summary>
    ///     Gets the Left Border.
    /// </summary>
    IBorder LeftBorder { get; }

    /// <summary>
    ///     Gets the Right Border.
    /// </summary>
    IBorder RightBorder { get; }
}

internal sealed class TableCell : ITableCell
{
    internal TableCell(A.TableCell aTableCell, int rowIndex, int columnIndex)
    {
        ATableCell = aTableCell;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        TextBox = new TableCellTextBox(ATableCell);
        var aTcPr = aTableCell.TableCellProperties!;
        Fill = new TableCellFill(aTcPr);
        TopBorder = new TopBorder(aTableCell.TableCellProperties!);
        BottomBorder = new BottomBorder(aTableCell.TableCellProperties!);
        LeftBorder = new LeftBorder(aTableCell.TableCellProperties!);
        RightBorder = new RightBorder(aTableCell.TableCellProperties!);
    }

    internal A.TableCell ATableCell { get; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }

    public bool IsMergedCell => ATableCell.GridSpan is not null ||
                                ATableCell.RowSpan is not null ||
                                ATableCell.HorizontalMerge is not null ||
                                ATableCell.VerticalMerge is not null;

    public IShapeFill Fill { get; }

    public IBorder TopBorder { get; }

    public IBorder BottomBorder { get; }

    public IBorder LeftBorder { get; }

    public IBorder RightBorder { get; }

    public ITextBox TextBox { get; }
}