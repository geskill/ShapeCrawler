namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft cell builder.
/// </summary>
public sealed class DraftCell
{
    internal string? SolidColorHex { get; private set; }

    internal string? FontColorHex { get; private set; }

    internal string? TextContent { get; private set; }

    /// <summary>
    ///     Sets the solid color fill for the cell.
    /// </summary>
    public DraftCell FillSolidColor(string hex)
    {
        SolidColorHex = hex;
        return this;
    }

    /// <summary>
    ///     Sets the font color for the cell.
    /// </summary>
    public DraftCell FontColor(string hex)
    {
        FontColorHex = hex;
        return this;
    }

    /// <summary>
    ///     Sets the text content for the cell.
    /// </summary>
    public DraftCell TextBox(string content)
    {
        TextContent = content;
        return this;
    }
}