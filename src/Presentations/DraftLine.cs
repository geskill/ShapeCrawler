using System;
using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft line.
/// </summary>
public sealed class DraftLine
{
    internal string DraftName { get; private set; } = "Line";

    internal int DraftX { get; private set; }

    internal int DraftY { get; private set; }

    internal int DraftWidth { get; private set; } = 100;

    internal int DraftHeight { get; private set; }

    internal DraftStroke? DraftStroke { get; private set; }

    internal LineEndValues? DraftHeadEndType { get; private set; }

    internal LineEndValues? DraftTailEndType { get; private set; }

    /// <summary>
    ///     Sets name.
    /// </summary>
    public DraftLine Name(string name)
    {
        DraftName = name;
        return this;
    }

    /// <summary>
    ///     Sets X-position of the start point in points.
    /// </summary>
    public DraftLine X(int x)
    {
        DraftX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position of the start point in points.
    /// </summary>
    public DraftLine Y(int y)
    {
        DraftY = y;
        return this;
    }

    /// <summary>
    ///     Sets width in points (endX = startX + width).
    /// </summary>
    public DraftLine Width(int width)
    {
        DraftWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height in points (endY = startY + height).
    /// </summary>
    public DraftLine Height(int height)
    {
        DraftHeight = height;
        return this;
    }

    /// <summary>
    ///     Configures the line stroke.
    /// </summary>
    public DraftLine Line(Action<DraftStroke> configure)
    {
        DraftStroke = new DraftStroke();
        configure(DraftStroke);
        return this;
    }

    /// <summary>
    ///     Sets the arrow head type for the end of the line.
    /// </summary>
    public DraftLine EndArrow(LineEndValues type)
    {
        DraftTailEndType = type;
        return this;
    }

    /// <summary>
    ///     Sets the arrow head type for the start of the line.
    /// </summary>
    public DraftLine StartArrow(LineEndValues type)
    {
        DraftHeadEndType = type;
        return this;
    }
}