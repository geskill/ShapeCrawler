using System.IO;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a shape fill.
/// </summary>
public interface IShapeFill
{
    /// <summary>
    ///     Gets fill type.
    /// </summary>
    FillType Type { get; }

    /// <summary>
    ///     Gets picture image if it is picture fill, otherwise <see langword="null" />.
    /// </summary>
    IImage? Picture { get; }

    /// <summary>
    ///     Gets color in hexadecimal representation if it is filled with solid color, otherwise <see langword="null" />.
    /// </summary>
    string? Color { get; }

    /// <summary>
    ///     Gets the opacity level of fill color in percentages.
    /// </summary>
    double Alpha { get; }

    /// <summary>
    ///     Gets the Luminance Modulation of fill color in percentages.
    /// </summary>
    double LuminanceModulation { get; }

    /// <summary>
    ///     Gets the Luminance Offset of fill color in percentages.
    /// </summary>
    double LuminanceOffset { get; }

    /// <summary>
    ///     Fills the shape with picture.
    /// </summary>
    void SetPicture(Stream image);

    /// <summary>
    ///     Fills the shape with solid color in hexadecimal representation.
    /// </summary>
    void SetColor(string hex);

    /// <summary>
    ///     Removes Fills from the shape.
    /// </summary>
    void SetNoFill();
}