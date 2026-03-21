using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a slide number font.
/// </summary>
public interface ISlideNumberFont : IFont
{
    /// <summary>
    ///     Gets or sets color.
    /// </summary>
    Color Color { get; set; }
}

internal sealed class SlideNumberFont : ISlideNumberFont
{
    private readonly A.DefaultRunProperties aDefaultRunProperties;
    private readonly MasterSlideNumberSize masterSlideNumberSize;

    internal SlideNumberFont(A.DefaultRunProperties aDefaultRunProperties)
    {
        this.aDefaultRunProperties = aDefaultRunProperties;
        masterSlideNumberSize = new MasterSlideNumberSize(aDefaultRunProperties);
    }

    public Color Color
    {
        get => ParseColor();
        set => UpdateColor(value);
    }

    public decimal Size
    {
        get => masterSlideNumberSize.Size;
        set => masterSlideNumberSize.Size = value;
    }

    private void UpdateColor(Color color)
    {
        var solidFill = aDefaultRunProperties.GetFirstChild<A.SolidFill>();
        solidFill?.Remove();

        var rgbColorModelHex = new A.RgbColorModelHex { Val = color.ToString() };
        solidFill = new A.SolidFill(rgbColorModelHex);

        aDefaultRunProperties.Append(solidFill);
    }

    private Color ParseColor()
    {
        var hex = aDefaultRunProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value!;

        return new Color(hex);
    }
}