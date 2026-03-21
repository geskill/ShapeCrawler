using System.Linq;
using ShapeCrawler.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130
using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a color scheme.
/// </summary>
public interface IThemeColorScheme
{
    /// <summary>
    ///     Gets or sets Dark 1 color in hexadecimal format.
    /// </summary>
    string Dark1 { get; set; }

    /// <summary>
    ///     Gets or sets Light 1 color in hexadecimal format.
    /// </summary>
    string Light1 { get; set; }

    /// <summary>
    ///     Gets or sets Dark 2 color in hexadecimal format.
    /// </summary>
    string Dark2 { get; set; }

    /// <summary>
    ///     Gets or sets Light 2 color in hexadecimal format.
    /// </summary>
    string Light2 { get; set; }

    /// <summary>
    ///     Gets or sets Accent 1 color in hexadecimal format.
    /// </summary>
    string Accent1 { get; set; }

    /// <summary>
    ///     Gets or sets Accent 2 color in hexadecimal format.
    /// </summary>
    string Accent2 { get; set; }

    /// <summary>
    ///     Gets or sets Accent 3 color in hexadecimal format.
    /// </summary>
    string Accent3 { get; set; }

    /// <summary>
    ///     Gets or sets Accent 4 color in hexadecimal format.
    /// </summary>
    string Accent4 { get; set; }

    /// <summary>
    ///     Gets or sets Accent 5 color in hexadecimal format.
    /// </summary>
    string Accent5 { get; set; }

    /// <summary>
    ///     Gets or sets Accent 6 color in hexadecimal format.
    /// </summary>
    string Accent6 { get; set; }

    /// <summary>
    ///     Gets or sets Hyperlink color in hexadecimal format.
    /// </summary>
    string Hyperlink { get; set; }

    /// <summary>
    ///     Gets or sets Followed Hyperlink color in hexadecimal format.
    /// </summary>
    string FollowedHyperlink { get; set; }
}

internal sealed class ThemeColorScheme : IThemeColorScheme
{
    private readonly A.ColorScheme aColorScheme;

    internal ThemeColorScheme(A.ColorScheme aColorScheme)
    {
        this.aColorScheme = aColorScheme;
    }

    public string Dark1
    {
        get => GetColor(aColorScheme.Dark1Color!);
        set => SetColor("dk1", value);
    }

    public string Light1
    {
        get => GetColor(aColorScheme.Light1Color!);
        set => SetColor("lt1", value);
    }

    public string Dark2
    {
        get => GetColor(aColorScheme.Dark2Color!);
        set => SetColor("dk2", value);
    }

    public string Light2
    {
        get => GetColor(aColorScheme.Light2Color!);
        set => SetColor("lt2", value);
    }

    public string Accent1
    {
        get => GetColor(aColorScheme.Accent1Color!);
        set => SetColor("accent1", value);
    }

    public string Accent2
    {
        get => GetColor(aColorScheme.Accent2Color!);
        set => SetColor("accent2", value);
    }

    public string Accent3
    {
        get => GetColor(aColorScheme.Accent3Color!);
        set => SetColor("accent3", value);
    }

    public string Accent4
    {
        get => GetColor(aColorScheme.Accent4Color!);
        set => SetColor("accent4", value);
    }

    public string Accent5
    {
        get => GetColor(aColorScheme.Accent5Color!);
        set => SetColor("accent5", value);
    }

    public string Accent6
    {
        get => GetColor(aColorScheme.Accent6Color!);
        set => SetColor("accent6", value);
    }

    public string Hyperlink
    {
        get => GetColor(aColorScheme.Hyperlink!);
        set => SetColor("hlink", value);
    }

    public string FollowedHyperlink
    {
        get => GetColor(aColorScheme.FollowedHyperlinkColor!);
        set => SetColor("folHlink", value);
    }

    private static string GetColor(A.Color2Type aColor2Type)
    {
        var color = HexParser.GetWithoutScheme(aColor2Type);
        return color!.Value.Hex;
    }

    private void SetColor(string name, string hex)
    {
        var color = aColorScheme.Elements().First(x => x.LocalName == name);
        foreach (var child in color)
        {
            child.Remove();
        }

        var aSrgbClr = new A.RgbColorModelHex { Val = hex };
        color.Append(aSrgbClr);
    }
}