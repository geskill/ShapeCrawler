using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130
using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a theme.
/// </summary>
public interface ITheme
{
    /// <summary>
    ///     Gets font scheme.
    /// </summary>
    IThemeFontScheme FontScheme { get; }

    /// <summary>
    ///     Gets the color scheme.
    /// </summary>
    IThemeColorScheme ColorScheme { get; }
}

internal sealed class Theme : ITheme
{
    private readonly A.Theme aTheme;
    private readonly OpenXmlPart sdkTypedOpenXmlPart;

    internal Theme(OpenXmlPart sdkTypedOpenXmlPart, A.Theme aTheme)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aTheme = aTheme;
    }

    public IThemeFontScheme FontScheme => new ThemeFontScheme(sdkTypedOpenXmlPart);

    public IThemeColorScheme ColorScheme => GetColorScheme();

    private IThemeColorScheme GetColorScheme()
    {
        return new ThemeColorScheme(aTheme.ThemeElements!.ColorScheme!);
    }
}