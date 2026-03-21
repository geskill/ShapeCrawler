using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a settings of theme font.
/// </summary>
public interface IThemeFontScheme
{
    /// <summary>
    ///     Gets or sets font name for head.
    /// </summary>
    string HeadLatinFont { get; set; }

    /// <summary>
    ///     Gets or sets font name for the Latin characters of the body.
    /// </summary>
    string BodyLatinFont { get; set; }

    /// <summary>
    ///     Gets or sets font name for the East Asian characters of the body.
    /// </summary>
    string BodyEastAsianFont { get; set; }

    /// <summary>
    ///     Gets or sets font name for the East Asian characters of the heading.
    /// </summary>
    string HeadEastAsianFont { get; set; }
}

internal sealed class ThemeFontScheme : IThemeFontScheme
{
    private readonly A.FontScheme aFontScheme;

    internal ThemeFontScheme(OpenXmlPart openXmlPart)
    {
        aFontScheme = openXmlPart switch
        {
            SlidePart slidePart => slidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FontScheme!,
            SlideLayoutPart slideLayoutPart => slideLayoutPart.SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FontScheme!,
            NotesSlidePart notesSlidePart => GetFontSchemeFromNotesSlidePart(notesSlidePart)!,
            _ => ((SlideMasterPart)openXmlPart).ThemePart!.Theme!.ThemeElements!.FontScheme!
        };
    }

    public string HeadLatinFont
    {
        get => GetHeadLatinFont();
        set => SetHeadLatinFont(value);
    }

    public string BodyLatinFont
    {
        get => GetBodyLatinFont();
        set => SetBodyLatinFont(value);
    }

    public string BodyEastAsianFont
    {
        get => GetBodyEastAsianFont();
        set => SetBodyEastAsianFont(value);
    }

    public string HeadEastAsianFont
    {
        get => GetHeadEastAsianFont();
        set => SetHeadEastAsianFont(value);
    }

    internal string MajorLatinFont()
    {
        return aFontScheme.MajorFont!.LatinFont!.Typeface!;
    }

    internal string MajorEastAsianFont()
    {
        return aFontScheme.MajorFont!.EastAsianFont!.Typeface!;
    }

    internal string MinorEastAsianFont()
    {
        return aFontScheme.MinorFont!.EastAsianFont!.Typeface!;
    }

    internal A.LatinFont MinorLatinFont()
    {
        return aFontScheme.MinorFont!.LatinFont!;
    }

    internal void UpdateMinorEastAsianFont(string eastAsianFont)
    {
        aFontScheme.MinorFont!.EastAsianFont!.Typeface = eastAsianFont;
    }

    private static A.FontScheme GetFontSchemeFromNotesSlidePart(NotesSlidePart notesSlidePart)
    {
        // If NotesMasterPart exists, use it
        var notesMasterFontScheme = notesSlidePart.NotesMasterPart?.ThemePart?.Theme!.ThemeElements!.FontScheme;
        if (notesMasterFontScheme != null)
        {
            return notesMasterFontScheme;
        }

        // Fall back to the slide's master part if NotesMasterPart is null
        var parentSlidePart = notesSlidePart.GetParentParts().OfType<SlidePart>().FirstOrDefault();
        var slideMasterFontScheme = parentSlidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme!.ThemeElements!
            .FontScheme;
        if (slideMasterFontScheme != null)
        {
            return slideMasterFontScheme;
        }

        throw new SCException("Could not find font scheme for notes slide part");
    }

    private string GetHeadLatinFont()
    {
        return aFontScheme.MajorFont!.LatinFont!.Typeface!.Value!;
    }

    private string GetHeadEastAsianFont()
    {
        return aFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value!;
    }

    private void SetHeadLatinFont(string fontName)
    {
        aFontScheme.MajorFont!.LatinFont!.Typeface!.Value = fontName;
    }

    private void SetHeadEastAsianFont(string fontName)
    {
        aFontScheme.MajorFont!.EastAsianFont!.Typeface!.Value = fontName;
    }

    private string GetBodyLatinFont()
    {
        return aFontScheme.MinorFont!.LatinFont!.Typeface!.Value!;
    }

    private string GetBodyEastAsianFont()
    {
        return aFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value!;
    }

    private void SetBodyLatinFont(string fontName)
    {
        aFontScheme.MinorFont!.LatinFont!.Typeface!.Value = fontName;
    }

    private void SetBodyEastAsianFont(string fontName)
    {
        aFontScheme.MinorFont!.EastAsianFont!.Typeface!.Value = fontName;
    }
}