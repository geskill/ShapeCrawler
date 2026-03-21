using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Colors;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Paragraphs;

internal sealed class TextParagraphPortion : IParagraphPortion
{
    private readonly A.Run aRun;
    private readonly Lazy<TextPortionFont> font;
    private readonly Lazy<Hyperlink> hyperlink;

    internal TextParagraphPortion(A.Run aRun)
    {
        AText = aRun.Text!;
        this.aRun = aRun;
        var openXmlPart = AText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        font = new Lazy<TextPortionFont>(() =>
            new TextPortionFont(
                new PortionFontSize(AText),
                new Lazy<FontColor>(() => new FontColor(AText)),
                new ThemeFontScheme(openXmlPart),
                AText
            )
        );
        hyperlink = new Lazy<Hyperlink>(() => new Hyperlink(this.aRun.RunProperties!));
    }

    internal A.Text AText { get; }

    public string Text
    {
        get => AText.Text;
        set => AText.Text = value;
    }

    public ITextPortionFont Font => font.Value;

    public IHyperlink Link => hyperlink.Value;

    public Color TextHighlightColor
    {
        get => GetTextHighlight();
        set => SetTextHighlight(value);
    }

    public void Remove()
    {
        aRun.Remove();
    }

    private Color GetTextHighlight()
    {
        var arPr = AText.PreviousSibling<A.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex is not A.RgbColorModelHex aSrgbClr
            || aSrgbClr.Val is null)
        {
            return Color.NoColor;
        }

        var hex = aSrgbClr.Val.ToString()!;

        var color = new Color(hex);

        var aAlphaValue = aSrgbClr.GetFirstChild<A.Alpha>()?.Val ?? 100000;
        color.Alpha = Color.Opacity * aAlphaValue / 100_000f;

        return color;
    }

    private void SetTextHighlight(Color color)
    {
        var arPr = AText.PreviousSibling<A.RunProperties>() ?? AText.Parent!.AddRunProperties();
        arPr.AddAHighlight(color);
    }
}