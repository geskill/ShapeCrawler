using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Colors;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class Field : IParagraphPortion
{
    private readonly A.Field aField;
    private readonly A.Text? aText;
    private readonly FieldPortionText fieldPortionText;
    private readonly Lazy<ITextPortionFont> font;
    private readonly Lazy<Hyperlink> hyperlink;

    internal Field(A.Field aField)
    {
        aText = aField.GetFirstChild<A.Text>()!;
        this.aField = aField;
        font = new Lazy<ITextPortionFont>(() =>
        {
            var textPortionSize = new PortionFontSize(aText!);
            var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
            return
                new TextPortionFont(
                    textPortionSize,
                    new Lazy<FontColor>(() => new FontColor(aText)),
                    new ThemeFontScheme(openXmlPart),
                    aText!
                );
        });
        fieldPortionText = new FieldPortionText(this.aField);
        hyperlink = new Lazy<Hyperlink>(() => new Hyperlink(this.aField.RunProperties!));
    }

    /// <inheritdoc />
    public string Text
    {
        get => fieldPortionText.Value;
        set => fieldPortionText.Update(value);
    }

    /// <inheritdoc />
    public ITextPortionFont Font => font.Value;

    public IHyperlink Link => hyperlink.Value;

    public Color TextHighlightColor
    {
        get => ParseTextHighlight();
        set => UpdateTextHighlight(value);
    }

    public void Remove()
    {
        aField.Remove();
    }

    private Color ParseTextHighlight()
    {
        var arPr = aText!.PreviousSibling<A.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex is not A.RgbColorModelHex aSrgbClr
            || aSrgbClr.Val is null)
        {
            return Color.NoColor;
        }

        // Gets node value.
        var hex = aSrgbClr.Val.ToString()!;

        // Check if the color value is valid, we are expecting values as "000000".
        var color = new Color(hex);

        // Calculate the alpha value if it is defined in the highlight node.
        var aAlphaValue = aSrgbClr.GetFirstChild<A.Alpha>()?.Val ?? 100000;
        color.Alpha = Color.Opacity * aAlphaValue / 100_000f;

        return color;
    }

    private void UpdateTextHighlight(Color color)
    {
        var arPr = aText!.PreviousSibling<A.RunProperties>() ?? aText.Parent!.AddRunProperties();

        arPr.AddAHighlight(color);
    }
}