using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class ShapeFill(OpenXmlCompositeElement openXmlCompositeElement) : IShapeFill
{
    private A.BlipFill? aBlipFill;
    private A.GradientFill? aGradFill;
    private A.SolidFill? aSolidFill;
    private SlidePictureImage? pictureImage;

    public string? Color
    {
        get
        {
            aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return aRgbColorModelHex.Val!.ToString();
                }

                return ColorHexOrNullOf(aSolidFill.SchemeColor!.Val!);
            }

            return null;
        }
    }

    public double Alpha
    {
        get
        {
            const int defaultAlphaPercentages = 100;
            aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    var alpha = aRgbColorModelHex.Elements<A.Alpha>().FirstOrDefault();
                    return alpha?.Val?.Value / 1000d ?? defaultAlphaPercentages;
                }

                var schemeColor = aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.Alpha>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? defaultAlphaPercentages;
            }

            return defaultAlphaPercentages;
        }
    }

    public double LuminanceModulation
    {
        get
        {
            const double luminanceModulation = 100;
            aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return luminanceModulation;
                }

                var schemeColor = aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.LuminanceModulation>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? luminanceModulation;
            }

            return luminanceModulation;
        }
    }

    public double LuminanceOffset
    {
        get
        {
            const double defaultValue = 0;
            aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return defaultValue;
                }

                var schemeColor = aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.LuminanceOffset>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? defaultValue;
            }

            return defaultValue;
        }
    }

    public IImage? Picture => GetPictureImage();

    public FillType Type => GetFillType();

    public void SetPicture(Stream image)
    {
        var openXmlPart = openXmlCompositeElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (Type == FillType.Picture)
        {
            pictureImage!.Update(image);
        }
        else
        {
            openXmlCompositeElement.GetFirstChild<A.SolidFill>()?.Remove();
            openXmlCompositeElement.GetFirstChild<A.GradientFill>()?.Remove();
            openXmlCompositeElement.GetFirstChild<A.PatternFill>()?.Remove();
            openXmlCompositeElement.GetFirstChild<A.NoFill>()?.Remove();

            var rId = openXmlPart.AddImagePart(image, "image/png");

            aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            var aOutline = openXmlCompositeElement.GetFirstChild<A.Outline>();
            if (aOutline != null)
            {
                openXmlCompositeElement.InsertBefore(aBlipFill, aOutline);
            }
            else
            {
                openXmlCompositeElement.Append(aBlipFill);
            }

            aSolidFill = null;
            aGradFill = null;
            pictureImage = new SlidePictureImage(aBlipFill.Blip!);
        }
    }

    public void SetColor(string hex)
    {
        InitSolidFillOr();
        openXmlCompositeElement.AddSolidFill(hex);
    }

    public void SetNoFill()
    {
        InitSolidFillOr();
        openXmlCompositeElement.AddNoFill();
    }

    private static A.ColorScheme GetColorScheme(OpenXmlPart openXmlPart)
    {
        return openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)openXmlPart).ThemePart!.Theme!.ThemeElements!.ColorScheme!
        };
    }

    private void InitSolidFillOr()
    {
        aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
        if (aSolidFill == null)
        {
            aGradFill = openXmlCompositeElement!.GetFirstChild<A.GradientFill>();
            if (aGradFill == null)
            {
                InitPictureFillOr();
            }
        }
    }

    private bool HasSolidFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.SolidFill>() != null;
    }

    private bool HasGradientFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.GradientFill>() != null;
    }

    private bool HasBlipFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.BlipFill>() != null;
    }

    private bool HasPatternFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.PatternFill>() != null;
    }

    private FillType GetFillType()
    {
        if (HasSolidFill())
        {
            return FillType.Solid;
        }

        if (HasGradientFill())
        {
            return FillType.Gradient;
        }

        if (HasBlipFill())
        {
            return FillType.Picture;
        }

        if (HasPatternFill())
        {
            return FillType.Pattern;
        }

        if (openXmlCompositeElement.Ancestors<P.Shape>().FirstOrDefault()?.UseBackgroundFill is not null)
        {
            return FillType.SlideBackground;
        }

        return FillType.NoFill;
    }

    private string? ColorHexOrNullOf(string schemeColor)
    {
        var openXmlPart = openXmlCompositeElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aColorScheme = GetColorScheme(openXmlPart);

        var aColor2Type = aColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == schemeColor);
        return aColor2Type?.RgbColorModelHex?.Val?.Value
               ?? aColor2Type?.SystemColor?.LastColor?.Value;
    }

    private void InitPictureFillOr()
    {
        aBlipFill = openXmlCompositeElement.GetFirstChild<A.BlipFill>();

        if (aBlipFill is not null)
        {
            var image = new SlidePictureImage(aBlipFill.Blip!);
            pictureImage = image;
        }
    }

    private SlidePictureImage? GetPictureImage()
    {
        InitSolidFillOr();

        return pictureImage;
    }
}