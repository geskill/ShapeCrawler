using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal class TableCellFill : IShapeFill
{
    private readonly A.TableCellProperties aTableCellProperties;
    private A.BlipFill? aBlipFill;
    private FillType fillType;
    private string? hexSolidColor;
    private bool isDirty;
    private ShapeFillImage? pictureImage;
    private A.GradientFill? sdkAGradFill;
    private A.PatternFill? sdkAPattFill;
    private A.SolidFill? sdkASolidFill;

    internal TableCellFill(A.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;
        isDirty = true;
    }

    public string? Color => GetHexSolidColor();

    public double Alpha { get; }

    public double LuminanceModulation { get; }

    public double LuminanceOffset { get; }

    public IImage? Picture => GetPicture();

    public FillType Type => GetFillType();

    public void SetPicture(Stream image)
    {
        var openXmlPart = aTableCellProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (isDirty)
        {
            Initialize();
        }

        if (Type == FillType.Picture)
        {
            pictureImage!.Update(image);
        }
        else
        {
            var rId = openXmlPart.AddImagePart(image, "image/png");

            // This could be refactored to DRY vs SlideShapes.CreatePPicture.
            // In the process, the image could be de-duped also.
            aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            aTableCellProperties.Append(aBlipFill);

            sdkASolidFill?.Remove();
            aBlipFill = null;
            sdkAGradFill?.Remove();
            sdkAGradFill = null;
            sdkAPattFill?.Remove();
            sdkAPattFill = null;
        }

        isDirty = true;
    }

    public void SetColor(string hex)
    {
        if (isDirty)
        {
            Initialize();
        }

        aTableCellProperties.AddSolidFill(hex);

        isDirty = true;
    }


    public void SetNoFill()
    {
        if (isDirty)
        {
            Initialize();
        }

        aTableCellProperties.AddNoFill();

        isDirty = true;
    }

    private void InitSlideBackgroundFillOr()
    {
        fillType = FillType.NoFill;
    }

    private FillType GetFillType()
    {
        if (isDirty)
        {
            Initialize();
        }

        return fillType;
    }

    private void Initialize()
    {
        InitSolidFillOr();
        isDirty = false;
    }

    private void InitSolidFillOr()
    {
        sdkASolidFill = aTableCellProperties.GetFirstChild<A.SolidFill>();
        if (sdkASolidFill != null)
        {
            var aRgbColorModelHex = sdkASolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                var hexColor = aRgbColorModelHex.Val!.ToString();
                hexSolidColor = hexColor;
            }

            fillType = FillType.Solid;
        }
        else
        {
            InitGradientFillOr();
        }
    }

    private void InitGradientFillOr()
    {
        sdkAGradFill = aTableCellProperties!.GetFirstChild<A.GradientFill>();
        if (sdkAGradFill != null)
        {
            fillType = FillType.Gradient;
        }
        else
        {
            InitPictureFillOr();
        }
    }

    private void InitPictureFillOr()
    {
        var openXmlPart = aTableCellProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        aBlipFill = aTableCellProperties.GetFirstChild<A.BlipFill>();

        if (aBlipFill is not null)
        {
            var blipEmbedValue = aBlipFill.Blip?.Embed?.Value;
            if (blipEmbedValue != null)
            {
                var imagePart = (ImagePart)openXmlPart.GetPartById(blipEmbedValue);
                var image = new ShapeFillImage(aBlipFill.Blip!, imagePart);
                pictureImage = image;
                fillType = FillType.Picture;
            }
        }
        else
        {
            InitPatternFillOr();
        }
    }

    private void InitPatternFillOr()
    {
        sdkAPattFill = aTableCellProperties.GetFirstChild<A.PatternFill>();
        if (sdkAPattFill != null)
        {
            fillType = FillType.Pattern;
        }
        else
        {
            InitSlideBackgroundFillOr();
        }
    }

    private string? GetHexSolidColor()
    {
        if (isDirty)
        {
            Initialize();
        }

        return hexSolidColor;
    }

    private ShapeFillImage? GetPicture()
    {
        if (isDirty)
        {
            Initialize();
        }

        return pictureImage;
    }
}