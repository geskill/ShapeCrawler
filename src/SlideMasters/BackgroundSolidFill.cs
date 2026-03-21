using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

internal sealed class BackgroundSolidFill(SlideLayoutPart slideLayoutPart) : ISolidFill
{
    public string Color
    {
        get
        {
            var pCommonSlideData = slideLayoutPart.SlideLayout!.CommonSlideData;
            var pBackground = pCommonSlideData?.GetFirstChild<Background>();
            var pBackgroundProperties = pBackground?.GetFirstChild<BackgroundProperties>();

            var aSolidFill = pBackgroundProperties?.GetFirstChild<SolidFill>();

            var aRgbColorModelHex = aSolidFill?.RgbColorModelHex;

            return aRgbColorModelHex != null ? aRgbColorModelHex.Val!.ToString()! : string.Empty;
        }
    }
}