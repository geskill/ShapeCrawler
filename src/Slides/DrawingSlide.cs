using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Slides;

/// <inheritdoc />
internal sealed class DrawingSlide(ILayoutSlide layoutSlide, UserSlideShapeCollection shapes, SlidePart slidePart)
    : UserSlide(layoutSlide, shapes, slidePart)
{
    /// <inheritdoc />
    public override void SaveImageTo(string file)
    {
        using var fileStream = File.Create(file);
        SaveImageTo(fileStream);
    }

    /// <inheritdoc />
    public override void SaveImageTo(Stream stream)
    {
        var presPart = GetSdkPresentationPart();
        var pSlideSize = presPart.Presentation!.SlideSize!;
        var width = new Emus(pSlideSize.Cx!.Value).AsPixels();
        var height = new Emus(pSlideSize.Cy!.Value).AsPixels();

        using var surface = SKSurface.Create(new SKImageInfo((int)width, (int)height));
        var canvas = surface.Canvas;

        RenderBackground(canvas);
        shapes.Render(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        data.SaveTo(stream);

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }
    }

    private SKColor GetSkColor()
    {
        var hex = Fill.Color!.TrimStart('#');

        // Validate hex length before parsing.
        if (hex.Length != 6 && hex.Length != 8)
        {
            return SKColors.White; // used by the PowerPoint application as the default background color
        }

        return new Color(hex).AsSkColor();
    }

    private void RenderBackground(SKCanvas canvas)
    {
        var slideFill = Fill;
        switch (slideFill)
        {
            case { Type: FillType.Solid, Color: not null }:
                {
                    var skColor = GetSkColor();
                    canvas.Clear(skColor);
                    break;
                }

            case { Type: FillType.Picture, Picture: not null }:
                {
                    var bytes = slideFill.Picture.AsByteArray();
                    using var stream = new MemoryStream(bytes);
                    using var bitmap = SKBitmap.Decode(stream);
                    var destRect = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
                    canvas.DrawBitmap(bitmap, destRect);
                    break;
                }

            default:
                // Default to white for unsupported backgrounds.
                canvas.Clear(SKColors.White);
                break;
        }
    }
}