using System.IO;
using System.Linq;
using ImageMagick;
using ImageMagick.Formats;
using ShapeCrawler.Drawing;

namespace ShapeCrawler.Slides;

/// <summary>
///     Represents processed image content ready for insertion into a PowerPoint slide.
/// </summary>
internal sealed class Image
{
    private readonly MagickFormat format;
    private readonly MagickImage image;
    private readonly Stream stream;

    internal Image(Stream stream)
    {
        this.stream = stream;

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        image = CreateMagickImage(stream);
        format = image.Format;

        EnsureSupportedImageFormat(image);
        HandleSvgFormat(image, format);

        var width = image.Width;
        var height = image.Height;

        if (format == MagickFormat.Svg)
        {
            ResizeSvgImageIfNeeded(image, ref width, ref height);
        }

        Width = width;
        Height = height;
    }

    internal uint Width { get; }

    internal uint Height { get; }

    internal bool IsSvg => format == MagickFormat.Svg;

    internal bool IsOriginalFormatPreserved =>
        format is MagickFormat.Gif or MagickFormat.Jpeg or MagickFormat.Png ||
        format is MagickFormat.Tif or MagickFormat.Tiff;

    internal string MimeType => GetMimeType(IsOriginalFormatPreserved ? format : image.Format);

    internal string Hash
    {
        get
        {
            if (IsOriginalFormatPreserved)
            {
                stream.Position = 0;
                return new ImageStream(stream).Base64Hash;
            }

            using var rasterStream = GetRasterStream();
            return new ImageStream(rasterStream).Base64Hash;
        }
    }

    internal string SvgHash
    {
        get
        {
            stream.Position = 0;
            return new ImageStream(stream).Base64Hash;
        }
    }

    /// <summary>
    ///     Gets the raster stream for the image.
    /// </summary>
    internal MemoryStream GetRasterStream()
    {
        var rasterStream = new MemoryStream();
        image.Settings.SetDefines(new PngWriteDefines { ExcludeChunks = PngChunkFlags.date });
        image.Settings.SetDefine("png:exclude-chunk", "tIME");
        image.Write(rasterStream);
        rasterStream.Position = 0;
        return rasterStream;
    }

    /// <summary>
    ///     Gets the original stream for formats that are preserved as-is.
    /// </summary>
    internal Stream GetOriginalStream()
    {
        stream.Position = 0;
        return stream;
    }

    private static MagickImage CreateMagickImage(Stream imageStream)
    {
        var format = IsIco(imageStream)
            ? MagickFormat.Ico
            : MagickFormat.Unknown;

        return new MagickImage(
            imageStream,
            new MagickReadSettings { Format = format, BackgroundColor = MagickColors.Transparent });
    }

    private static bool IsIco(Stream stream)
    {
        if (stream.Length < 6)
        {
            return false;
        }

        var originalPosition = stream.Position;
        stream.Seek(0, SeekOrigin.Begin);

        try
        {
            var header = new byte[6];
            var bytesRead = stream.Read(header, 0, 6);

            if (bytesRead < 6)
            {
                return false;
            }

            // ICO file signature:
            // Bytes 0-1: Reserved (must be 0x00 0x00)
            // Bytes 2-3: Image type (must be 0x01 0x00 for ICO)
            // Bytes 4-5: Number of images (must be > 0)
            // https://docs.fileformat.com/image/ico/#header
            return header[0] == 0x00 &&
                   header[1] == 0x00 &&
                   header[2] == 0x01 &&
                   header[3] == 0x00 &&
                   (header[4] > 0 || header[5] > 0);
        }
        finally
        {
            stream.Seek(originalPosition, SeekOrigin.Begin);
        }
    }

    private static void EnsureSupportedImageFormat(MagickImage image)
    {
        MagickFormat[] supportedImageFormats =
        [
            MagickFormat.Jpeg,
            MagickFormat.Png,
            MagickFormat.Gif,
            MagickFormat.Tif,
            MagickFormat.Tiff,
            MagickFormat.Svg
        ];

        if (!supportedImageFormats.Contains(image.Format))
        {
            image.Format = image.HasAlpha ? MagickFormat.Png : MagickFormat.Jpeg;
        }
    }

    private static void HandleSvgFormat(MagickImage image, MagickFormat originalFormat)
    {
        if (originalFormat == MagickFormat.Svg)
        {
            image.Format = MagickFormat.Png;
            image.Density =
                new Density(384,
                    DensityUnit
                        .PixelsPerInch); // in PowerPoint, the resolution of the rasterized version of SVG is set to 384 PPI
        }
    }

    private static void ResizeSvgImageIfNeeded(MagickImage image, ref uint width, ref uint height)
    {
        if (height > 500 || width > 500)
        {
            if (height > 500)
            {
                height = 500;
            }

            if (width > 500)
            {
                width = 500;
            }

            if (height == 500)
            {
                width = (uint)(height * (decimal)image.Width / image.Height);
            }

            if (width == 500)
            {
                height = (uint)(width * (decimal)image.Height / image.Width);
            }

            image.Resize(width, height);
        }
    }

    private static string GetMimeType(MagickFormat format)
    {
        var mime = MagickFormatInfo.Create(format)?.MimeType;

        return mime ?? throw new SCException("Unsupported image format.");
    }
}