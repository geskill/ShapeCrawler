#if NETSTANDARD2_0
using System.Collections.Generic;
using ShapeCrawler.Extensions;
#endif
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Presentations;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <inheritdoc />
public sealed class Presentation : IPresentation
{
    private readonly string? inputPresFile;
    private readonly Stream? inputPresStream;
    internal readonly PresentationDocument PresDocument;
    private readonly MemoryStream presStream = new();
    private readonly SlideSize slideSize;

    /// <summary>
    ///     Opens presentation from the specified stream.
    /// </summary>
    public Presentation(Stream stream)
    {
        inputPresStream = stream;
        inputPresStream.Position = 0;
        inputPresStream.CopyTo(presStream);

        PresDocument = PresentationDocument.Open(presStream, true);
        slideSize = new SlideSize(PresDocument.PresentationPart!.Presentation!.SlideSize!);
        MasterSlides = new MasterSlideCollection(PresDocument.PresentationPart!.SlideMasterParts);
        Sections = new SectionCollection(PresDocument);
        Slides = new UpdatedSlideCollection(
            new UserSlideCollection(PresDocument.PresentationPart.SlideParts),
            PresDocument.PresentationPart);
        Footer = new Footer(new UpdatedSlideCollection(
            new UserSlideCollection(PresDocument.PresentationPart.SlideParts), PresDocument.PresentationPart));
        Properties =
            PresDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(PresDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
    }

    /// <summary>
    ///     Opens presentation from the specified file.
    /// </summary>
    public Presentation(string file)
    {
        inputPresFile = file;
        using var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read);
        fileStream.CopyTo(presStream);

        PresDocument = PresentationDocument.Open(presStream, true);
        slideSize = new SlideSize(PresDocument.PresentationPart!.Presentation!.SlideSize!);
        MasterSlides = new MasterSlideCollection(PresDocument.PresentationPart!.SlideMasterParts);
        Sections = new SectionCollection(PresDocument);
        Slides = new UpdatedSlideCollection(
            new UserSlideCollection(PresDocument.PresentationPart.SlideParts),
            PresDocument.PresentationPart);
        Footer = new Footer(new UpdatedSlideCollection(
            new UserSlideCollection(PresDocument.PresentationPart.SlideParts), PresDocument.PresentationPart));
        Properties =
            PresDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(PresDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public Presentation()
    {
        presStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("new presentation.pptx");

        PresDocument = PresentationDocument.Open(presStream, true);
        slideSize = new SlideSize(PresDocument.PresentationPart!.Presentation!.SlideSize!);
        MasterSlides = new MasterSlideCollection(PresDocument.PresentationPart!.SlideMasterParts);
        Sections = new SectionCollection(PresDocument);
        Slides = new UpdatedSlideCollection(
            new UserSlideCollection(PresDocument.PresentationPart.SlideParts),
            PresDocument.PresentationPart);
        Footer = new Footer(new UpdatedSlideCollection(
            new UserSlideCollection(PresDocument.PresentationPart.SlideParts), PresDocument.PresentationPart));
        Properties =
            PresDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(PresDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
        Properties.Modified = SCSettings.TimeProvider.UtcNow;
    }

    /// <summary>
    ///     Creates a new presentation using fluent configuration.
    /// </summary>
    public Presentation(Action<DraftPresentation> configure)
        : this()
    {
        var draft = new DraftPresentation(this);
        configure(draft);
        draft.ApplyTo(this);
    }

    /// <inheritdoc />
    public ISlideCollection Slides { get; }

    /// <inheritdoc />
    public decimal SlideHeight
    {
        get => slideSize.Height;
        set => slideSize.Height = value;
    }

    /// <inheritdoc />
    public decimal SlideWidth
    {
        get => slideSize.Width;
        set => slideSize.Width = value;
    }

    /// <inheritdoc />
    public IMasterSlideCollection MasterSlides { get; }

    /// <inheritdoc />
    public ISectionCollection Sections { get; }

    /// <inheritdoc />
    public IFooter Footer { get; }

    /// <inheritdoc />
    public IPresentationProperties Properties { get; }

    /// <inheritdoc />
    public IUserSlide Slide(int number)
    {
        if (number < 0)
        {
            throw new SCException($"Specified slide number is must {number} be more than zero.");
        }

        return number > Slides.Count
            ? throw new SCException(
                $"Specified slide number {number} exceeds the number of slides {Slides.Count} in the presentation.")
            : Slides[number - 1];
    }

    /// <inheritdoc />
    public void Save()
    {
        // Materialize initial template slide if SlideIdList is empty but slide parts exist
        EnsureInitialSlideId();
        PresDocument.PresentationPart!.Presentation!.Save();
        PresDocument.Save();
        if (inputPresStream is not null)
        {
            PresDocument.Clone(inputPresStream);
        }
        else if (inputPresFile is not null)
        {
            var savedPres = PresDocument.Clone(inputPresFile);
            savedPres.Dispose();
        }
    }

    /// <inheritdoc />
    public void Save(Stream stream)
    {
        Properties.Modified = SCSettings.TimeProvider.UtcNow;
        EnsureInitialSlideId();
        PresDocument.PresentationPart!.Presentation!.Save();

        if (stream is FileStream fileStream)
        {
            var mStream = new MemoryStream();
            PresDocument.Clone(mStream);
            mStream.Position = 0;
            mStream.CopyTo(fileStream);
        }
        else
        {
            PresDocument.Clone(stream);
        }
    }

    /// <inheritdoc />
    public void Save(string file)
    {
        Save();
        using var stream = new FileStream(file, FileMode.Create);
        Save(stream);
    }

    /// <inheritdoc />
    public string AsMarkdown()
    {
        var markdown = new StringBuilder();
        foreach (var slide in Slides)
        {
            markdown.AppendLine($"# Slide {slide.Number}");
            var textShapes = slide.Shapes
                .Select(shape => new { shape, shapeText = shape.TextBox })
                .Where(x => x.shapeText is not null
                            && x.shapeText.Text != string.Empty
                            && x.shape.PlaceholderType != PlaceholderType.SlideNumber);

            var titleShape = textShapes.FirstOrDefault(x =>
                x.shape.Name.StartsWith("Title", StringComparison.OrdinalIgnoreCase));
            if (titleShape != null)
            {
                markdown.AppendLine($"## {titleShape.shapeText!.Text}");
            }

            foreach (var nonTitleShape in textShapes
                         .Where(x => !x.shape.Name.StartsWith("Title", StringComparison.OrdinalIgnoreCase)))
            {
                markdown.AppendLine(nonTitleShape.shapeText!.Text);
            }

            markdown.AppendLine();
        }

        return markdown.ToString();
    }

    /// <inheritdoc />
    public string AsBase64()
    {
        using var stream = new MemoryStream();
        Save(stream);

        return Convert.ToBase64String(stream.ToArray());
    }

    /// <inheritdoc />
    public PresentationDocument GetSdkPresentationDocument()
    {
        return PresDocument.Clone();
    }

    /// <summary>
    ///     Releases all resources used by the presentation.
    /// </summary>
    public void Dispose()
    {
        PresDocument.Dispose();
    }

    /// <summary>
    ///     Starts a fluent creation of a new presentation.
    /// </summary>
    public static DraftPresentation Create(Action<DraftPresentation> configure)
    {
        var draft = new DraftPresentation();
        configure(draft);
        return draft;
    }

    /// <summary>
    ///     Gets Slide Master by number.
    /// </summary>
    public IMasterSlide SlideMaster(int number)
    {
        return MasterSlides[number - 1];
    }

    private void EnsureInitialSlideId()
    {
        var presentationPart = PresDocument.PresentationPart!;
        var presentation = presentationPart.Presentation!;
        presentation.SlideIdList ??= new P.SlideIdList();
#if NETSTANDARD2_0
        var existingIds = new HashSet<string>(
            presentation.SlideIdList
                .OfType<P.SlideId>()
                .Select(s => (string)s.RelationshipId!));
#else
        var existingIds = presentation.SlideIdList
            .OfType<P.SlideId>()
            .Select(s => (string)s.RelationshipId!)
            .ToHashSet();
#endif
        var nextIdVal = presentation.SlideIdList.OfType<P.SlideId>().Any()
            ? presentation.SlideIdList.OfType<P.SlideId>().Max(s => s.Id!.Value) + 1u
            : 256u;

        // Ensure all slide parts are represented in SlideIdList
        foreach (var slidePart in presentationPart.SlideParts)
        {
            var relId = presentationPart.GetIdOfPart(slidePart);
            if (!existingIds.Contains(relId))
            {
                presentation.SlideIdList.Append(new P.SlideId { Id = nextIdVal++, RelationshipId = relId });
            }
        }
    }
}