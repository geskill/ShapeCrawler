using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Drawing;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using ApplicationNonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;

namespace ShapeCrawler.Slides;

/// <inheritdoc />
internal abstract class UserSlide(ILayoutSlide layoutSlide, UserSlideShapeCollection shapes, SlidePart slidePart)
    : IUserSlide
{
    public ILayoutSlide LayoutSlide => layoutSlide;

    public IUserSlideShapeCollection Shapes => shapes;

    public int Number
    {
        get
        {
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
            var presPart = presDocument.PresentationPart!;
            var currentSlidePartId = presPart.GetIdOfPart(slidePart);
            var slideIdList =
                presPart.Presentation!.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
            for (var i = 0; i < slideIdList.Count; i++)
            {
                if (slideIdList[i].RelationshipId == currentSlidePartId)
                {
                    return i + 1;
                }
            }

            throw new SCException("An error occurred while parsing slide number.");
        }

        set
        {
            if (Number == value)
            {
                throw new SCException("Slide number is already set to the specified value.");
            }

            var currentIndex = Number - 1;
            var newIndex = value - 1;
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
            if (newIndex < 0 || newIndex >= presDocument.PresentationPart!.SlideParts.Count())
            {
                throw new SCException("Slide number is out of range.");
            }

            var presentationPart = presDocument.PresentationPart!;
            var presentation = presentationPart.Presentation;
            var slideIdList = presentation!.SlideIdList!;

            // Get the slide ID of the source slide.
            var sourceSlide = (SlideId)slideIdList.ChildElements[currentIndex];

            SlideId? targetSlide;

            // Identify the position of the target slide after which to move the source slide
            if (newIndex == 0)
            {
                targetSlide = null;
            }
            else if (currentIndex < newIndex)
            {
                targetSlide = (SlideId)slideIdList.ChildElements[newIndex];
            }
            else
            {
                targetSlide = (SlideId)slideIdList.ChildElements[newIndex - 1];
            }

            // Remove the source slide from its current position.
            sourceSlide.Remove();
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            presentation.Save();
        }
    }

    public string? CustomData
    {
        get => GetCustomData();
        set => SetCustomData(value);
    }

    public ITextBox? Notes => GetNotes();

    public IShapeFill Fill
    {
        get
        {
            if (field is not null)
            {
                return field!;
            }

            var pcSld = slidePart.Slide!.CommonSlideData
                        ?? slidePart.Slide!.AppendChild(
                            new CommonSlideData());

            // Background element needs to be first, else it gets ignored.
            var pBg = pcSld.GetFirstChild<Background>()
                      ?? pcSld.InsertAt<Background>(new Background(), 0);

            var pBgPr = pBg.GetFirstChild<BackgroundProperties>();
            if (pBgPr is null)
            {
                // PowerPoint always keeps background properties schema-valid.
                // If we create an empty p:bgPr, Open XML validation fails because it must contain a fill element.
                pBgPr = new BackgroundProperties(new A.NoFill());
                pBg.AppendChild(pBgPr);
            }
            else
            {
                var hasFill =
                    pBgPr.GetFirstChild<A.BlipFill>() is not null
                    || pBgPr.GetFirstChild<A.GradientFill>() is not null
                    || pBgPr.GetFirstChild<A.NoFill>() is not null;
                hasFill = hasFill
                          || pBgPr.GetFirstChild<A.PatternFill>() is not null
                          || pBgPr.GetFirstChild<A.SolidFill>() is not null;
                if (!hasFill)
                {
                    // Keep schema-valid even if p:bgPr was previously created empty.
                    pBgPr.InsertAt(new A.NoFill(), 0);
                }
            }

            field = new ShapeFill(pBgPr);

            return field!;
        }
    }

    public bool Hidden()
    {
        return slidePart.Slide!.Show is not null && !slidePart.Slide!.Show.Value;
    }

    public void Hide()
    {
        if (slidePart.Slide!.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            slidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            slidePart.Slide.Show = false;
        }
    }

    public IShape Shape(string name)
    {
        return Shapes.Shape<IShape>(name);
    }

    public IShape Shape(int id)
    {
        return Shapes.GetById<IShape>(id);
    }

    public T Shape<T>(string name)
        where T : IShape
    {
        return Shapes.Shape<T>(name);
    }

    /// <inheritdoc />
    public abstract void SaveImageTo(string file);

    /// <inheritdoc />
    public abstract void SaveImageTo(Stream stream);

    public PresentationPart GetSdkPresentationPart()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;

        return presDocument.Clone().PresentationPart!;
    }

    public T First<T>()
    {
        return (T)Shapes.First(shape => shape is T);
    }

    public IList<ITextBox> GetTexts()
    {
        var collectedTextBoxes = new List<ITextBox>();

        foreach (var shape in Shapes)
        {
            CollectTextBoxes(shape, collectedTextBoxes);
        }

        return collectedTextBoxes;
    }

    /// <inheritdoc />
    public void AddNotes(IEnumerable<string> lines)
    {
        var notes = Notes;
        if (notes is null)
        {
            AddNotesSlide(lines);
        }
        else
        {
            var paragraphs = notes.Paragraphs;
            foreach (var line in lines)
            {
                paragraphs.Add();
                paragraphs[paragraphs.Count - 1].Text = line;
            }
        }
    }

    public void Remove()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presDocument.PresentationPart!.Presentation!;
        var slideIdList = pPresentation.SlideIdList!;

        // Find the exact SlideId corresponding to this slide
        var slideIdRelationship = presPart.GetIdOfPart(slidePart);
        var removingPSlideId = slideIdList.Elements<SlideId>()
                                   .FirstOrDefault(slideId => slideId.RelationshipId!.Value == slideIdRelationship) ??
                               throw new SCException("Could not find slide ID in presentation.");

        var sectionList = pPresentation.PresentationExtensionList?.Descendants<SectionList>().FirstOrDefault();
        var removingSectionSlideIdListEntry = sectionList?.Descendants<SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingPSlideId.Id!);
        removingSectionSlideIdListEntry?.Remove();

        slideIdList.RemoveChild(removingPSlideId);
        pPresentation.Save();

        var removingSlideIdRelationshipId = removingPSlideId.RelationshipId!;
        new SCPPresentation(pPresentation).RemoveSlideIdFromCustomShow(removingSlideIdRelationshipId.Value!);

        var removingSlidePart = (SlidePart)presPart.GetPartById(removingSlideIdRelationshipId!);
        presPart.DeletePart(removingSlidePart);

        presPart.Presentation!.Save();
    }

    /// <summary>
    ///     Gets the underlying <see cref="SlidePart" />.
    /// </summary>
    /// <returns>Slide part instance.</returns>
    internal SlidePart GetSdkSlidePart()
    {
        return slidePart;
    }

    private void CollectTextBoxes(IShape shape, List<ITextBox> buffer)
    {
        if (shape.TextBox is not null)
        {
            buffer.Add(shape.TextBox);
        }

        if (shape.Table is not null)
        {
            foreach (var cell in shape.Table.Rows.SelectMany(row => row.Cells))
            {
                buffer.Add(cell.TextBox);
            }
        }

        if (shape.GroupedShapes is not null)
        {
            foreach (var innerShape in shape.GroupedShapes)
            {
                CollectTextBoxes(innerShape, buffer);
            }
        }
    }

    private ITextBox? GetNotes()
    {
        var notesSlidePart = slidePart.NotesSlidePart;

        if (notesSlidePart is null)
        {
            return null;
        }

        var notesShapes = new ShapeCollection(notesSlidePart);
        var notesPlaceholder = notesShapes
            .FirstOrDefault(shape =>
                shape is { PlaceholderType: not null, TextBox: not null, PlaceholderType: PlaceholderType.Text });
        return notesPlaceholder?.TextBox;
    }

    private void AddNotesSlide(IEnumerable<string> lines)
    {
        // Build up the children of the text body element
        var textBodyChildren = new List<OpenXmlElement> { new A.BodyProperties(), new A.ListStyle() };

        // Add in the text lines
        textBodyChildren.AddRange(
            lines
                .Select(line => new A.Paragraph(
                    new A.ParagraphProperties(),
                    new A.Run(
                        new A.RunProperties(),
                        new A.Text(line)),
                    new A.EndParagraphRunProperties())));

        // Always add at least one paragraph, even if empty
        if (!lines.Any())
        {
            textBodyChildren.Add(
                new A.Paragraph(
                    new A.EndParagraphRunProperties()));
        }

        // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides
        var rid = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var notesSlidePart1 = slidePart.AddNewPart<NotesSlidePart>(rid);
        var notesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = string.Empty },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = (UInt32Value)2U, Name = "Notes Placeholder 2" },
                            new NonVisualShapeDrawingProperties(
                                new A.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body })),
                        new ShapeProperties(),
                        new TextBody(
                            textBodyChildren)))),
            new ColorMapOverride(new A.MasterColorMapping()));
        notesSlidePart1.NotesSlide = notesSlide;
    }

    private string? GetCustomData()
    {
        var getCustomXmlPart = GetCustomXmlPartOrNull();
        if (getCustomXmlPart == null)
        {
            return null;
        }

        var customXmlPartStream = getCustomXmlPart.GetStream();
        using var customXmlStreamReader = new StreamReader(customXmlPartStream);
        var raw = customXmlStreamReader.ReadToEnd();
        return raw[3..];
    }

    private void SetCustomData(string? value)
    {
        var getCustomXmlPart = GetCustomXmlPartOrNull();
        Stream customXmlPartStream;
        if (getCustomXmlPart == null)
        {
            var newSlideCustomXmlPart = slidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPartStream = newSlideCustomXmlPart.GetStream();
        }
        else
        {
            customXmlPartStream = getCustomXmlPart.GetStream();
        }

        using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
        customXmlStreamReader.Write($"ctd{value}");
    }

    private CustomXmlPart? GetCustomXmlPartOrNull()
    {
        foreach (var customXmlPart in slidePart.CustomXmlParts)
        {
            using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
            var customXmlPartText = customXmlPartStream.ReadToEnd();
            if (customXmlPartText.StartsWith(
                    "ctd",
                    StringComparison.CurrentCulture))
            {
                return customXmlPart;
            }
        }

        return null;
    }
}