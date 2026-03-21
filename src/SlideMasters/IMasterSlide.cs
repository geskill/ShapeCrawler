using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a PowerPoint Slide Master.
/// </summary>
public interface IMasterSlide
{
    /// <summary>
    ///     Gets slide master order number.
    /// </summary>
    int Number { get; }

    /// <summary>
    ///     Gets background image if slide master has background, otherwise <see langword="null" />.
    /// </summary>
    IImage? Background { get; }

    /// <summary>
    ///     Gets slide layout collection.
    /// </summary>
    ILayoutSlideCollection LayoutSlides { get; }

    /// <summary>
    ///     Gets the collection of master shapes.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }

    /// <summary>
    ///     Gets slide number. Returns <see langword="null" /> if slide master does not have slide number.
    /// </summary>
    IMasterSlideNumber? SlideNumber { get; }

    /// <summary>
    ///     Returns a shape from the slide master.
    /// </summary>
    /// <param name="shape">The name of the shape.</param>
    /// <returns>The requested shape.</returns>
    IShape Shape(string shape);

    /// <summary>
    ///     Gets slide layout by name.
    /// </summary>
    ILayoutSlide SlideLayout(string name);

    /// <summary>
    ///     Gets slide layout by number.
    /// </summary>
    ILayoutSlide SlideLayout(int number);
}

internal sealed class MasterSlide : IMasterSlide
{
    private readonly LayoutSlideCollection layouts;
    private readonly SlideMasterPart slideMasterPart;
    private readonly Lazy<MasterSlideNumber?> slideNumber;

    internal MasterSlide(SlideMasterPart slideMasterPart)
    {
        this.slideMasterPart = slideMasterPart;
        layouts = new LayoutSlideCollection(slideMasterPart);
        slideNumber = new Lazy<MasterSlideNumber?>(CreateSlideNumber);
        Shapes = new ShapeCollection(this.slideMasterPart);
    }

    public IImage? Background => null;

    public ILayoutSlideCollection LayoutSlides => layouts;

    public IShapeCollection Shapes { get; }

    public ITheme Theme => new Theme(slideMasterPart, slideMasterPart.ThemePart!.Theme!);

    public IMasterSlideNumber? SlideNumber => slideNumber.Value;

    public int Number
    {
        get
        {
            var match = Regex.Match(slideMasterPart.Uri.ToString(), @"\d+", RegexOptions.None, TimeSpan.FromSeconds(1));
            return int.Parse(match.Value);
        }
    }

    public IShape Shape(string shape)
    {
        return Shapes.Shape(shape);
    }

    public ILayoutSlide SlideLayout(string name)
    {
        return layouts.First(l => l.Name == name);
    }

    public ILayoutSlide SlideLayout(int number)
    {
        return InternalSlideLayout(number);
    }

    internal LayoutSlide InternalSlideLayout(int number)
    {
        return layouts.Layout(number);
    }

    private MasterSlideNumber? CreateSlideNumber()
    {
        var pSldNum = slideMasterPart.SlideMaster!.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .FirstOrDefault(s =>
                s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value ==
                P.PlaceholderValues.SlideNumber);

        return pSldNum is null ? null : new MasterSlideNumber(pSldNum);
    }
}