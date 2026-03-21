using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

/// <summary>
///     Represents a slide layout collection.
/// </summary>
public interface ILayoutSlideCollection : IEnumerable<ILayoutSlide>
{
    /// <summary>
    ///     Gets slide layout by index.
    /// </summary>
    ILayoutSlide this[int index] { get; }
}

internal sealed class LayoutSlideCollection(SlideMasterPart slideMasterPart) : ILayoutSlideCollection
{
    public ILayoutSlide this[int index] => Layouts().ElementAt(index);

    public IEnumerator<ILayoutSlide> GetEnumerator()
    {
        return Layouts().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    internal LayoutSlide Layout(int number)
    {
        return Layouts().First(l => l.Number == number);
    }

    private IEnumerable<LayoutSlide> Layouts()
    {
        var rIdList = slideMasterPart.SlideMaster!.SlideLayoutIdList!.OfType<P.SlideLayoutId>()
            .Select(layoutId => layoutId.RelationshipId!);
        foreach (var rId in rIdList)
        {
            var layoutPart = (SlideLayoutPart)slideMasterPart.GetPartById(rId.Value!);

            yield return new LayoutSlide(layoutPart);
        }
    }
}