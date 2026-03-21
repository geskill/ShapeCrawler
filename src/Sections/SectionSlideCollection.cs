using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Sections;

internal sealed class SectionSlideCollection(P14.Section p14Section) : IReadOnlyList<IUserSlide>
{
    public int Count => GetSlides().Count;

    public IUserSlide this[int index] => GetSlides()[index];

    public IEnumerator<IUserSlide> GetEnumerator()
    {
        return GetSlides().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private UserSlideCollection GetSlides()
    {
        var p14SectionSlideIdListEntryList = p14Section.Descendants<P14.SectionSlideIdListEntry>();
        var presDocument = new SCOpenXmlElement(p14Section).ParentPresentationDocument;
        var slideParts = new List<SlidePart>();
        var idToRId = presDocument.PresentationPart!.Presentation!.SlideIdList!.ChildElements.OfType<P.SlideId>()
            .ToDictionary(slideId => slideId.Id!, slideId => slideId.RelationshipId);
        foreach (var p14SectionSlideIdListEntry in p14SectionSlideIdListEntryList)
        {
            var rId = idToRId[p14SectionSlideIdListEntry.Id!]!.Value!;
            var slidePart = (SlidePart)presDocument.PresentationPart!.GetPartById(rId);
            slideParts.Add(slidePart);
        }

        return new UserSlideCollection(slideParts);
    }
}