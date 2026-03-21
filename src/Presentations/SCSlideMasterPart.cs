using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

internal readonly ref struct SCSlideMasterPart
{
    private readonly SlideMasterPart slideMasterPart;

    internal SCSlideMasterPart(SlideMasterPart slideMasterPart)
    {
        this.slideMasterPart = slideMasterPart;
    }

    internal void RemoveLayoutsExcept(SlideLayoutPart exceptSlideLayoutPart)
    {
        var pSlideLayoutIds = slideMasterPart.SlideMaster!.SlideLayoutIdList!.OfType<SlideLayoutId>();
        foreach (var slideLayoutPart in slideMasterPart.SlideLayoutParts.ToList())
        {
            if (slideLayoutPart == exceptSlideLayoutPart)
            {
                continue;
            }

            var id = slideMasterPart.GetIdOfPart(slideLayoutPart);
            var layoutId = pSlideLayoutIds.First(x => x.RelationshipId == id);
            layoutId.Remove();
            slideMasterPart.DeletePart(slideLayoutPart);
        }
    }
}