using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Positions;

internal sealed class Position(OpenXmlElement pShapeTreeElement)
{
    internal decimal X
    {
        get
        {
            var emus = GetAOffset().X!.Value;
            return new Emus(emus).AsPoints();
        }

        set
        {
            var emus = new Points(value).AsEmus();
            GetAOffset().X = new Int64Value(emus);
        }
    }

    internal decimal Y
    {
        get
        {
            var emus = GetAOffset().Y!.Value;
            return new Emus(emus).AsPoints();
        }

        set
        {
            var emus = new Points(value).AsEmus();
            GetAOffset().Y = new Int64Value(emus);
        }
    }

    private A.Offset GetAOffset()
    {
        var aOffset = pShapeTreeElement.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset != null)
        {
            return aOffset;
        }

        return new ReferencedPShape(pShapeTreeElement).ATransform2D().Offset!;
    }
}