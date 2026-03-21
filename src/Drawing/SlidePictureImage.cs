using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class SlidePictureImage : IImage
{
    private readonly A.Blip aBlip;
    private readonly OpenXmlPart openXmlPart;
    private ImagePart imagePart;

    internal SlidePictureImage(A.Blip aBlip)
    {
        this.aBlip = aBlip;
        openXmlPart = aBlip.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        imagePart = (ImagePart)openXmlPart.GetPartById(aBlip.Embed!.Value!);
    }

    public string Mime => imagePart.ContentType;

    public string Name => Path.GetFileName(imagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var presDocument = (PresentationDocument)openXmlPart.OpenXmlPackage;
        var slideParts = presDocument.PresentationPart!.SlideParts;
        var allABlips =
            slideParts.SelectMany(slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!.Descendants<A.Blip>());
        var isSharedImagePart = allABlips.Count(blip => blip.Embed!.Value == aBlip.Embed!.Value) > 1;
        if (isSharedImagePart)
        {
            var rId = RelationshipId.New();
            imagePart = openXmlPart.AddNewPart<ImagePart>("image/png", rId);
            aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        imagePart.FeedData(stream);
    }

    public byte[] AsByteArray()
    {
        return new SCImagePart(imagePart).AsBytes();
    }
}