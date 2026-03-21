using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class ShapeFillImage : IImage
{
    private readonly A.Blip aBlip;
    private ImagePart imagePart;

    internal ShapeFillImage(A.Blip aBlip, ImagePart imagePart)
    {
        this.aBlip = aBlip;
        this.imagePart = imagePart;
    }

    public string Mime => imagePart.ContentType;

    public string Name => Path.GetFileName(imagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var openXmlPart = aBlip.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var isSharedImagePart =
            openXmlPart.GetPartsOfType<ImagePart>().Count(imagePart => imagePart == this.imagePart) > 1;
        if (isSharedImagePart)
        {
            var rId = RelationshipId.New();
            imagePart = openXmlPart.AddNewPart<ImagePart>("image/png", rId);
            aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.imagePart.FeedData(stream);
    }

    public byte[] AsByteArray()
    {
        return new SCImagePart(imagePart).AsBytes();
    }
}