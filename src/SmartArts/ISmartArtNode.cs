using ShapeCrawler.SmartArts;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a node in a SmartArt graphic.
/// </summary>
public interface ISmartArtNode
{
    /// <summary>
    ///     Gets or sets the text of the SmartArt node.
    /// </summary>
    string Text { get; set; }
}

/// <summary>
///     Represents a node in a SmartArt graphic.
/// </summary>
internal class SmartArtNode : ISmartArtNode
{
    private readonly SmartArtNodeCollection nodeCollection;
    private string textValue;

    internal SmartArtNode(string modelId, string text, SmartArtNodeCollection nodeCollection)
    {
        ModelId = modelId;
        textValue = text;
        this.nodeCollection = nodeCollection;
    }

    internal string ModelId { get; }

    /// <summary>
    ///     Gets or sets the text of the SmartArt node.
    /// </summary>
    public string Text
    {
        get => textValue;
        set
        {
            if (textValue != value)
            {
                textValue = value;
                nodeCollection?.UpdateNodeText(ModelId, value);
            }
        }
    }

    internal void UpdateText(string text)
    {
        textValue = text;
    }
}