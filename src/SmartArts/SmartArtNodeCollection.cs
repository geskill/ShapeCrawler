using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ShapeCrawler.SmartArts;

/// <summary>
///     Represents a collection of SmartArt nodes.
/// </summary>
internal class SmartArtNodeCollection : ISmartArtNodeCollection
{
    private readonly List<SmartArtNode> nodes = [];
    private int nextNodeId = 1;

    /// <summary>
    ///     Gets the number of nodes in the collection.
    /// </summary>
    public int Count => nodes.Count;

    /// <summary>
    ///     Adds a new node to the SmartArt graphic with the specified text.
    /// </summary>
    /// <param name="text">The text for the new node.</param>
    /// <returns>The newly added SmartArt node.</returns>
    public ISmartArtNode AddNode(string text)
    {
        var nodeId = $"p{nextNodeId++}";
        var node = new SmartArtNode(nodeId, text, this);
        nodes.Add(node);
        return node;
    }

    public IEnumerator<ISmartArtNode> GetEnumerator()
    {
        return nodes.Cast<ISmartArtNode>().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    internal void UpdateNodeText(string nodeId, string text)
    {
        var node = nodes.FirstOrDefault(n => n.ModelId == nodeId);
        node?.UpdateText(text);
    }
}