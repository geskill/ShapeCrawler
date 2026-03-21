using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of presentation sections.
/// </summary>
public interface ISectionCollection : IReadOnlyCollection<ISection>
{
    /// <summary>
    ///     Gets the section by index.
    /// </summary>
    ISection this[int index] { get; }

    /// <summary>
    ///     Removes specified section.
    /// </summary>
    void Remove(ISection removingSection);

    /// <summary>
    ///     Gets section by section name.
    /// </summary>
    ISection GetByName(string sectionName);
}

internal sealed class SectionCollection(PresentationDocument presDocument) : ISectionCollection
{
    public int Count => SectionList().Count;

    public ISection this[int index] => SectionList()[index];

    public IEnumerator<ISection> GetEnumerator()
    {
        return SectionList().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Remove(ISection removingSection)
    {
        if (removingSection is not IRemovable removeable)
        {
            throw new SCException("Section cannot be removed.");
        }

        var total = Count;
        removeable.Remove();

        if (total == 1)
        {
            presDocument.PresentationPart!.Presentation!.PresentationExtensionList
                ?.Descendants<P14.SectionList>().First()
                .Remove();
        }
    }

    public ISection GetByName(string sectionName)
    {
        return SectionList().First(section => section.Name == sectionName);
    }

    private List<Section> SectionList()
    {
        var p14SectionList = presDocument.PresentationPart!.Presentation!.PresentationExtensionList
            ?.Descendants<P14.SectionList>().FirstOrDefault();
        return p14SectionList == null
            ? []
            : [.. p14SectionList.OfType<P14.Section>().Select(p14Section => new Section(p14Section))];
    }
}