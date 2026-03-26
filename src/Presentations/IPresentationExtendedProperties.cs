using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the extended presentation properties.
/// </summary>
public interface IPresentationExtendedProperties
{
    /// <summary>
    /// Gets or sets the name of the company.
    /// </summary>
    string? Company { get; set; }
}

internal class PresentationExtendedProperties : IPresentationExtendedProperties
{
    private ExtendedFilePropertiesPart? _extendedFilePropertiesPart { get; set; }


    public PresentationExtendedProperties()
    {
    }

    public PresentationExtendedProperties(ExtendedFilePropertiesPart? extendedFilePropertiesPart)
    {
        if (extendedFilePropertiesPart != null)
        {
            if (!extendedFilePropertiesPart.IsRootElementLoaded)
            {
                extendedFilePropertiesPart?.RootElement?.Reload();
            }
        }

        _extendedFilePropertiesPart = extendedFilePropertiesPart;

    }

    public string? Company
    {
        get => _extendedFilePropertiesPart?.Properties?.Company?.Text;
        set
        {
            if (_extendedFilePropertiesPart?.Properties != null)
            {
                if (_extendedFilePropertiesPart.Properties.Company == null)
                {
                    _extendedFilePropertiesPart.Properties.Company = new Company { Text = value ?? string.Empty };
                }
                else
                {
                    _extendedFilePropertiesPart.Properties.Company.Text = value ?? string.Empty;
                }
            }
        }
    }
}
