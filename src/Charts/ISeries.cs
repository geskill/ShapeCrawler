using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using C = DocumentFormat.OpenXml.Drawing.Charts;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a chart series.
/// </summary>
public interface ISeries
{
    /// <summary>
    ///     Gets series name.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets chart type.
    /// </summary>
    ChartType Type { get; }

    /// <summary>
    ///     Gets the collection of chart points.
    /// </summary>
    IReadOnlyList<IChartPoint> Points { get; }

    /// <summary>
    ///     Gets the collection of X-values points of the series.
    ///     Returns <see langword="null" /> when the series doesn't support X-values.
    /// </summary>
    IReadOnlyList<IChartPoint>? XPoints { get; }

    /// <summary>
    ///     Gets the collection of bubble size points of the series.
    ///     Returns <see langword="null" /> when the series doesn't support bubble size values.
    /// </summary>
    IReadOnlyList<IChartPoint>? BubbleSizePoints { get; }

    /// <summary>
    ///     Gets a value indicating whether chart has name.
    /// </summary>
    bool HasName { get; }
}

internal sealed class Series : ISeries
{
    private readonly ChartPart chartPart;
    private readonly OpenXmlElement cSer;

    internal Series(ChartPart sdkChartPart, OpenXmlElement cSer, ChartType type)
    {
        chartPart = sdkChartPart;
        this.cSer = cSer;
        Type = type;
        Points = new ChartPoints(chartPart, this.cSer);
        XPoints = type is ChartType.ScatterChart or ChartType.BubbleChart
            ? new SeriesXPoints(chartPart, this.cSer)
            : null;
        BubbleSizePoints = type is ChartType.BubbleChart
            ? new SeriesBubbleSizePoints(chartPart, this.cSer)
            : null;
    }

    public ChartType Type { get; }

    public IReadOnlyList<IChartPoint> Points { get; }

    public IReadOnlyList<IChartPoint>? XPoints { get; }

    public IReadOnlyList<IChartPoint>? BubbleSizePoints { get; }

    public bool HasName => cSer.GetFirstChild<C.SeriesText>()?.StringReference != null;

    public string Name => ParseName();

    private string ParseName()
    {
        var cStrRef = cSer.GetFirstChild<C.SeriesText>()?.StringReference ??
                      throw new SCException(
                          $"Series does not have name. Use {nameof(HasName)} property to check if series has name.");
        var fromCache = cStrRef.StringCache?.GetFirstChild<C.StringPoint>()!.Single().InnerText;

        return fromCache ?? new Workbook(chartPart.EmbeddedPackagePart!).FormulaValues(cStrRef.Formula!.Text)[0]
            .ToString();
    }
}