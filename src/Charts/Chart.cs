using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class Chart : IBarChart, IColumnChart, ILineChart, IPieChart, IScatterChart, IBubbleChart,
    IAreaChart
{
    private readonly Categories? categories;
    private readonly ChartPart chartPart;
    private readonly Lazy<ChartTitle> chartTitle;
    private readonly ShapeFill fill;
    private readonly SlideShapeOutline outline;
    private readonly SeriesCollection seriesCollection;
    private readonly XAxis? xAxis;

    internal Chart(
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart,
        Categories categories,
        XAxis xAxis)
    {
        this.seriesCollection = seriesCollection;
        this.outline = outline;
        this.fill = fill;
        this.chartPart = chartPart;
        this.categories = categories;
        this.xAxis = xAxis;
        chartTitle = new Lazy<ChartTitle>(() =>
            new ChartTitle(chartPart, Type, SeriesCollection, new ChartTitleAlignment(chartPart)));
    }

    internal Chart(
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart,
        XAxis xAxis)
    {
        this.seriesCollection = seriesCollection;
        this.outline = outline;
        this.fill = fill;
        this.chartPart = chartPart;
        this.xAxis = xAxis;
        chartTitle = new Lazy<ChartTitle>(() =>
            new ChartTitle(chartPart, Type, SeriesCollection, new ChartTitleAlignment(chartPart)));
    }

    internal Chart(
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart,
        Categories categories)
    {
        this.seriesCollection = seriesCollection;
        this.outline = outline;
        this.fill = fill;
        this.chartPart = chartPart;
        this.categories = categories;
        chartTitle = new Lazy<ChartTitle>(() =>
            new ChartTitle(chartPart, Type, SeriesCollection, new ChartTitleAlignment(chartPart)));
    }

    public IShapeOutline Outline => outline;

    public IShapeFill Fill => fill;

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("It is not possible to set the geometry type for the chart shape.");
    }

    public ChartType Type
    {
        get
        {
            var plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            if (cXCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            var chartName = cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType enumChartType);

            return enumChartType;
        }
    }

    public IChartTitle? Title => chartTitle.Value;

    public IReadOnlyList<ICategory>? Categories => categories;

    public IXAxis? XAxis => xAxis;

    public ISeriesCollection SeriesCollection => seriesCollection;

    public byte[] GetWorksheetByteArray()
    {
        return new Workbook(chartPart.EmbeddedPackagePart!).AsByteArray();
    }
}