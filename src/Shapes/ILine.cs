using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine
{
    /// <summary>
    ///     Gets the start point of the line.
    /// </summary>
    Point StartPoint { get; }

    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    Point EndPoint { get; }
}

internal sealed class Line(P.ConnectionShape pConnectionShape, LineShape parentLineShape) : ILine
{
    private readonly P.ConnectionShape connectionShape = pConnectionShape;
    private readonly LineShape lineShape = parentLineShape;

    public Geometry GeometryType => Geometry.Line;

    public Point StartPoint
    {
        get
        {
            var aTransform2D = connectionShape.GetFirstChild<P.ShapeProperties>()!.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            var startX = flipH ? lineShape.X + lineShape.Width : lineShape.X;
            var startY = flipV ? lineShape.Y + lineShape.Height : lineShape.Y;
            return new Point(startX, startY);
        }
    }

    public Point EndPoint
    {
        get
        {
            var aTransform2D = connectionShape.GetFirstChild<P.ShapeProperties>()!.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            var endX = flipH ? lineShape.X : lineShape.X + lineShape.Width;
            var endY = flipV ? lineShape.Y : lineShape.Y + lineShape.Height;
            return new Point(endX, endY);
        }
    }
}