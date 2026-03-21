using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Groups;

internal sealed class GroupedShape : Shape
{
    private readonly P.Shape pShape;

    internal GroupedShape(Position position, ShapeSize shapeSize, ShapeId shapeId, P.Shape pShape)
        : base(position, shapeSize, shapeId, pShape)
    {
        this.pShape = pShape;
    }

    public override decimal X
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.X;
            }

            var absoluteX = base.X;

            // Apply the formula for each parent group in the hierarchy, from innermost to outermost
            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childOffset = transformGroup.ChildOffset!;
                var childExtents = transformGroup.ChildExtents!;
                var offset = transformGroup.Offset!;
                var extents = transformGroup.Extents!;

                // Calculate scale factor (ratio of group extents to child extents)
                var scaleFactor = 1.0m;
                if (childExtents.Cx!.Value != 0)
                {
                    scaleFactor = (decimal)extents.Cx!.Value / childExtents.Cx!.Value;
                }

                // Apply the formula: (childOffset - groupChildOffset) * scaleFactor + groupOffset
                var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
                absoluteX = ((absoluteX - childOffsetX) * scaleFactor) + new Emus(offset.X!.Value).AsPoints();
            }

            return absoluteX;
        }

        set
        {
            base.X = LocalX(value);
            var pGroupShape = pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedShapeXEmus = new Points(value).AsEmus();
            var groupShapeXEmus = aOffset.X!.Value;

            if (groupedShapeXEmus < groupShapeXEmus)
            {
                var diffParent = groupShapeXEmus - groupedShapeXEmus;
                var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
                aOffset.X = new Int64Value(aOffset.X!.Value - diffParent);
                aExtents.Cx = new Int64Value(aExtents.Cx!.Value + diffParent);
                aChildOffset.X = new Int64Value(aChildOffset.X!.Value - diffChild);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value + diffChild);

                return;
            }

            var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
            var groupedRightEmu = new Points(X + Width).AsEmus();
            if (groupedRightEmu > groupRightEmu)
            {
                var diffParent = groupedRightEmu - groupRightEmu;
                var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
                aExtents.Cx = new Int64Value(aExtents.Cx!.Value + diffParent);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value + diffChild);
            }
        }
    }

    public override decimal Y
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.Y;
            }

            // Start with the shape's relative Y coordinate
            var absoluteY = base.Y;

            // Apply the formula for each parent group in the hierarchy, from innermost to outermost
            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childOffset = transformGroup.ChildOffset!;
                var childExtents = transformGroup.ChildExtents!;
                var offset = transformGroup.Offset!;
                var extents = transformGroup.Extents!;

                // Calculate scale factor (ratio of group extents to child extents)
                var scaleFactor = 1.0m;
                if (childExtents.Cy!.Value != 0)
                {
                    scaleFactor = (decimal)extents.Cy!.Value / childExtents.Cy!.Value;
                }

                // Apply the formula: (childOffset - groupChildOffset) * scaleFactor + groupOffset
                var childOffsetY = new Emus(childOffset.Y!.Value).AsPoints();
                absoluteY = ((absoluteY - childOffsetY) * scaleFactor) + new Emus(offset.Y!.Value).AsPoints();
            }

            return absoluteY;
        }

        set
        {
            base.Y = LocalY(value);
            var pGroupShape = pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedYEmu = new Points(value).AsEmus();
            var groupYEmu = aOffset.Y!.Value;
            if (groupedYEmu < groupYEmu)
            {
                var diffParent = groupYEmu - groupedYEmu;
                var diffChild = ChildDiff(diffParent, aExtents.Cy!.Value, aChildExtents.Cy!.Value);
                aOffset.Y = new Int64Value(aOffset.Y!.Value - diffParent);
                aExtents.Cy = new Int64Value(aExtents.Cy!.Value + diffParent);
                aChildOffset.Y = new Int64Value(aChildOffset.Y!.Value - diffChild);
                aChildExtents.Cy = new Int64Value(aChildExtents.Cy!.Value + diffChild);

                return;
            }

            var groupBottomEmu = aOffset.Y!.Value + aExtents.Cy!.Value;
            var groupedBottomEmu = groupedYEmu + new Points(Height).AsEmus();
            if (groupedBottomEmu > groupBottomEmu)
            {
                var diffParent = groupedBottomEmu - groupBottomEmu;
                var diffChild = ChildDiff(diffParent, aExtents.Cy!.Value, aChildExtents.Cy!.Value);
                aExtents.Cy = new Int64Value(aExtents.Cy!.Value + diffParent);
                aChildExtents.Cy = new Int64Value(aChildExtents.Cy!.Value + diffChild);
            }
        }
    }

    public override decimal Width
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.Width;
            }

            // Calculate cumulative scale factor through all parent groups
            var cumulativeScaleFactor = 1.0m;

            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childExtentsWidth = transformGroup.ChildExtents!.Cx!.Value;
                var extentsWidth = transformGroup.Extents!.Cx!.Value;

                // Skip if either value is zero to avoid division by zero
                if (childExtentsWidth == 0)
                {
                    continue;
                }

                var scaleFactor = (decimal)extentsWidth / childExtentsWidth;
                cumulativeScaleFactor *= scaleFactor;
            }

            return base.Width * cumulativeScaleFactor;
        }

        set
        {
            base.Width = LocalWidth(value);
            var pGroupShape = pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var groupedShapeWidthEmus = new Points(value).AsEmus();
            var groupShapeWidthEmus = aExtents.Cx!.Value;

            if (groupedShapeWidthEmus < groupShapeWidthEmus)
            {
                var diffParent = groupShapeWidthEmus - groupedShapeWidthEmus;
                var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
                aExtents.Cx = new Int64Value(aExtents.Cx!.Value - diffParent);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value - diffChild);

                return;
            }

            var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
            var groupedRightEmu = new Points(X + Width).AsEmus();
            if (groupedRightEmu > groupRightEmu)
            {
                var diffParent = groupedRightEmu - groupRightEmu;
                var diffChild = ChildDiff(diffParent, aExtents.Cx!.Value, aChildExtents.Cx!.Value);
                aExtents.Cx = new Int64Value(aExtents.Cx!.Value + diffParent);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx!.Value + diffChild);
            }
        }
    }

    public override decimal Height
    {
        get
        {
            // Get all ancestor group shapes to account for nested groups
            var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
            if (pGroupShapes.Length == 0)
            {
                return base.Height;
            }

            // Calculate cumulative scale factor through all parent groups
            var cumulativeScaleFactor = 1.0m;

            foreach (var pGroupShape in pGroupShapes)
            {
                var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
                var childExtentsCy = transformGroup.ChildExtents!.Cy!.Value;
                var extentsCy = transformGroup.Extents!.Cy!.Value;

                // Skip if either value is zero to avoid division by zero
                if (childExtentsCy == 0)
                {
                    continue;
                }

                var scaleFactor = (decimal)extentsCy / childExtentsCy;
                cumulativeScaleFactor *= scaleFactor;
            }

            return base.Height * cumulativeScaleFactor;
        }
        set => base.Height = LocalHeight(value);
    }

    private static long ChildDiff(long parentDiff, long extents, long childExtents)
    {
        if (parentDiff == 0)
        {
            return 0;
        }

        if (childExtents == 0)
        {
            return parentDiff;
        }

        var scaleFactor = (decimal)extents / childExtents;
        if (scaleFactor == 0)
        {
            return parentDiff;
        }

        return (long)decimal.Round(parentDiff / scaleFactor, 0, MidpointRounding.AwayFromZero);
    }

    private decimal LocalX(decimal absoluteX)
    {
        var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return absoluteX;
        }

        var localX = absoluteX;
        for (var i = pGroupShapes.Length - 1; i >= 0; i--)
        {
            var pGroupShape = pGroupShapes[i];
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childOffset = transformGroup.ChildOffset!;
            var childExtents = transformGroup.ChildExtents!;
            var offset = transformGroup.Offset!;
            var extents = transformGroup.Extents!;

            var scaleFactor = 1.0m;
            if (childExtents.Cx!.Value != 0)
            {
                scaleFactor = (decimal)extents.Cx!.Value / childExtents.Cx!.Value;
            }

            if (scaleFactor == 0)
            {
                scaleFactor = 1.0m;
            }

            var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
            var offsetX = new Emus(offset.X!.Value).AsPoints();
            localX = ((localX - offsetX) / scaleFactor) + childOffsetX;
        }

        return localX;
    }

    private decimal LocalY(decimal absoluteY)
    {
        var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return absoluteY;
        }

        var localY = absoluteY;
        for (var i = pGroupShapes.Length - 1; i >= 0; i--)
        {
            var pGroupShape = pGroupShapes[i];
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childOffset = transformGroup.ChildOffset!;
            var childExtents = transformGroup.ChildExtents!;
            var offset = transformGroup.Offset!;
            var extents = transformGroup.Extents!;

            var scaleFactor = 1.0m;
            if (childExtents.Cy!.Value != 0)
            {
                scaleFactor = (decimal)extents.Cy!.Value / childExtents.Cy!.Value;
            }

            if (scaleFactor == 0)
            {
                scaleFactor = 1.0m;
            }

            var childOffsetY = new Emus(childOffset.Y!.Value).AsPoints();
            var offsetY = new Emus(offset.Y!.Value).AsPoints();
            localY = ((localY - offsetY) / scaleFactor) + childOffsetY;
        }

        return localY;
    }

    private decimal LocalWidth(decimal absoluteWidth)
    {
        var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return absoluteWidth;
        }

        var cumulativeScaleFactor = 1.0m;

        foreach (var pGroupShape in pGroupShapes)
        {
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childExtentsWidth = transformGroup.ChildExtents!.Cx!.Value;
            if (childExtentsWidth == 0)
            {
                continue;
            }

            var extentsWidth = transformGroup.Extents!.Cx!.Value;
            var scaleFactor = (decimal)extentsWidth / childExtentsWidth;
            if (scaleFactor == 0)
            {
                scaleFactor = 1.0m;
            }

            cumulativeScaleFactor *= scaleFactor;
        }

        if (cumulativeScaleFactor == 0)
        {
            return absoluteWidth;
        }

        return absoluteWidth / cumulativeScaleFactor;
    }

    private decimal LocalHeight(decimal absoluteHeight)
    {
        var pGroupShapes = pShape.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return absoluteHeight;
        }

        var cumulativeScaleFactor = 1.0m;

        foreach (var pGroupShape in pGroupShapes)
        {
            var transformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var childExtentsHeight = transformGroup.ChildExtents!.Cy!.Value;
            if (childExtentsHeight == 0)
            {
                continue;
            }

            var extentsHeight = transformGroup.Extents!.Cy!.Value;
            var scaleFactor = (decimal)extentsHeight / childExtentsHeight;
            if (scaleFactor == 0)
            {
                scaleFactor = 1.0m;
            }

            cumulativeScaleFactor *= scaleFactor;
        }

        if (cumulativeScaleFactor == 0)
        {
            return absoluteHeight;
        }

        return absoluteHeight / cumulativeScaleFactor;
    }
}