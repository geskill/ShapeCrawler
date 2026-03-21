using System;

namespace ShapeCrawler.Drawing;

internal struct RelationshipId
{
    internal static string New()
    {
        return $"rId-{Guid.NewGuid().ToString("N")[..5]}";
    }
}