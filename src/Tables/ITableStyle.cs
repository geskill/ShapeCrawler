using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a table style.
/// </summary>
public interface ITableStyle
{
    /// <summary>
    ///     Gets the name.
    /// </summary>
    string Name { get; }
}

internal class TableStyle(string name) : ITableStyle
{
    public string Guid { get; init; } = string.Empty;
    public string Name { get; } = name;

    public override bool Equals(object? obj)
    {
        return obj is TableStyle style &&
               Name == style.Name &&
               Guid == style.Guid;
    }

    public override int GetHashCode()
    {
        var hashCode = 1242478914;
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(Name);
        hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(Guid);
        return hashCode;
    }
}