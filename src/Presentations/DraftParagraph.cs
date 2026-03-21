using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft paragraph for fluent API.
/// </summary>
public sealed class DraftParagraph
{
    internal string? Content { get; private set; }

    internal bool IsBulletedList { get; private set; }

    internal string BulletCharacter { get; private set; } = "•";

    /// <summary>
    ///     Gets draft font.
    /// </summary>
    internal DraftFont? FontDraft { get; private set; }

    /// <summary>
    ///     Gets draft indentation.
    /// </summary>
    internal DraftIndentation? IndentationDraft { get; private set; }

    /// <summary>
    ///     Sets paragraph text.
    /// </summary>
    public DraftParagraph Text(string text)
    {
        Content = text;
        return this;
    }

    /// <summary>
    ///     Configures font using a nested builder.
    /// </summary>
    public DraftParagraph Font(Action<DraftFont> configure)
    {
        FontDraft = new DraftFont();
        configure(FontDraft);
        return this;
    }

    /// <summary>
    ///     Configures indentation using a nested builder.
    /// </summary>
    public DraftParagraph Indentation(Action<DraftIndentation> configure)
    {
        IndentationDraft = new DraftIndentation();
        configure(IndentationDraft);
        return this;
    }

    /// <summary>
    ///     Makes this paragraph a bulleted list item.
    /// </summary>
    public DraftParagraph BulletedList()
    {
        IsBulletedList = true;
        return this;
    }

    /// <summary>
    ///     Makes this paragraph a bulleted list item with a custom character.
    /// </summary>
    public DraftParagraph BulletedList(string character)
    {
        IsBulletedList = true;
        BulletCharacter = character;
        return this;
    }
}