using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Units;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a paragraph bullet.
/// </summary>
public sealed class Bullet
{
    private readonly ParagraphProperties aParagraphProperties;
    private readonly Lazy<string?> character;
    private readonly Lazy<string?> colorHex;
    private readonly Lazy<string?> fontName;
    private readonly Lazy<int> size;
    private readonly Lazy<BulletType> type;

    internal Bullet(ParagraphProperties aParagraphProperties)
    {
        this.aParagraphProperties = aParagraphProperties;
        type = new Lazy<BulletType>(ParseType);
        colorHex = new Lazy<string?>(ParseColorHex);
        character = new Lazy<string?>(ParseChar);
        fontName = new Lazy<string?>(ParseFontName);
        size = new Lazy<int>(ParseSize);
    }

    /// <summary>
    ///     Applies default PowerPoint spacing between the bullet and text.
    /// </summary>
    internal void ApplyDefaultSpacing()
    {
        // PowerPoint uses a hanging indent where bullet is at 0 and text starts at 22.5pt.
        var leftMarginEmu = (int)new Points(22.5m).AsEmus();
        aParagraphProperties.LeftMargin = new Int32Value(leftMarginEmu);
        aParagraphProperties.Indent = new Int32Value(-leftMarginEmu);
    }

    private BulletType ParseType()
    {
        if (aParagraphProperties == null)
        {
            return BulletType.None;
        }

        var aAutoNumeredBullet = aParagraphProperties.GetFirstChild<AutoNumberedBullet>();
        if (aAutoNumeredBullet != null)
        {
            return BulletType.Numbered;
        }

        var aPictureBullet = aParagraphProperties.GetFirstChild<PictureBullet>();
        if (aPictureBullet != null)
        {
            return BulletType.Picture;
        }

        var aCharBullet = aParagraphProperties.GetFirstChild<CharacterBullet>();
        if (aCharBullet != null)
        {
            return BulletType.Character;
        }

        return BulletType.None;
    }

    private string? ParseColorHex()
    {
        if (Type == BulletType.None)
        {
            return null;
        }

        var aRgbClrModelHexCollection = aParagraphProperties.Descendants<RgbColorModelHex>();
        if (aRgbClrModelHexCollection.Any())
        {
            return aRgbClrModelHexCollection.Single().Val;
        }

        return null;
    }

    private string? ParseChar()
    {
        if (Type == BulletType.None)
        {
            return null;
        }

        var aCharBullet = aParagraphProperties.GetFirstChild<CharacterBullet>() ??
                          throw new SCException($"This is not {nameof(BulletType.Character)} type bullet.");
        return aCharBullet.Char?.Value;
    }

    private string? ParseFontName()
    {
        if (Type == BulletType.None)
        {
            return null;
        }

        var aBulletFont = aParagraphProperties.GetFirstChild<BulletFont>();
        return aBulletFont?.Typeface?.Value;
    }

    private int ParseSize()
    {
        if (Type == BulletType.None)
        {
            return 0;
        }

        var aBulletSizePercent = aParagraphProperties.GetFirstChild<BulletSizePercentage>();
        var basicPoints = aBulletSizePercent?.Val?.Value ?? 100000;

        return basicPoints / 1000;
    }

    #region Public Properties

    /// <summary>
    ///     Gets RGB color in HEX format.
    /// </summary>
    public string? ColorHex => colorHex.Value;

    /// <summary>
    ///     Gets or sets bullet character. Returns <see langword="null" /> if bullet doesn't exist.
    /// </summary>
    public string? Character
    {
        get => character.Value;
        set
        {
            if (Type != BulletType.Character)
            {
                return;
            }

            var aCharBullet = aParagraphProperties.GetFirstChild<CharacterBullet>();
            if (aCharBullet == null)
            {
                aCharBullet = new CharacterBullet();
                aParagraphProperties.AddChild(aCharBullet);
            }

            aCharBullet.Char = value;
        }
    }

    /// <summary>
    ///     Gets or sets bullet font name. Returns <see langword="null" /> if bullet doesn't exist.
    /// </summary>
    public string? FontName
    {
        get => fontName.Value;
        set
        {
            if (Type == BulletType.None)
            {
                return;
            }

            var aBulletFont = aParagraphProperties.GetFirstChild<BulletFont>();
            if (aBulletFont == null)
            {
                aBulletFont = new BulletFont();
                aParagraphProperties.AddChild(aBulletFont);
            }

            aBulletFont.Typeface = value;
        }
    }

    /// <summary>
    ///     Gets or sets bullet size in percentages of text.
    /// </summary>
    public int Size
    {
        get => size.Value;
        set
        {
            if (aParagraphProperties == null)
            {
                return;
            }

            var aBulletSizePercent = aParagraphProperties.GetFirstChild<BulletSizePercentage>();
            if (aBulletSizePercent == null)
            {
                aBulletSizePercent = new BulletSizePercentage();
                aParagraphProperties.AddChild(aBulletSizePercent);
            }

            aBulletSizePercent.Val = value * 1000;
        }
    }

    /// <summary>
    ///     Gets or sets bullet type.
    /// </summary>
    public BulletType Type
    {
        get => type.Value;
        set
        {
            if (aParagraphProperties == null)
            {
                return;
            }

            var aAutoNumeredBullet = aParagraphProperties.GetFirstChild<AutoNumberedBullet>();
            aParagraphProperties.RemoveChild(aAutoNumeredBullet);

            var aPictureBullet = aParagraphProperties.GetFirstChild<PictureBullet>();
            aParagraphProperties.RemoveChild(aPictureBullet);

            var aCharBullet = aParagraphProperties.GetFirstChild<CharacterBullet>();
            aParagraphProperties.RemoveChild(aCharBullet);

            if (value == BulletType.Numbered)
            {
                var child = new AutoNumberedBullet
                {
                    // replace at property
                    Type = TextAutoNumberSchemeValues.ArabicPeriod
                };

                aParagraphProperties.AddChild(child);
            }

            if (value == BulletType.Picture)
            {
                var child = new PictureBullet();
                aParagraphProperties.AddChild(child);
            }

            if (value == BulletType.Character)
            {
                var child = new CharacterBullet();
                aParagraphProperties.AddChild(child);
            }
        }
    }

    #endregion Public Properties
}