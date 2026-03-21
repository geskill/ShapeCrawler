using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Texts;

internal abstract class TextBox : ITextBox
{
    private readonly TextAutofit autofit;
    private readonly TextBoxMargins margins;
    private readonly OpenXmlElement textBody;
    private TextDirection? textDirection;
    private TextVerticalAlignment? vAlignment;

    private protected TextBox(TextBoxMargins margins, OpenXmlElement textBody)
    {
        this.margins = margins;
        this.textBody = textBody;
        var shapeSize = new ShapeSize(textBody.Parent!);
        autofit = new TextAutofit(
            Paragraphs,
            () => AutofitType,
            shapeSize,
            this.margins,
            () => TextWrapped,
            this.textBody);
    }

    public IParagraphCollection Paragraphs => new ParagraphCollection(textBody);

    public string Text
    {
        get
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(Paragraphs[0].Text);

            var paragraphsCount = Paragraphs.Count;
            var index = 1; // we've already added the text of first paragraph
            while (index < paragraphsCount)
            {
                stringBuilder.AppendLine();
                stringBuilder.Append(Paragraphs[index].Text);

                index++;
            }

            return stringBuilder.ToString();
        }
    }

    public AutofitType AutofitType
    {
        get
        {
            var aBodyPr = textBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
            {
                return AutofitType.Shrink;
            }

            if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
            {
                return AutofitType.Resize;
            }

            return AutofitType.None;
        }

        set
        {
            var currentType = AutofitType;
            if (currentType == value)
            {
                return;
            }

            var aBodyPr = textBody.GetFirstChild<A.BodyProperties>()!;

            RemoveExistingAutofitElements(aBodyPr);

            switch (value)
            {
                case AutofitType.None:
                    aBodyPr.Append(new A.NoAutoFit());
                    break;
                case AutofitType.Shrink:
                    aBodyPr.Append(new A.NormalAutoFit());
                    break;
                case AutofitType.Resize:
                    aBodyPr.Append(new A.ShapeAutoFit());
                    ResizeParentShapeOnDemand();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }

    public decimal LeftMargin
    {
        get => margins.Left;
        set => margins.Left = value;
    }

    public decimal RightMargin
    {
        get => margins.Right;
        set => margins.Right = value;
    }

    public decimal TopMargin
    {
        get => margins.Top;
        set => margins.Top = value;
    }

    public decimal BottomMargin
    {
        get => margins.Bottom;
        set => margins.Bottom = value;
    }

    public bool TextWrapped
    {
        get
        {
            var aBodyPr = textBody.GetFirstChild<A.BodyProperties>()!;
            var wrap = aBodyPr.GetAttributes().FirstOrDefault(a => a.LocalName == "wrap");

            if (wrap.Value == "none")
            {
                return false;
            }

            return true;
        }
    }

    public string SdkXPath => new XmlPath(textBody).XPath;

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (vAlignment.HasValue)
            {
                return vAlignment.Value;
            }

            var aBodyPr = textBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.Anchor?.Value == A.TextAnchoringTypeValues.Center)
            {
                vAlignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr.Anchor?.Value == A.TextAnchoringTypeValues.Bottom)
            {
                vAlignment = TextVerticalAlignment.Bottom;
            }
            else
            {
                vAlignment = TextVerticalAlignment.Top;
            }

            return vAlignment.Value;
        }

        set => SetVerticalAlignment(value);
    }

    public TextDirection TextDirection
    {
        get
        {
            if (!textDirection.HasValue)
            {
                var textDirectionVal = textBody.GetFirstChild<A.BodyProperties>()!.Vertical?.Value;

                if (textDirectionVal == A.TextVerticalValues.Vertical)
                {
                    textDirection = TextDirection.Rotate90;
                }
                else if (textDirectionVal == A.TextVerticalValues.Vertical270)
                {
                    textDirection = TextDirection.Rotate270;
                }
                else if (textDirectionVal == A.TextVerticalValues.WordArtVertical)
                {
                    textDirection = TextDirection.Stacked;
                }
                else
                {
                    textDirection = TextDirection.Horizontal;
                }
            }

            return textDirection.Value;
        }

        set => SetTextDirection(value);
    }

    public void SetMarkdownText(string text)
    {
        var markdownText = new MarkdownText(
            text,
            Paragraphs,
            () => AutofitType,
            autofit.ShrinkFont,
            autofit.Apply);
        markdownText.ApplyTo();
    }

    public void SetText(string text)
    {
        var textContent = new TextContent(
            text,
            Paragraphs,
            () => AutofitType,
            autofit.ShrinkFont,
            autofit.Apply);
        textContent.ApplyTo();
    }

    /// <summary>
    ///     Disables text wrapping in the text box.
    /// </summary>
    internal void DisableWrapping()
    {
        var bodyProperties = textBody.GetFirstChild<A.BodyProperties>()!;
        bodyProperties.SetAttribute(new OpenXmlAttribute("wrap", string.Empty, "none"));
    }

    internal void ResizeParentShapeOnDemand()
    {
        autofit.Apply();
    }

    private static void RemoveExistingAutofitElements(A.BodyProperties bodyProperties)
    {
        bodyProperties.GetFirstChild<A.NoAutoFit>()?.Remove();
        bodyProperties.GetFirstChild<A.NormalAutoFit>()?.Remove();
        bodyProperties.GetFirstChild<A.ShapeAutoFit>()?.Remove();
    }

    private void SetVerticalAlignment(TextVerticalAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            TextVerticalAlignment.Top => A.TextAnchoringTypeValues.Top,
            TextVerticalAlignment.Middle => A.TextAnchoringTypeValues.Center,
            TextVerticalAlignment.Bottom => A.TextAnchoringTypeValues.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
        };

        var aBodyPr = textBody.GetFirstChild<A.BodyProperties>();

        aBodyPr?.Anchor = aTextAlignmentTypeValue;

        vAlignment = alignmentValue;
    }

    private void SetTextDirection(TextDirection direction)
    {
        var aBodyPr = textBody.GetFirstChild<A.BodyProperties>()!;

        aBodyPr.Vertical = direction switch
        {
            TextDirection.Rotate90 => A.TextVerticalValues.Vertical,
            TextDirection.Rotate270 => A.TextVerticalValues.Vertical270,
            TextDirection.Stacked => A.TextVerticalValues.WordArtVertical,
            _ => A.TextVerticalValues.Horizontal
        };

        textDirection = direction;
    }
}