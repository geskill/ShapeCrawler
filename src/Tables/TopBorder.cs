using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Units;

namespace ShapeCrawler.Tables;

internal class TopBorder(TableCellProperties aTableCellProperties) : IBorder
{
    public decimal Width
    {
        get => GetWidth();
        set => UpdateWidth(value);
    }

    public string? Color { get => GetColor(); set => SetColor(value!); }

    private string? GetColor()
    {
        return aTableCellProperties.TopBorderLineProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
    }

    private void SetColor(string color)
    {
        aTableCellProperties.TopBorderLineProperties ??= new TopBorderLineProperties
        {
            Width = (Int32Value)new Points(1).AsEmus()
        };

        var solidFill = aTableCellProperties.TopBorderLineProperties.GetFirstChild<SolidFill>();

        if (solidFill is null)
        {
            solidFill = new SolidFill();
            aTableCellProperties.TopBorderLineProperties.AppendChild(solidFill);
        }

        solidFill.RgbColorModelHex ??= new RgbColorModelHex();

        solidFill.RgbColorModelHex.Val = new HexBinaryValue(color);
    }

    private void UpdateWidth(decimal points)
    {
        if (aTableCellProperties.TopBorderLineProperties is null)
        {
            var aSolidFill = new SolidFill { SchemeColor = new SchemeColor { Val = SchemeColorValues.Text1 } };
            aTableCellProperties.TopBorderLineProperties = new TopBorderLineProperties();
            aTableCellProperties.TopBorderLineProperties.AppendChild(aSolidFill);
        }

        var emus = new Points(points).AsEmus();
        aTableCellProperties.TopBorderLineProperties.Width = new Int32Value((int)emus);
    }

    private decimal GetWidth()
    {
        if (aTableCellProperties.TopBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = aTableCellProperties.TopBorderLineProperties!.Width!.Value;

        return new Emus(emus).AsPoints();
    }
}