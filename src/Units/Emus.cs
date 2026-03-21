namespace ShapeCrawler.Units;

internal readonly ref struct Emus(long emus)
{
    internal decimal AsPoints()
    {
        return emus / 12700m;
    }

    internal decimal AsPixels()
    {
        return emus / 9525m;
    }
}