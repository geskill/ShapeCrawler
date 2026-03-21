namespace ShapeCrawler.Examples;

public class ShapeCollectionExamples
{
    [Test]
    [Explicit]
    public void Groups_shapes()
    {
        using var pres = new Presentation("pres.pptx");
        var shapes = pres.Slide(1).Shapes;
        var shape1 = shapes.Shape("Shape 1");
        var shape2 = shapes.Shape("Shape 2");

        var group = shapes.Group([shape1, shape2]);
    }

    [Test]
    [Explicit]
    public void Add_Line_shape()
    {
        using var pres = new Presentation("some.pptx");
        var shapes = pres.Slide(1).Shapes;

        shapes.AddLine(100, 50, 100, 50);

        pres.Save();
    }

    [Test]
    [Explicit]
    public void Add_shape()
    {
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;

        shapes.AddShape(50, 60, 100, 200, Geometry.Rectangle, "Test");
    }
}
