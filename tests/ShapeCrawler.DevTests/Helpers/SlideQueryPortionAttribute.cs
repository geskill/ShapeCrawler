using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.DevTests.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlideQueryPortionAttribute(
    string pptxName,
    int slideNumber,
    string shapeName,
    int paragraphNumber,
    int portionNumber)
    : Attribute, ITestBuilder
{
    private readonly object expectedResult;
    private readonly int shapeId;

    public SlideQueryPortionAttribute(string pptxName, int slideNumber, int shapeId, int paragraphNumber,
        int portionNumber) : this(pptxName, slideNumber, null, paragraphNumber, portionNumber)
    {
        this.shapeId = shapeId;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new Presentation(SCTest.TestAsset(pptxName));
        var portionQuery = shapeName == null
            ? new TestSlidePortionQuery(slideNumber, shapeId, paragraphNumber, portionNumber)
            : new TestSlidePortionQuery(slideNumber, shapeName, paragraphNumber, portionNumber);

        var parameters = new TestCaseParameters(new object[] { pres, portionQuery });

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}
