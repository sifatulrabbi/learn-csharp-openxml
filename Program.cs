using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace CsharpOpenXML;

public class Program
{
    private static readonly string testPptxFile = "./test-files/universal-presentation-fdsfa3432412f.pptx";

    public static void Main()
    {
        TestingOutOpenXML testingOutOpenXML = new(testPptxFile);
        testingOutOpenXML.TestOutPptxFile();
    }
}

internal class TestingOutOpenXML(string filePath)
{
    public void TestOutPptxFile()
    {
        Console.WriteLine("Creating a PPTX via OpenXML");
        PresentationDocument pptx = PresentationDocument.Open(filePath, false);
        if (pptx.PresentationPart == null || pptx.PresentationPart.Presentation == null)
        {
            throw new MissingMemberException(nameof(PresentationDocument), "Presentation or PresentationPart is not found!");
        }
        List<string> slideXmls = GetAllSlides(pptx);
        Console.WriteLine($"file type {pptx.GetType().Name}");
    }

    private static List<string> GetAllSlides(PresentationDocument pptx)
    {
        List<string> xmlSlides = ExtractSlidePartXml(pptx.PresentationPart!.SlideParts);
        Console.WriteLine($"Total slides found: {pptx.PresentationPart!.SlideParts.Count()}");
        Console.WriteLine($"Total slides in the XML: {xmlSlides.Count}");
        return xmlSlides;
    }

    private static List<string> ExtractSlidePartXml(IEnumerable<SlidePart> slideParts)
    {
        List<string> xmlSlides = [];
        foreach (SlidePart slidePart in slideParts)
        {
            xmlSlides.Add(slidePart.Slide.OuterXml);
            if (slidePart.SlideParts.Any()) { xmlSlides.AddRange(ExtractSlidePartXml(slidePart.SlideParts)); }
        }
        return xmlSlides;
    }

    private static void ExtractThemeXml(PresentationDocument pptx)
    {
    }
}
