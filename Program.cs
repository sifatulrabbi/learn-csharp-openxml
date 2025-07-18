using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace CsharpOpenXML;

public class Program
{
    private static readonly string testPptxFile =
        "./test-files/universal-presentation-fdsfa3432412f.pptx";

    public static void Main()
    {
        TestingOutOpenXML testingOutOpenXML = new(testPptxFile);
        testingOutOpenXML.TestOutPptxFile();
    }
}

internal class TestingOutOpenXML(string filePath)
{
    private static readonly JsonSerializerOptions jsonSerializerOptions = new()
    {
        WriteIndented = true,
    };

    public void TestOutPptxFile()
    {
        using PresentationDocument pptx = PresentationDocument.Open(filePath, true);
        if (pptx.PresentationPart == null || pptx.PresentationPart.Presentation == null)
        {
            throw new MissingMemberException(
                nameof(PresentationDocument),
                "Presentation or PresentationPart is not found!"
            );
        }
        // AnalyzePresentationPropertiesStructure(pptx);

        var extractedInfo = ExtractAllSlideMasterAndThemeInfo(pptx);
        Console.WriteLine(JsonSerializer.Serialize(extractedInfo, jsonSerializerOptions));

        pptx.Save();
        pptx.Dispose();
    }

    private static Dictionary<string, object> ExtractAllSlideMasterAndThemeInfo(
        PresentationDocument pptx
    )
    {
        if (pptx.PresentationPart == null || pptx.PresentationPart.Presentation == null)
        {
            throw new MissingMemberException(
                nameof(PresentationDocument),
                "Presentation or PresentationPart is not found!"
            );
        }

        Dictionary<string, object> result = new();

        // Extract presentation-level properties
        if (pptx.PresentationPart.Presentation.SlideSize != null)
        {
            var slideSize = pptx.PresentationPart.Presentation.SlideSize;
            result["slideSize"] = new Dictionary<string, object>
            {
                ["width"] = slideSize.Cx?.Value ?? 0,
                ["height"] = slideSize.Cy?.Value ?? 0,
                ["type"] = slideSize.Type?.ToString() ?? "",
            };
        }
        if (pptx.PresentationPart.Presentation?.NotesSize != null)
        {
            var noteSize = pptx.PresentationPart.Presentation.NotesSize;
            result["noteSize"] = new Dictionary<string, object>
            {
                ["width"] = noteSize.Cx?.Value ?? 0,
                ["height"] = noteSize.Cy?.Value ?? 0,
            };
        }

        // Extract presentation-level theme (even when masters exist)
        if (
            pptx.PresentationPart.ThemePart != null
            && pptx.PresentationPart.ThemePart.Theme != null
        )
            result["globalTheme"] = ExtractThemeInfo(pptx.PresentationPart.ThemePart);

        // Extract slide masters
        List<object> masterParts = [];
        if (pptx.PresentationPart?.SlideMasterParts != null)
            foreach (var slideMasterPart in pptx.PresentationPart.SlideMasterParts)
                if (slideMasterPart.SlideMaster != null)
                    masterParts.Add(ExtractMasterInfo(slideMasterPart));
        result["slideMasters"] = masterParts;

        // Extract all slides
        var slideInfos = new List<object>();
        var slideIds = pptx.PresentationPart!.Presentation!.SlideIdList?.Elements<SlideId>();
        if (slideIds != null)
        {
            foreach (var slideId in slideIds)
            {
                var relId = slideId.RelationshipId;
                if (relId != null)
                {
                    var slidePart = (SlidePart)pptx.PresentationPart.GetPartById(relId!);
                    if (slidePart.Slide != null && slideId.Id != null)
                        slideInfos.Add(ExtractSlideInfo(slideId.Id.Value, slidePart));
                }
            }
        }
        result["slides"] = slideInfos;

        return result;
    }

    private static Dictionary<string, object> ExtractMasterInfo(SlideMasterPart masterPart)
    {
        Dictionary<string, object> slideMasterInfo = new();
        if (masterPart.SlideMaster != null)
        {
            slideMasterInfo["name"] =
                masterPart.SlideMaster.CommonSlideData?.Name?.Value ?? "slide-master-part";
            // Extract layout count for slide masters
            if (masterPart.SlideMaster.SlideLayoutIdList != null)
            {
                var slideLayoutIdList = new List<object>();
                foreach (var id in masterPart.SlideMaster.SlideLayoutIdList.Elements())
                {
                    if (id != null && !string.IsNullOrEmpty(id.InnerText))
                        slideLayoutIdList.Add(id.InnerText);
                }
                slideMasterInfo["slideLayoutIds"] = slideLayoutIdList;
            }
        }
        // Extract layouts for slide masters
        if (masterPart.SlideLayoutParts != null)
        {
            var layoutParts = new List<object>();
            foreach (var layoutPart in masterPart.SlideLayoutParts)
                layoutParts.Add(ExtractLayoutInfo(layoutPart));
            slideMasterInfo["slideLayouts"] = layoutParts;
        }
        // Extract theme information
        if (masterPart.ThemePart != null)
            slideMasterInfo["theme"] = ExtractThemeInfo(masterPart.ThemePart);

        return slideMasterInfo;
    }

    private static Dictionary<string, object> ExtractLayoutInfo(SlideLayoutPart layoutPart)
    {
        var layoutInfo = new Dictionary<string, object>();
        var layout = layoutPart.SlideLayout;
        if (layout == null)
            return layoutInfo;
        layoutInfo["name"] = layout.CommonSlideData?.Name?.Value ?? "Default layout";
        layoutInfo["typeName"] = layout.Type?.ToString() ?? "";
        return layoutInfo;
    }

    private static Dictionary<string, object> ExtractThemeInfo(ThemePart themePart)
    {
        var colorScheme = themePart.Theme.ThemeElements?.ColorScheme;
        if (colorScheme == null)
            return new Dictionary<string, object>();
        return new Dictionary<string, object>
        {
            ["dark1"] = colorScheme.Dark1Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["light1"] = colorScheme.Light1Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["dark2"] = colorScheme.Dark2Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["light2"] = colorScheme.Light2Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["accent1"] = colorScheme.Accent1Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["accent2"] = colorScheme.Accent2Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["accent3"] = colorScheme.Accent3Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["accent4"] = colorScheme.Accent4Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["accent5"] = colorScheme.Accent5Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["accent6"] = colorScheme.Accent6Color?.RgbColorModelHex?.Val?.Value ?? "",
            ["hyperlink"] = colorScheme.Hyperlink?.RgbColorModelHex?.Val?.Value ?? "",
            ["followedHyperlink"] =
                colorScheme.FollowedHyperlinkColor?.RgbColorModelHex?.Val?.Value ?? "",
        };
    }

    private static Dictionary<string, object> ExtractSlideInfo(uint id, SlidePart slidePart)
    {
        PrintDescendentTree(slidePart);
        return new Dictionary<string, object>
        {
            ["slideId"] = id,
            ["layoutName"] =
                slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Name?.Value ?? "",
            ["texts"] = slidePart.Slide.Descendants<D.Text>().Select(t => t.Text).ToList(),
        };
    }

    private static void PrintDescendentTree(SlidePart slidePart)
    {
        PrintElement(slidePart.Slide, 0);
    }

    private static void PrintElement(DocumentFormat.OpenXml.OpenXmlElement element, int level)
    {
        var indent = new string(' ', level * 4);
        Console.WriteLine($"{indent}{element.GetType().Name}");
        foreach (var child in element.Elements())
            PrintElement(child, level + 1);
    }

    private static void AnalyzePresentationPropertiesStructure(PresentationDocument pptx)
    {
        Console.WriteLine("--- Old values ---");
        var coreProps = pptx.PackageProperties;
        Console.WriteLine("Title:           " + coreProps.Title);
        Console.WriteLine("Subject:         " + coreProps.Subject);
        Console.WriteLine("Creator:         " + coreProps.Creator);
        Console.WriteLine("Keywords:        " + coreProps.Keywords);
        Console.WriteLine("Description:     " + coreProps.Description);
        Console.WriteLine("Last Modified By:" + coreProps.LastModifiedBy);
        Console.WriteLine("Revision:        " + coreProps.Revision);
        Console.WriteLine("Created:         " + coreProps.Created);
        Console.WriteLine("Modified:        " + coreProps.Modified);

        var extPart = pptx.ExtendedFilePropertiesPart;
        if (extPart != null)
        {
            var props = extPart.Properties;
            Console.WriteLine();
            Console.WriteLine("Application:         " + props.Application?.Text);
            Console.WriteLine("Company:             " + props.Company?.Text);
            Console.WriteLine("Manager:             " + props.Manager?.Text);
            Console.WriteLine("Total Editing Time:  " + props.TotalTime?.Text + " minutes");
            Console.WriteLine("Presentation Format: " + props.PresentationFormat?.Text);
        }
        else
        {
            Console.WriteLine("No extended properties found.");
        }

        Console.WriteLine("--- New values ---");

        coreProps.Modified = DateTime.Now;
        Console.WriteLine("+ Modified:        " + coreProps.Modified);
    }
}

// public class PresentationData
// {
//     public required string Title { get; set; }
//     public required string Subject { get; set; }
//     public required string Description { get; set; }
//     public string Creator { get; set; } = "James";
//     public DateTime Created { get; set; } = DateTime.UtcNow;
//     public DateTime Modified { get; set; } = DateTime.UtcNow;

//     public required PresentationTheme Theme { get; set; }
// }
