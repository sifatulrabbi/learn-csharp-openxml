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
        PptxDataExtractor extractor = new(pptx);
        var extractedInfo = extractor.ExtractAll();
        Console.WriteLine(JsonSerializer.Serialize(extractedInfo, jsonSerializerOptions));

        // Test text modification
        PptxDataModifier modifier = new(pptx);
        modifier.ModifySlideTextTest();

        // AnalyzePresentationPropertiesStructure(pptx);

        pptx.Save();
        pptx.Dispose();
    }
}

internal class PptxDataExtractor
{
    private PresentationDocument pptx;

    public PptxDataExtractor(PresentationDocument pptx)
    {
        if (pptx.PresentationPart == null || pptx.PresentationPart.Presentation == null)
        {
            throw new MissingMemberException(
                nameof(PresentationDocument),
                "Presentation or PresentationPart is not found!"
            );
        }
        this.pptx = pptx;
    }

    public PresentationData ExtractAll()
    {
        // Extract presentation metadata
        var coreProps = pptx.PackageProperties;
        var title = coreProps.Title ?? "";
        var subject = coreProps.Subject ?? "";
        var description = coreProps.Description ?? "";
        var creator = coreProps.Creator ?? "";
        var created = coreProps.Created ?? DateTime.MinValue;
        var modified = coreProps.Modified ?? DateTime.MinValue;

        // Extract slide size
        var slideSize = new SlideSize
        {
            Width = pptx.PresentationPart!.Presentation!.SlideSize?.Cx?.Value ?? 0,
            Height = pptx.PresentationPart.Presentation.SlideSize?.Cy?.Value ?? 0,
            Type = pptx.PresentationPart.Presentation.SlideSize?.Type?.ToString() ?? "",
        };

        // Extract note size
        var noteSize = new NoteSize
        {
            Width = pptx.PresentationPart.Presentation.NotesSize?.Cx?.Value ?? 0,
            Height = pptx.PresentationPart.Presentation.NotesSize?.Cy?.Value ?? 0,
        };

        // Extract presentation-level theme (even when masters exist)
        var globalTheme =
            pptx.PresentationPart.ThemePart != null && pptx.PresentationPart.ThemePart.Theme != null
                ? ExtractThemeInfo(pptx.PresentationPart.ThemePart)
                : new Theme
                {
                    Dark1 = "",
                    Light1 = "",
                    Dark2 = "",
                    Light2 = "",
                    Accent1 = "",
                    Accent2 = "",
                    Accent3 = "",
                    Accent4 = "",
                    Accent5 = "",
                    Accent6 = "",
                    Hyperlink = "",
                    FollowedHyperlink = "",
                };

        // Extract slide masters
        var masterParts = new List<SlideMaster>();
        foreach (var slideMasterPart in pptx.PresentationPart.SlideMasterParts)
            if (slideMasterPart.SlideMaster != null)
                masterParts.Add(ExtractMasterInfo(slideMasterPart));

        // Extract all slides
        var slideInfos = new List<Slide>();
        var slideIds = pptx.PresentationPart.Presentation!.SlideIdList?.Elements<SlideId>();
        if (slideIds != null)
            foreach (var slideId in slideIds)
            {
                var relId = slideId.RelationshipId;
                if (relId != null)
                {
                    var slidePart = (SlidePart)pptx.PresentationPart.GetPartById(relId.Value!);
                    if (slidePart.Slide != null && slideId.Id != null)
                        slideInfos.Add(ExtractSlideInfo(slideId.Id.Value, slidePart));
                }
            }

        return new PresentationData
        {
            Title = title,
            Subject = subject,
            Description = description,
            Creator = creator,
            Created = created,
            Modified = modified,
            SlideSize = slideSize,
            NoteSize = noteSize,
            GlobalTheme = globalTheme,
            SlideMasters = masterParts,
            Slides = slideInfos,
        };
    }

    private static SlideMaster ExtractMasterInfo(SlideMasterPart masterPart)
    {
        var name = masterPart.SlideMaster?.CommonSlideData?.Name?.Value ?? "slide-master-part";

        // Extract layout IDs for slide masters
        var slideLayoutIdList = new List<string>();
        if (masterPart.SlideMaster?.SlideLayoutIdList != null)
        {
            foreach (var id in masterPart.SlideMaster.SlideLayoutIdList.Elements())
            {
                if (id != null && !string.IsNullOrEmpty(id.InnerText))
                    slideLayoutIdList.Add(id.InnerText);
            }
        }

        // Extract layouts for slide masters
        var layoutParts = new List<SlideLayout>();
        if (masterPart.SlideLayoutParts != null)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
                layoutParts.Add(ExtractLayoutInfo(layoutPart));
        }

        // Extract theme information
        var theme =
            masterPart.ThemePart != null
                ? ExtractThemeInfo(masterPart.ThemePart)
                : new Theme
                {
                    Dark1 = "",
                    Light1 = "",
                    Dark2 = "",
                    Light2 = "",
                    Accent1 = "",
                    Accent2 = "",
                    Accent3 = "",
                    Accent4 = "",
                    Accent5 = "",
                    Accent6 = "",
                    Hyperlink = "",
                    FollowedHyperlink = "",
                };

        return new SlideMaster
        {
            Name = name,
            SlideLayoutIds = slideLayoutIdList,
            SlideLayouts = layoutParts,
            Theme = theme,
        };
    }

    private static SlideLayout ExtractLayoutInfo(SlideLayoutPart layoutPart)
    {
        var layout = layoutPart.SlideLayout;
        return new SlideLayout
        {
            Name = layout?.CommonSlideData?.Name?.Value ?? "Default layout",
            TypeName = layout?.Type?.ToString() ?? "",
        };
    }

    private static Theme ExtractThemeInfo(ThemePart themePart)
    {
        var colorScheme = themePart.Theme.ThemeElements?.ColorScheme;
        if (colorScheme == null)
            return new Theme
            {
                Dark1 = "",
                Light1 = "",
                Dark2 = "",
                Light2 = "",
                Accent1 = "",
                Accent2 = "",
                Accent3 = "",
                Accent4 = "",
                Accent5 = "",
                Accent6 = "",
                Hyperlink = "",
                FollowedHyperlink = "",
            };

        return new Theme
        {
            Dark1 = colorScheme.Dark1Color?.RgbColorModelHex?.Val?.Value ?? "",
            Light1 = colorScheme.Light1Color?.RgbColorModelHex?.Val?.Value ?? "",
            Dark2 = colorScheme.Dark2Color?.RgbColorModelHex?.Val?.Value ?? "",
            Light2 = colorScheme.Light2Color?.RgbColorModelHex?.Val?.Value ?? "",
            Accent1 = colorScheme.Accent1Color?.RgbColorModelHex?.Val?.Value ?? "",
            Accent2 = colorScheme.Accent2Color?.RgbColorModelHex?.Val?.Value ?? "",
            Accent3 = colorScheme.Accent3Color?.RgbColorModelHex?.Val?.Value ?? "",
            Accent4 = colorScheme.Accent4Color?.RgbColorModelHex?.Val?.Value ?? "",
            Accent5 = colorScheme.Accent5Color?.RgbColorModelHex?.Val?.Value ?? "",
            Accent6 = colorScheme.Accent6Color?.RgbColorModelHex?.Val?.Value ?? "",
            Hyperlink = colorScheme.Hyperlink?.RgbColorModelHex?.Val?.Value ?? "",
            FollowedHyperlink =
                colorScheme.FollowedHyperlinkColor?.RgbColorModelHex?.Val?.Value ?? "",
        };
    }

    private static Slide ExtractSlideInfo(uint id, SlidePart slidePart)
    {
        PrintDescendentTree(slidePart);
        return new Slide
        {
            SlideId = id,
            LayoutName = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Name?.Value ?? "",
            Texts = slidePart.Slide.Descendants<D.Text>().Select(t => t.Text).ToList(),
        };
    }

    private static void PrintDescendentTree(SlidePart slidePart)
    {
        var shapeTrees = slidePart.Slide.Descendants<ShapeTree>();
        Console.WriteLine($"{slidePart.Slide.LocalName}, Found {shapeTrees.Count()} shape trees");
        foreach (var shapeTree in shapeTrees)
            PrintElement(shapeTree, 1);
    }

    private static void PrintElement(DocumentFormat.OpenXml.OpenXmlElement element, int level)
    {
        var indent = new string(' ', level * 2);
        Console.WriteLine($"{indent}{element.GetType().Name}");
        foreach (var child in element.Elements())
            PrintElement(child, level + 1);
    }
}

internal class PptxDataModifier
{
    private PresentationDocument pptx;

    public PptxDataModifier(PresentationDocument pptx)
    {
        if (pptx.PresentationPart == null || pptx.PresentationPart.Presentation == null)
        {
            throw new MissingMemberException(
                nameof(PresentationDocument),
                "Presentation or PresentationPart is not found!"
            );
        }
        this.pptx = pptx;
    }

    public void ModifySlideTextTest()
    {
        var textToFind = "Basic presentation";
        var newText = "Updated Basic Presentation";

        var slideToChange = pptx.PresentationPart!.SlideParts.FirstOrDefault(sp =>
            sp.Slide.Descendants<D.Text>().Any(t => t.Text == textToFind)
        );

        if (slideToChange != null)
        {
            var updated = TryUpdateSlideText(slideToChange, textToFind, newText);
            Console.WriteLine(
                updated
                    ? $"Successfully updated text from '{textToFind}' to '{newText}'"
                    : $"Failed to update text '{textToFind}'"
            );
        }
        else
        {
            Console.WriteLine($"No slide found containing text '{textToFind}'");
        }
    }

    private static bool TryUpdateSlideText(SlidePart slidePart, string oldText, string newText)
    {
        var texts = slidePart.Slide.Descendants<D.Text>().Where(t => t.Text == oldText).ToList();
        if (texts.Count == 0)
            return false;
        foreach (var t in texts)
            t.Text = newText;
        slidePart.Slide.Save();
        return true;
    }
}

public class PresentationData
{
    public required string Title { get; set; }
    public required string Subject { get; set; }
    public required string Description { get; set; }
    public string Creator { get; set; } = "James SEQUESTO";
    public DateTime Created { get; set; } = DateTime.UtcNow;
    public DateTime Modified { get; set; } = DateTime.UtcNow;
    public required SlideSize SlideSize { get; set; }
    public required NoteSize NoteSize { get; set; }
    public required Theme GlobalTheme { get; set; }
    public required List<SlideMaster> SlideMasters { get; set; }
    public required List<Slide> Slides { get; set; }
}

public class SlideSize
{
    public required long Width { get; set; }
    public required long Height { get; set; }
    public required string Type { get; set; }
}

public class NoteSize
{
    public required long Width { get; set; }
    public required long Height { get; set; }
}

public class Theme
{
    public required string Dark1 { get; set; }
    public required string Light1 { get; set; }
    public required string Dark2 { get; set; }
    public required string Light2 { get; set; }
    public required string Accent1 { get; set; }
    public required string Accent2 { get; set; }
    public required string Accent3 { get; set; }
    public required string Accent4 { get; set; }
    public required string Accent5 { get; set; }
    public required string Accent6 { get; set; }
    public required string Hyperlink { get; set; }
    public required string FollowedHyperlink { get; set; }
}

public class SlideMaster
{
    public required string Name { get; set; }
    public required List<string> SlideLayoutIds { get; set; }
    public required List<SlideLayout> SlideLayouts { get; set; }
    public required Theme Theme { get; set; }
}

public class SlideLayout
{
    public required string Name { get; set; }
    public required string TypeName { get; set; }
}

public class Slide
{
    public required uint SlideId { get; set; }
    public required string LayoutName { get; set; }
    public required List<string> Texts { get; set; }
}
