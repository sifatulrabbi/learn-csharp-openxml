using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace CsharpOpenXML;

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
        var theme = ExtractThemeInfo(masterPart.ThemePart);

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

    private static Theme ExtractThemeInfo(ThemePart? themePart)
    {
        if (themePart == null || themePart.Theme.ThemeElements?.ColorScheme == null)
        {
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
        }
        var colorScheme = themePart.Theme.ThemeElements!.ColorScheme!;

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
        IEnumerable<ShapeTree> shapeTrees = slidePart.Slide.Descendants<ShapeTree>();
        IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> contentElements = shapeTrees.Any()
            ? shapeTrees.ElementAt(0).Elements()
            : [];

        foreach (var element in contentElements)
        {
            switch (element)
            {
                case Shape shape:
                    // TODO:
                    break;
                case Picture picture:
                    // TODO:
                    break;
                case GraphicFrame frame:
                    // TODO:
                    break;
                case GroupShape groupShape:
                    // TODO:
                    break;
                case ConnectionShape connector:
                    // TODO:
                    break;
                case ContentPart contentPart:
                    break;
                case NonVisualGroupShapeProperties:
                case GroupShapeProperties:
                case null:
                    break;
                default:
                    Console.WriteLine(
                        $"Unknown content type encountered: {element.GetType().Name}"
                    );
                    break;
            }
        }

        return new Slide
        {
            SlideId = id,
            LayoutName = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Name?.Value ?? "",
            Texts = [],
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
