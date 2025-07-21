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
        var slideSize = new PptxSlideSize
        {
            Width = pptx.PresentationPart!.Presentation!.SlideSize?.Cx?.Value ?? 0,
            Height = pptx.PresentationPart.Presentation.SlideSize?.Cy?.Value ?? 0,
            Type = pptx.PresentationPart.Presentation.SlideSize?.Type?.ToString() ?? "",
        };

        // Extract note size
        var noteSize = new PptxNoteSize
        {
            Width = pptx.PresentationPart.Presentation.NotesSize?.Cx?.Value ?? 0,
            Height = pptx.PresentationPart.Presentation.NotesSize?.Cy?.Value ?? 0,
        };

        // Extract presentation-level theme (even when masters exist)
        var globalTheme = ExtractThemeInfo(pptx.PresentationPart.ThemePart);

        // Extract slide masters
        var masterParts = new List<PptxSlideMaster>();
        foreach (var slideMasterPart in pptx.PresentationPart.SlideMasterParts)
            if (slideMasterPart.SlideMaster != null)
                masterParts.Add(ExtractMasterInfo(slideMasterPart));

        // Extract all slides
        var slideInfos = new List<PptxSlide>();
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

    private static PptxSlideMaster ExtractMasterInfo(SlideMasterPart masterPart)
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
        var layoutParts = new List<PptxSlideLayout>();
        if (masterPart.SlideLayoutParts != null)
        {
            foreach (var layoutPart in masterPart.SlideLayoutParts)
                layoutParts.Add(ExtractLayoutInfo(layoutPart));
        }

        // Extract theme information
        var theme = ExtractThemeInfo(masterPart.ThemePart);

        return new PptxSlideMaster
        {
            Name = name,
            SlideLayoutIds = slideLayoutIdList,
            SlideLayouts = layoutParts,
            Theme = theme,
        };
    }

    private static PptxSlideLayout ExtractLayoutInfo(SlideLayoutPart layoutPart)
    {
        var layout = layoutPart.SlideLayout;
        return new PptxSlideLayout
        {
            Name = layout?.CommonSlideData?.Name?.Value ?? "Default layout",
            TypeName = layout?.Type?.ToString() ?? "",
        };
    }

    private static PptxTheme ExtractThemeInfo(ThemePart? themePart)
    {
        if (themePart == null || themePart.Theme.ThemeElements?.ColorScheme == null)
        {
            return new PptxTheme
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

        return new PptxTheme
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

    private static PptxSlide ExtractSlideInfo(uint id, SlidePart slidePart)
    {
        PrintDescendentTree(slidePart);

        List<PptxSlideContent> contents = [];
        IEnumerable<ShapeTree> shapeTrees = slidePart.Slide.Descendants<ShapeTree>();
        IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> contentElements = shapeTrees.Any()
            ? shapeTrees.ElementAt(0).Elements()
            : [];

        foreach (var element in contentElements)
        {
            switch (element)
            {
                case Shape shape:
                    if (shape.TextBody == null)
                        break;
                    var paragraphTexts = shape
                        .TextBody.Elements<D.Paragraph>()
                        .Select(p => string.Concat(p.Elements<D.Run>().Select(r => r.Text?.Text)))
                        .Where(t => !string.IsNullOrEmpty(t));
                    foreach (var txt in paragraphTexts)
                        contents.Add(
                            new() { ContentType = PptxSlideContentTypes.Text, Text = txt }
                        );
                    break;

                case Picture picture:
                    var imageUrl = GetImageUrl(picture, slidePart);
                    if (!string.IsNullOrEmpty(imageUrl))
                        contents.Add(
                            new() { ContentType = PptxSlideContentTypes.Image, ImageUrl = imageUrl }
                        );
                    break;

                case GraphicFrame frame:
                    var graphicFrame = ExtractGraphicFrame(frame, slidePart);
                    if (graphicFrame != null)
                        contents.Add(graphicFrame);
                    break;

                case ContentPart:
                case GroupShape:
                case ConnectionShape:
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

        return new PptxSlide
        {
            SlideId = id,
            LayoutName = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Name?.Value ?? "",
            Contents = contents,
        };
    }

    private static string GetImageUrl(Picture picture, SlidePart slidePart)
    {
        var embedId = picture.BlipFill?.Blip?.Embed?.Value;
        if (embedId == null)
            return "";
        var imagePart = (ImagePart)slidePart.GetPartById(embedId);
        return imagePart.Uri.ToString();
    }

    private static PptxSlideContent? ExtractGraphicFrame(GraphicFrame frame, SlidePart _)
    {
        var table = frame.Graphic?.GraphicData?.GetFirstChild<D.Table>();
        if (table != null)
        {
            PptxTableContent tableContent = new();
            foreach (var row in table.Elements<D.TableRow>())
            {
                PptxTableRow tableRow = new();
                foreach (var cell in row.Elements<D.TableCell>())
                {
                    var cellText = "";
                    if (cell.TextBody != null)
                    {
                        cellText = string.Join(
                            Environment.NewLine,
                            cell.TextBody.Elements<D.Paragraph>()
                                .Select(p =>
                                    string.Concat(p.Elements<D.Run>().Select(r => r.Text?.Text))
                                )
                        );
                    }
                    tableRow.Cells.Add(
                        new() { ContentType = PptxSlideContentTypes.Text, Text = cellText }
                    );
                }
                tableContent.Rows.Add(tableRow);
            }
            return new() { ContentType = PptxSlideContentTypes.Table, Table = tableContent };
        }
        return null;
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
