using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;

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
        AnalyzePresentationPropertiesStructure(pptx);

        var extractedInfo = ExtractAllSlideMasterAndThemeInfo(pptx);
        Console.WriteLine(JsonSerializer.Serialize(extractedInfo, jsonSerializerOptions));

        pptx.Save();
        pptx.Dispose();
    }

    private static Dictionary<string, object> ExtractAllSlideMasterAndThemeInfo(
        PresentationDocument pptx
    )
    {
        Dictionary<string, object> result = new();
        List<object> masterParts = [];

        if (pptx.PresentationPart?.SlideMasterParts != null)
        {
            foreach (var slideMasterPart in pptx.PresentationPart.SlideMasterParts)
            {
                var masterInfo = new Dictionary<string, object>
                {
                    ["type"] = "slideMaster",
                    ["uri"] = slideMasterPart.Uri?.ToString() ?? "",
                    ["relationshipType"] = slideMasterPart.RelationshipType,
                    ["contentType"] = slideMasterPart.ContentType,
                    ["slideMasterName"] =
                        slideMasterPart.SlideMaster?.CommonSlideData?.Name?.Value
                        ?? (
                            slideMasterPart.Uri != null
                                ? Path.GetFileNameWithoutExtension(slideMasterPart.Uri.ToString())
                                : "default-slide-master"
                        ),
                    ["preserveElements"] = slideMasterPart.SlideMaster?.Preserve?.Value ?? false,
                    ["colorMap"] =
                        slideMasterPart
                            .SlideMaster?.ColorMap?.Elements()
                            .ToDictionary(e => e.LocalName, e => e.InnerText)
                        ?? new Dictionary<string, string>(),
                    ["slideLayoutIdList"] =
                        slideMasterPart.SlideMaster?.SlideLayoutIdList?.Elements()?.Count() ?? 0,
                    ["layouts"] = new List<object>(),
                };

                if (slideMasterPart.SlideLayoutParts != null)
                {
                    var layouts = new List<object>();
                    foreach (var layoutPart in slideMasterPart.SlideLayoutParts)
                    {
                        var layoutInfo = new Dictionary<string, object>
                        {
                            ["type"] = "slideLayout",
                            ["uri"] = layoutPart.Uri?.ToString() ?? "",
                            ["relationshipType"] = layoutPart.RelationshipType,
                            ["contentType"] = layoutPart.ContentType,
                            ["layoutName"] =
                                layoutPart.SlideLayout?.CommonSlideData?.Name?.Value
                                ?? "default-layout",
                            ["layoutType"] =
                                layoutPart.SlideLayout?.Type?.ToString() ?? "default-layout",
                            ["preserveElements"] = layoutPart.SlideLayout?.Preserve?.Value ?? true,
                            ["showMasterShapes"] =
                                layoutPart.SlideLayout?.ShowMasterShapes?.Value ?? true,
                            ["showMasterPlaceholderAnimations"] =
                                layoutPart.SlideLayout?.ShowMasterPlaceholderAnimations?.Value
                                ?? true,
                            ["placeholderCount"] =
                                layoutPart
                                    .SlideLayout?.CommonSlideData?.ShapeTree?.Elements()
                                    ?.Count(e => e.LocalName == "sp") ?? 0,
                        };
                        layouts.Add(layoutInfo);
                    }
                    masterInfo["layouts"] = layouts;
                }

                if (slideMasterPart.ThemePart != null)
                {
                    var theme = slideMasterPart.ThemePart.Theme;
                    masterInfo["theme"] = new Dictionary<string, object>
                    {
                        ["type"] = "theme",
                        ["uri"] = slideMasterPart.ThemePart.Uri?.ToString() ?? "",
                        ["relationshipType"] = slideMasterPart.ThemePart.RelationshipType,
                        ["contentType"] = slideMasterPart.ThemePart.ContentType,
                        ["themeName"] = theme?.Name?.Value ?? "",
                        ["colorScheme"] = new Dictionary<string, object>
                        {
                            ["name"] = theme?.ThemeElements?.ColorScheme?.Name?.Value ?? "",
                            ["dark1"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Dark1Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["light1"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Light1Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["dark2"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Dark2Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["light2"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Light2Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["accent1"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Accent1Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["accent2"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Accent2Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["accent3"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Accent3Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["accent4"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Accent4Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["accent5"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Accent5Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["accent6"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Accent6Color
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["hyperlink"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.Hyperlink
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                            ["followedHyperlink"] =
                                theme
                                    ?.ThemeElements
                                    ?.ColorScheme
                                    ?.FollowedHyperlinkColor
                                    ?.RgbColorModelHex
                                    ?.Val
                                    ?.Value ?? "",
                        },
                        ["fontScheme"] = new Dictionary<string, object>
                        {
                            ["name"] = theme?.ThemeElements?.FontScheme?.Name?.Value ?? "",
                            ["majorFontLatin"] =
                                theme
                                    ?.ThemeElements
                                    ?.FontScheme
                                    ?.MajorFont
                                    ?.LatinFont
                                    ?.Typeface
                                    ?.Value ?? "",
                            ["minorFontLatin"] =
                                theme
                                    ?.ThemeElements
                                    ?.FontScheme
                                    ?.MinorFont
                                    ?.LatinFont
                                    ?.Typeface
                                    ?.Value ?? "",
                        },
                        ["formatScheme"] = new Dictionary<string, object>
                        {
                            ["name"] = theme?.ThemeElements?.FormatScheme?.Name?.Value ?? "",
                            ["fillStyleListCount"] =
                                theme
                                    ?.ThemeElements?.FormatScheme?.FillStyleList?.Elements()
                                    ?.Count() ?? 0,
                            ["lineStyleListCount"] =
                                theme
                                    ?.ThemeElements?.FormatScheme?.LineStyleList?.Elements()
                                    ?.Count() ?? 0,
                            ["effectStyleListCount"] =
                                theme
                                    ?.ThemeElements?.FormatScheme?.EffectStyleList?.Elements()
                                    ?.Count() ?? 0,
                            ["backgroundFillStyleListCount"] =
                                theme
                                    ?.ThemeElements?.FormatScheme?.BackgroundFillStyleList?.Elements()
                                    ?.Count() ?? 0,
                        },
                    };
                }

                masterParts.Add(masterInfo);
            }
        }

        if (
            pptx.PresentationPart?.ThemePart != null
            && !pptx.PresentationPart.SlideMasterParts.Any()
        )
        {
            var theme = pptx.PresentationPart.ThemePart.Theme;
            masterParts.Add(
                new Dictionary<string, object>
                {
                    ["type"] = "presentationTheme",
                    ["uri"] = pptx.PresentationPart.ThemePart.Uri?.ToString() ?? "",
                    ["relationshipType"] = pptx.PresentationPart.ThemePart.RelationshipType,
                    ["contentType"] = pptx.PresentationPart.ThemePart.ContentType,
                    ["themeName"] = theme?.Name?.Value ?? "",
                    ["colorScheme"] = new Dictionary<string, object>
                    {
                        ["name"] = theme?.ThemeElements?.ColorScheme?.Name?.Value ?? "",
                        ["dark1"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Dark1Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["light1"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Light1Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["dark2"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Dark2Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["light2"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Light2Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["accent1"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Accent1Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["accent2"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Accent2Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["accent3"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Accent3Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["accent4"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Accent4Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["accent5"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Accent5Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["accent6"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Accent6Color
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["hyperlink"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.Hyperlink
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                        ["followedHyperlink"] =
                            theme
                                ?.ThemeElements
                                ?.ColorScheme
                                ?.FollowedHyperlinkColor
                                ?.RgbColorModelHex
                                ?.Val
                                ?.Value ?? "",
                    },
                    ["fontScheme"] = new Dictionary<string, object>
                    {
                        ["name"] = theme?.ThemeElements?.FontScheme?.Name?.Value ?? "",
                        ["majorFontLatin"] =
                            theme?.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value
                            ?? "",
                        ["minorFontLatin"] =
                            theme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value
                            ?? "",
                    },
                }
            );
        }

        result["masterParts"] = masterParts;
        return result;
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

public class PresentationData
{
    public required string Title { get; set; }
    public required string Subject { get; set; }
    public required string Description { get; set; }
    public string Creator { get; set; } = "James";
    public DateTime Created { get; set; } = DateTime.UtcNow;
    public DateTime Modified { get; set; } = DateTime.UtcNow;

    public required PresentationTheme Theme { get; set; }
}

public class PresentationTheme
{
    /// <summary>
    /// Name or ID of the theme.
    /// </summary>
    public required string Name { get; set; }

    /// <summary>
    /// Color scheme (e.g. "accent1" -> "#FF0000").
    /// </summary>
    public required Dictionary<string, string> ColorScheme { get; set; }

    /// <summary>
    /// Font scheme (e.g. "heading" -> "Calibri", "body" -> "Arial").
    /// </summary>
    public required Dictionary<string, string> FontScheme { get; set; }
}

public class SlidePart
{
    /// <summary>
    /// Unique identifier for this slide.
    /// </summary>
    public required string Id { get; set; }

    /// <summary>
    /// Reference to the layout this slide uses.
    /// </summary>
    public required SlideLayoutPart Layout { get; set; }

    /// <summary>
    /// Raw slide content (e.g. xml or serializable shape tree).
    /// </summary>
    public required string Content { get; set; }
}

public class SlideLayoutPart
{
    /// <summary>
    /// Unique identifier for this layout.
    /// </summary>
    public required string Id { get; set; }

    /// <summary>
    /// Which master this layout belongs to.
    /// </summary>
    public required SlideMasterPart Master { get; set; }

    /// <summary>
    /// Any placeholders defined on this layout.
    /// </summary>
    public required List<string> PlaceholderTypes { get; set; }
}

public class SlideMasterPart
{
    /// <summary>
    /// Unique identifier for this master.
    /// </summary>
    public required string Id { get; set; }

    /// <summary>
    /// All layouts under this master.
    /// </summary>
    public required List<SlideLayoutPart> Layouts { get; set; }

    /// <summary>
    /// The theme applied to this master.
    /// </summary>
    public required ThemePart Theme { get; set; }
}
