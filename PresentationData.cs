namespace CsharpOpenXML;

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
    public required List<SlideContent> Contents { get; set; }
}

public class SlideContent
{
    public required string ContentType { get; set; }
    public string? Text { get; set; }
    public string? ImageUrl { get; set; }
    public string? ImageBase64 { get; set; }
}

public class SlideContentTypes
{
    public const string Text = "text";
    public const string Image = "image";
    public const string GraphicFrame = "graphicFrame";
}
