namespace CsharpOpenXML;

public class PresentationData
{
    public required string Title { get; set; }
    public required string Subject { get; set; }
    public required string Description { get; set; }
    public string Creator { get; set; } = "James SEQUESTO";
    public DateTime Created { get; set; } = DateTime.UtcNow;
    public DateTime Modified { get; set; } = DateTime.UtcNow;
    public required PptxSlideSize SlideSize { get; set; }
    public required PptxNoteSize NoteSize { get; set; }
    public required PptxTheme GlobalTheme { get; set; }
    public required List<PptxSlideMaster> SlideMasters { get; set; }
    public required List<PptxSlide> Slides { get; set; }
}

public class PptxSlideSize
{
    public required long Width { get; set; }
    public required long Height { get; set; }
    public required string Type { get; set; }
}

public class PptxNoteSize
{
    public required long Width { get; set; }
    public required long Height { get; set; }
}

public class PptxTheme
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

public class PptxSlideMaster
{
    public required string Name { get; set; }
    public required List<string> SlideLayoutIds { get; set; }
    public required List<PptxSlideLayout> SlideLayouts { get; set; }
    public required PptxTheme Theme { get; set; }
}

public class PptxSlideLayout
{
    public required string Name { get; set; }
    public required string TypeName { get; set; }
}

public class PptxSlide
{
    public required uint SlideId { get; set; }
    public required string LayoutName { get; set; }
    public required List<PptxSlideContent> Contents { get; set; }
}

public class PptxSlideContent
{
    public required string ContentType { get; set; }
    public string? Text { get; set; }
    public string? ImageUrl { get; set; }
    public PptxTableContent? Table { get; set; }
}

public class PptxSlideContentTypes
{
    public const string Text = "text";
    public const string Image = "image";
    public const string Table = "table";
    // public const string GraphicFrame = "graphicFrame";
}

public class PptxTableContent
{
    public List<PptxTableRow> Rows { get; set; } = [];
}

public class PptxTableRow
{
    public List<PptxTableCell> Cells { get; set; } = [];
}

public class PptxTableCell
{
    public required string ContentType { get; set; }
    public string? Text { get; set; }
}
