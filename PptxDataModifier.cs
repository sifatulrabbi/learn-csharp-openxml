using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace CsharpOpenXML;

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

    public void UpdateFromJsonFile(string jsonFilePath)
    {
        var presentationData = LoadPresentationDataFromJson(jsonFilePath);
        UpdatePresentationMetadata(presentationData);
        UpdateSlideTexts(presentationData);
    }

    private PresentationData LoadPresentationDataFromJson(string jsonFilePath)
    {
        var jsonContent = File.ReadAllText(jsonFilePath);
        var presentationData = JsonSerializer.Deserialize<PresentationData>(jsonContent);
        if (presentationData == null)
        {
            throw new InvalidOperationException("Failed to deserialize JSON to PresentationData");
        }
        return presentationData;
    }

    private void UpdatePresentationMetadata(PresentationData data)
    {
        var coreProps = pptx.PackageProperties;
        if (!string.IsNullOrEmpty(data.Title))
            coreProps.Title = data.Title;
        if (!string.IsNullOrEmpty(data.Subject))
            coreProps.Subject = data.Subject;
        if (!string.IsNullOrEmpty(data.Description))
            coreProps.Description = data.Description;
        if (!string.IsNullOrEmpty(data.Creator))
            coreProps.Creator = data.Creator;
        if (data.Created != DateTime.MinValue)
            coreProps.Created = data.Created;
        if (data.Modified != DateTime.MinValue)
            coreProps.Modified = data.Modified;
    }

    private void UpdateSlideTexts(PresentationData data)
    {
        var slideParts = pptx.PresentationPart!.SlideParts.ToList();
        var slideIds = pptx
            .PresentationPart.Presentation!.SlideIdList?.Elements<SlideId>()
            .ToList();

        if (slideIds == null)
            return;

        foreach (var slideData in data.Slides)
        {
            var slideId = slideIds.FirstOrDefault(s => s.Id?.Value == slideData.SlideId);
            if (slideId?.RelationshipId?.Value == null)
                continue;

            var slidePart = (SlidePart)
                pptx.PresentationPart.GetPartById(slideId.RelationshipId.Value);
            UpdateSlideText(slidePart, slideData.Contents);
        }
    }

    private static void UpdateSlideText(SlidePart slidePart, List<PptxSlideContent> newTexts)
    {
        var textElements = slidePart.Slide.Descendants<D.Text>().ToList();
        for (int i = 0; i < Math.Min(textElements.Count, newTexts.Count); i++)
        {
            if (string.IsNullOrEmpty(newTexts[i].Text))
                textElements[i].Text = newTexts[i].Text!;
            // TODO: implement image insertion.
            // else if (string.IsNullOrEmpty(newTexts[i].ImageBase64))
            //     textElements[i].Text = newTexts[i].ImageBase64!;
            // else if (string.IsNullOrEmpty(newTexts[i].ImageUrl))
            //     textElements[i].Text = newTexts[i].ImageUrl!;
        }
        slidePart.Slide.Save();
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
