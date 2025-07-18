using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;

namespace CsharpOpenXML;

public class Program
{
    private static readonly string testPptxFile =
        "./test-files/universal-presentation-fdsfa3432412f.pptx";

    public static void Main(string[] args)
    {
        var updateMode = args.Length > 0 && args[0] == "update";
        TestingOutOpenXML testingOutOpenXML = new(testPptxFile);
        testingOutOpenXML.TestOutPptxFile(updateMode);
    }
}

internal class TestingOutOpenXML(string filePath)
{
    private static readonly JsonSerializerOptions jsonSerializerOptions = new()
    {
        WriteIndented = true,
    };

    public void TestOutPptxFile(bool updateFile)
    {
        using PresentationDocument pptx = PresentationDocument.Open(filePath, updateFile);
        if (updateFile)
        {
            PptxDataModifier modifier = new(pptx);
            modifier.UpdateFromJsonFile("./test-files/update.json");
            pptx.Save();
        }
        else
        {
            PptxDataExtractor extractor = new(pptx);
            var extractedInfo = extractor.ExtractAll();
            // Save to JSON file
            var jsonFilePath = "./test-files/pptx-data.json";
            var jsonContent = JsonSerializer.Serialize(extractedInfo, jsonSerializerOptions);
            File.WriteAllText(jsonFilePath, jsonContent);
            // Also print to console for reference
            Console.WriteLine(JsonSerializer.Serialize(extractedInfo, jsonSerializerOptions));
        }
        pptx.Dispose();
    }
}
